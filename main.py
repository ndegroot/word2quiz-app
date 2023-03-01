""" main module of word2quiz app: GUI and CMD version"""
import re
import os
from os.path import exists
import locale
import sys
import pathlib
import io
import glob
import logging
from typing import Union, Callable
import gettext
from functools import partial
import threading
import queue
from pprint import PrettyPrinter

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from tkinter.filedialog import askopenfilename, asksaveasfile

import customtkinter
from tktooltip import ToolTip
import customtkinter as ctk
from tkhtmlview import HTMLLabel

from rich.console import Console
from rich.pretty import pprint as rich_pprint
from rich.prompt import Prompt

import word2quiz as w2q  # includes canvasrobot

if sys.platform == 'darwin':
    from macos_messagebox import root, macos_messagebox as messagebox

def get_translator() -> callable(str):
    def setup_env_windows(system_lang=True):
        """ Check environment variables used by gettext
        and setup LANG if there is none.
        """
        if _get_lang_env_var() is not None:
            return
        lang = get_language_windows(system_lang)
        if lang:
            os.environ['LANGUAGE'] = ':'.join(lang)

    def get_language_windows(system_lang=True):
        """ Get language code based on current Windows settings.
        @return: list of languages.
        """
        try:
            import ctypes
        except ImportError:
            return [locale.getdefaultlocale()[0]]
        # get all locales using Windows API
        lcid_user = ctypes.windll.kernel32.GetUserDefaultLCID()
        lcid_system = ctypes.windll.kernel32.GetSystemDefaultLCID()
        if system_lang and lcid_user != lcid_system:
            lcids = [lcid_user, lcid_system]
        else:
            lcids = [lcid_user]
        return filter(None, [locale.windows_locale.get(i) for i in lcids]) or None

    def setup_env_other(system_lang=True):
        pass

    def get_language_other(system_lang=True):
        # standard behavior for POSIX
        lang = _get_lang_env_var()
        if lang is not None:
            return lang.split(':')
        # next lines needed for MACOS when there are no LC_ or LANG env vars
        lang, encoding = locale.getdefaultlocale()
        os.environ["LANG"] = f"{lang}.{encoding}"
        return [lang]

    def _get_lang_env_var():
        for i in ('LANGUAGE', 'LC_ALL', 'LC_MESSAGES', 'LANG'):
            lang = os.environ.get(i)
            if lang:
                return lang
        return None

    if sys.platform == 'win32':
        setup_env = setup_env_windows
        get_language = get_language_windows
    else:
        setup_env = setup_env_other
        get_language = get_language_other

    en = gettext.translation('base', localedir='locales', languages=['en'])
    nl = gettext.translation('base', localedir='locales', languages=['nl'])

    # en.install()  # assume en
    # nl.install()
    # loc = locale.getlocale()
    # locale.setlocale(locale.LC_ALL, 'nl_NL')
    # print(os.environ["LC_ALL"])
    # os.environ["LC_ALL"] = "nl_NL"

    loc = locale.getlocale()

    language = get_language()
    if language and 'nl' in language[0]:
        return nl.gettext
    else:
        return en.gettext



class Word2QuizApp(ctk.CTkFrame):
    gettext: Union[Callable[[str], str], Callable[[str], str]] = get_translator()

    def __init__(self):
        super().__init__()

        self.notebook = None
        self.entry_num_questions = None
        self.tkvar_font_normalize = None
        self.cb_font_normalize = None
        self.html_lbl_docsample = None
        self.btn_convert2data = None
        self.data_dict = None
        self.txt_quiz_data = None
        self.entry_file_name = None
        self.background_image = None
        self.background_image_label = None
        self.thread_create_quizzes = None

    def get_course_combobox(self):
        db = canvasrobot.db
        fields = (db.course.course_id, db.course.sis_code)
        courses = canvasrobot.get_courses_from_database(skip_courses_without_students=True,
                                                        fields=fields)
        # create a list of dicts with sis_code as key, course_id as value
        courses_lookup = {course.sis_code: course.sis_code for course in courses}
        courses_sis_codes = courses_lookup.keys()

        self.tkvar_course_id = tk.StringVar(self.master)
        # self.tkvar_course_id.set(_('no change'))  # set the default option
        return  customtkinter.CTkComboBox(master=canvas_frame,
                                                   values=courses_sis_code,
                                                   variable=self.tkvar_course_id)

    def init_ui(self):
        """
        ---------------------------------------------
                [filename]            box text docx

                         [ Open file]
        ---------------------------------------------

            #question [dropbox]
                                     box parsed data
            [v] check box testrun

                        [ Convert ]
        ---------------------------------------------

            course_id [ input ]

                                    browserbox/link

                        [ Create quiz]
        ---------------------------------------------

        :return:
        """

        _ = self.gettext
        #  root = self.master
        self.master.title(_("Word to Canvasquiz Converter"))  # that's the tk root
        self.pack(fill="both", expand=True)

        # img_filepath = os.path.abspath(os.path.join(os.pardir, "data", "witraster.png"))
        # assert os.path.exists(img_filepath)

        # self.background_image = tk.PhotoImage(file=img_filepath)
        # self.background_image_label = tk.Label(self, image=self.background_image)
        # self.background_image_label.place(x=0, y=0)

        # self.canvas = tk.Canvas(self, width=500, height=700,
        #                        background='white',
        #                        highlightthickness=0,
        #                        borderwidth=0)
        # self.canvas.place(x=50, y=60)
        try:
            self.master.wm_iconbitmap("../data/word2quiz.ico")
        except FileNotFoundError:
            print('icon file is not available')
            pass
        file = ""
        default_text = (_("Your extracted quizdata will "
                          "appear here.\n\n please check the data"))

        # the frames
        self.notebook = ttk.Notebook(self)
        file_frame = ctk.CTkFrame(self.notebook)
        file_frame.pack(fill=tk.BOTH)
        data_frame = ctk.CTkFrame(self.notebook)
        data_frame.pack(fill=tk.BOTH)
        canvas_frame = ctk.CTkFrame(self.notebook)
        canvas_frame.pack(fill=tk.BOTH)

        self.notebook.add(file_frame, text=_('Docx File'))
        self.notebook.add(data_frame, text=_('Quiz Data'))
        self.notebook.add(canvas_frame, text=_('Canvas'))
        self.notebook.pack(expand=1, fill="both")

        # File_frame Input: button Output: filename label, box with sample
        btn_open_file = ctk.CTkButton(file_frame,
                                      text=_(' Open '),
                                      width=30,
                                      corner_radius=8,
                                      command=self.open_word_file)
        btn_open_file.pack(side="bottom", padx=5, pady=5)
        ToolTip(btn_open_file, msg=_("Open a docx file with special quiz-format. See help.") )
        # # Select word file
        lbl_file = ctk.CTkLabel(file_frame,
                                width=120,
                                height=25,
                                text=_("Select docx file"),
                                text_font=("Arial", 20)
                                )
        lbl_file.pack(side="top", pady=5)

        self.entry_file_name = ctk.CTkEntry(file_frame,
                                            placeholder_text=_("no file selected"),
                                            width=400,
                                            height=25,
                                            border_width=2,
                                            corner_radius=10)
        self.entry_file_name.pack(side="top", pady=5, padx=5)

        # # Select word file
        lbl_font_normalize = ctk.CTkLabel(file_frame,
                                          width=80,
                                          height=25,
                                          text=_("Normalize\nfontsize?"),
                                          text_font=("Arial", 12)
                                          )
        lbl_font_normalize.pack(side="left", pady=5)

        # Dictionary with options
        choices = {_('no change'), '12', '14', '16'}
        self.tkvar_font_normalize = tk.StringVar(self.master)
        self.tkvar_font_normalize.set(_('no change'))  # set the default option
        self.cb_font_normalize = tk.OptionMenu(file_frame, self.tkvar_font_normalize, *choices, )
        self.cb_font_normalize.config(width=6)

        # link function to change dropdown
        self.tkvar_font_normalize.trace('w', self.on_change_cb_normalize_fontsize)

        self.cb_font_normalize.pack(side="left", pady=5)

        self.html_lbl_docsample = HTMLLabel(file_frame,
                                            height=400,
                                            # width=100,
                                            html=_("<p><i>no content yet</i></p>"))

        self.html_lbl_docsample.pack(side="right", pady=20, padx=20)

        # data_frame Inputs: num questions, testrun Output: box quizdata
        self.btn_convert2data = ctk.CTkButton(data_frame,
                                              text=_("Convert"),
                                              width=30,
                                              command=self.create_quizdata,
                                              state=tk.DISABLED)
        self.btn_convert2data.pack(side="bottom", padx=5, pady=5)

        lbl_num_questions = ctk.CTkLabel(data_frame,
                                         width=120,
                                         height=25,
                                         text=_("How many questions\n(in  a section)"),
                                         text_font=("Arial", 12)
                                         )
        lbl_num_questions.pack(side="left", pady=5)

        self.entry_num_questions = ctk.CTkEntry(data_frame,
                                                placeholder_text="0",
                                                width=30,
                                                height=25,
                                                border_width=2)
        self.entry_num_questions.pack(side="left", pady=5, padx=5)

        #  ======================= Box to show quizdata
        self.txt_quiz_data = tk.scrolledtext.ScrolledText(data_frame,
                                                          width=350,
                                                          height=500,
                                                          bg="lightgrey",
                                                          fg="black"
                                                          )

        # self.txt_quiz_data.configure(text=default_text)
        self.txt_quiz_data.pack(side="right", padx=20, pady=10)

        lbl_course_id = ctk.CTkLabel(canvas_frame,
                                     width=120,
                                     height=25,
                                     text=_("Choose Course"),
                                     text_font=("Arial", 12)
                                     )
        lbl_course_id.pack(side="top", pady=5)

        # Dictionary with options
        # todo: get list of course ids (or names+ids) from canvas

        choices = ['34', '12723', '16']
        # self.tkvar_course_id = tk.StringVar(self.master)
        # self.tkvar_course_id.set(_('no change'))  # set the default option
        self.om_course = customtkinter.CTkComboBox(master=canvas_frame,
                                                   values = choices,
                                                   width=120,
                                                   command=self.on_select_course)
        # link function to choice of course
        #self.tkvar_course_id.trace('w', self.on_select_course)
        self.om_course.pack(side="top", pady=5)
        # progressbar
        self.pb_create_quizzes = ttk.Progressbar(canvas_frame, orient=tk.HORIZONTAL,
                               length=300, mode='determinate')
        self.pb_create_quizzes.pack(side="top", pady=10)
        self.btn_create_quizzes = ctk.CTkButton(
                canvas_frame,
                text=_(' Create Quizzes in Canvas '),
                width=30,
                corner_radius=8,
                command=self.create_quizzes,
                state=ctk.DISABLED,
                                                )
        self.btn_create_quizzes.pack(side="top", pady=10)
        # ===============================Button to access save2word method=================
        # save2canvas = Button(root, text="Save to Word File", font=('arial', 10, 'bold'),
        #                      bg="RED", fg='WHITE', command=save2canvas)
        # save2canvas.place(x=255, y=320)

        # button = InterActiveButton(self,
        #                           text="Button",
        #                           width=200,
        #                           height=50)
        # Using `anchor="w"` forces the button to expand to the right.
        # If it's removed, the button will expand in both directions
        # button.pack(padx=20, pady=20, anchor="w")

    def on_change_cb_normalize_fontsize(self, *args):
        print(self.tkvar_font_normalize.get())

    def on_select_course(self, *args):
        """ enables button on Tab Canvas that can start the creation of the quizzes
        by calling self.create_quizzes"""
        self.btn_create_quizzes.configure(state=ctk.NORMAL)

    def open_word_file(self):
        """
        Open the Word file
        save the filename in self.entry_filename
        normalize Q&A fontsize if asked"""

        _ = self.gettext

        f = askopenfilename(defaultextension=".docx",
                            filetypes=[("Word docx", "*.docx")])
        if f == "":
            f = None
        else:
            self.entry_file_name.delete(0, tk.END)
            # self.entry_file_name.config(fg="blue")
            self.entry_file_name.insert(0, f)

        normalize = self.tkvar_font_normalize.get()
        normalize = int(normalize) if normalize.isdigit() else 0
        par_list, not_recognized_list = w2q.get_document_html(filename=f,
                                                              normalized_fontsize=normalize)
        tot_html = ''
        for p_type, ans_weight, text, html in par_list:
            tot_html += f'<p style="color: green">{p_type} {html}</p>' \
                if ans_weight else f"<p>{p_type} {html}</p>"
        if not_recognized_list:
            tot_html += f"<h3>{_('Note that the next lines are not recognized')}</h3>"
            for html in not_recognized_list:
                tot_html += html
        self.html_lbl_docsample.set_html(tot_html)
        # enable next step
        # root.tk.eval('::msgcat::mclocale nl')
        result = messagebox.askquestion(title="docx",
                                        message=_('Is the first question block ok?'))
        # todo: change symbol
        if result == 'yes':
            self.notebook.select(1)
            self.btn_convert2data.configure(state=tk.NORMAL)
            return
        title = "docx"
        message = _("Check the Word doc, save it and try again")
        messagebox.showinfo(title=title, message=message)

    def create_quizdata(self):
        """
        GUI: show the quiz_data in textbox as text
        :return:
        """
        _ = self.gettext

        def to_int(var):
            return int(var) if var.isdigit() else 0

        self.data_dict, not_recognized = \
            w2q.parse_document_d2p(filename=self.entry_file_name.get(),
                                   check_num_questions=to_int(self.entry_num_questions.get()),
                                   normalize_fontsize=to_int(self.tkvar_font_normalize.get()))
        pprinter=PrettyPrinter(indent=6)
        data_text = pprinter.pformat(self.data_dict)
        self.txt_quiz_data.insert(tk.END, data_text)
        if not_recognized:
            not_recognized_text = pprinter.pformat(not_recognized)
            self.txt_quiz_data.insert(tk.END, _('\n- Not recognized lines -'))
            self.txt_quiz_data.insert(tk.END, not_recognized_text)
        else:
            self.txt_quiz_data.insert(tk.END, _('\n- All lines were recognized -'))

    def create_quizzes(self):
        """
        Create the quiz in Canvas using the quizdata
        :return: not used
        """
        _ = self.gettext

        # create a thread for the (slow) creating of quizzes in Canvas
        # stats = canvasrobot.create_quizzes_from_data(course_id=course_id,
        #                                              data=self.data_dict)
        course_id = self.om_course.get()
        kwargs = dict(course_id=course_id,
                      data=self.data_dict,
                      gui_root=self.master,
                      gui_queue=queue)

        self.thread_create_quizzes = threading.Thread(target=canvasrobot.create_quizzes_from_data,
                                                      kwargs=kwargs,
                                                      daemon=True)
        self.thread_create_quizzes.start()
        # print(f"Ready to create quiz for course {course_id}")
        # todo: call create function and show result in box
        #self.create_quiz(course_id)
        self.btn_create_quizzes.configure(state= ctk.DISABLED,
                                          text=_('Working...'))
        # idea: first check if quiz with name {} already exists, ask for overwrite


if __name__ == '__main__':
    GUI = True
    queue = queue.Queue()
    canvasrobot = w2q.CanvasRobot()

    level = logging.INFO  # global logging level, this effects canvasapi too
    logging.basicConfig(filename='word2quiz.log', encoding='utf-8', level=level)

    if GUI:
        ctk.set_appearance_mode("Light")  # Modes: system (default/Mac), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        root = root or ctk.CTk()
        root.geometry("600x800")
        root.resizable(False, False)
        # localization
        p = root.tk.eval('::msgcat::mcpackagelocale preferences')
        r = root.tk.eval('::msgcat::mcload [file join [file dirname [info script]] msgs]')
        # root.tk.eval('::msgcat::mclocale nl')

        app = Word2QuizApp()
        app.init_ui()

        # function updates the value of a progressbar
        def pb_create_updater(pb, queue, event):
            pb['value'] += 100 * queue.get()

        # connect an event used to updating the progressbar
        # the event is generated in canvasrobot-method _create_quiz_
        update_handler_create_quizzes = partial(pb_create_updater, app.pb_create_quizzes, queue)
        root.bind('<<CreateQuizzes:Progress>>', update_handler_create_quizzes)
        def show_export_ready(event):
            #todo: change buttontext and disable
            messagebox.showinfo(title="Done", message="Export to Canvas ready")

        root.bind('<<CreateQuizzes:Done>>', show_export_ready)

        root.mainloop()
        exit()


    # start of CMD version ignore when GUI is True

    TEST_COURSE_ID = 34
    _: Union[Callable[[str], str], Callable[[str], str]] = get_translator()
    console = Console(force_terminal=True)  # to trick Pycharm console into showing textattributes

    files = glob.glob('data/*.docx')
    console.print(f"aanwezig: {files}")
    filename = Prompt.ask(_("Enter filename"),
                          choices=files,
                          show_choices=True)
    with console.status(_("Working..."), spinner="dots"):
        try:
            result = w2q.word2quiz(filename,
                               course_id=TEST_COURSE_ID,
                               check_num_questions=6,
                               testrun=False)
        except FileNotFoundError as e:
            console.print(f'\n[bold red]Error:[/] {e}')
        except (IncorrectNumberofQuestions, IncorrectAnswerMarking) as e:
            console.print(f'\n[bold red]Error:[/] {e}')
        else:
            rich_pprint(result)
