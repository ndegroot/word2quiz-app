""" main module of word2quiz: GUI and CMD version"""
import re
import os
# from os.path import exists
import locale
import sys
# import pathlib
import io
import glob
import logging
import tkinter
from dataclasses import dataclass
import gettext
from pprint import PrettyPrinter

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from typing import Union, Callable

import customtkinter as ctk
from tkhtmlview import HTMLLabel
# from tkinter import filedialog
from tkinter.filedialog import askopenfilename #, asksaveasfile

from lxml import etree
from rich.console import Console
from rich.pretty import pprint as rich_pprint
from rich.prompt import Prompt
import docx2python as d2p
from docx2python.iterators import iter_paragraphs
from canvas_robot import CanvasRobot, Answer
if sys.platform == 'darwin':
    from macos_messagebox import root, macos_messagebox as messagebox

GUI = True
pprint = PrettyPrinter()


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
        return list(filter(None, [locale.windows_locale.get(i) for i in lcids])) or None

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


class InterActiveButton(tk.Button):
    """
    This button expands when the user hovers over it and shrinks when
    the cursor leaves the button.

    If you want the button to expand in both directions just use:
        button = InterActiveButton(root, text="Button", width=200, height=50)
        button.pack()
    If you want the button to only expand to the right use:
        button = InterActiveButton(root, text="Button", width=200, height=50)
        button.pack(anchor="w")

    This button should work with all geometry managers.
    """

    def __init__(self, master, max_expansion: int = 12, bg="dark blue",
                 fg="#dad122", **kwargs):
        # Save some variables for later:
        self.max_expansion = max_expansion
        self.bg = bg
        self.fg = fg

        # To use the button's width in pixels:
        # From here: https://stackoverflow.com/a/46286221/11106801
        self.pixel = tk.PhotoImage(width=1, height=1)

        # The default button arguments:
        button_args = dict(cursor="hand2", bd=0, font=("arial", 18, "bold"),
                           height=50, compound="c", activebackground=bg,
                           image=self.pixel, activeforeground=fg)
        button_args.update(kwargs)
        super().__init__(master, bg=bg, fg=fg, **button_args)

        # Bind to the cursor entering and exiting the button:
        super().bind("<Enter>", self.on_hover)
        super().bind("<Leave>", self.on_leave)

        # Save some variables for later:
        self.base_width = button_args.pop("width", 200)
        self.width = self.base_width
        # `self.mode` can be "increasing"/"decreasing"/None only
        # It stops a bug where if the user wuickly hovers over the button
        # it doesn't go back to normal
        self.mode = None

    def increase_width(self) -> None:
        if self.width <= self.base_width + self.max_expansion:
            if self.mode == "increasing":
                self.width += 1
                super().config(width=self.width)
                super().after(5, self.increase_width)

    def decrease_width(self) -> None:
        if self.width > self.base_width:
            if self.mode == "decreasing":
                self.width -= 1
                super().config(width=self.width)
                super().after(5, self.decrease_width)

    def on_hover(self, event: tk.Event = None) -> None:
        # Improvement: use integers instead of "increasing" and "decreasing"
        self.mode = "increasing"
        # Swap the `bg` and the `fg` of the button
        super().config(bg=self.fg, fg=self.bg)
        super().after(5, self.increase_width)

    def on_leave(self, event: tk.Event = None) -> None:
        # Improvement: use integers instead of "increasing" and "decreasing"
        self.mode = "decreasing"
        # Reset the `fg` and `bg` of the button
        super().config(bg=self.bg, fg=self.fg)
        super().after(5, self.decrease_width)


class Word2Quiz(ctk.CTkFrame):
    gettext: Union[Callable[[str], str], Callable[[str], str]] = get_translator()

    def __init__(self):
        super().__init__()

        self.quiz_data = None
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
        self.course_id = None

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
            self.master.iconbitmap("../images/word2quiz.ico")
        except (FileNotFoundError, tkinter.TclError):
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

        self.notebook.add(file_frame, text='Docx File')
        self.notebook.add(data_frame, text='Quiz Data')
        self.notebook.add(canvas_frame, text='Canvas')
        self.notebook.pack(expand=1, fill="both")

        # File_frame Input: button Output: filename label, box with sample
        btn_open_file = ctk.CTkButton(file_frame,
                                      text=_(' Open '),
                                      width=30,
                                      corner_radius=8,
                                      command=self.open_word_file)
        btn_open_file.pack(side="bottom", padx=5, pady=5)
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
                                            html="""
            <p><i>no content yet</i></p>
            """)

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
            self.entry_file_name.config(fg="blue")
            self.entry_file_name.insert(0, f)

        normalize = self.tkvar_font_normalize.get()
        normalize = int(normalize) if normalize.isdigit() else 0
        par_list, not_recognized_list = get_document_html(filename=f,
                                                          normalized_fontsize=normalize)
        html = ''
        tot_html = ''
        for p_type, ans_weight, text, html in par_list:
            tot_html += f'<p style="color: green">{p_type} {html}</p>' \
                if ans_weight else f"<p>{p_type} {html}</p>"
        if not_recognized_list:
            html += "<hr><p>Not recognized:</p>"
            for d1, d2, html in not_recognized_list:
                tot_html += html
        self.html_lbl_docsample.set_html(tot_html)
        # enable next step
        # root.tk.eval('::msgcat::mclocale nl')
        result = messagebox.askquestion(title="docx",
                                           message=_('Is the first Q&A ok?'))
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
            parse_document_d2p(filename=self.entry_file_name.get(),
                               check_num_questions=to_int(self.entry_num_questions.get()),
                               normalize_fontsize=to_int(self.tkvar_font_normalize.get()))
        data_text = pprint.pformat(self.data_dict)
        self.txt_quiz_data.insert(tk.END, data_text)
        if not_recognized:
            not_recognized_text = pprint.pformat(not_recognized)
            self.txt_quiz_data.insert(tk.END, _('\n- Not recognized lines -'))
            self.txt_quiz_data.insert(tk.END, not_recognized_text)
        else:
            self.txt_quiz_data.insert(tk.END, _('\n- All lines were recognized -'))

    def create_quiz(self, course_id):
        """
        Create the quiz in Canvas using the quizdata
        :return: not used"
        """
        canvasrobot = CanvasRobot()
        result, stats = canvasrobot.create_quizzes_from_data(course_id=course_id,
                                                             data=self.quiz_data)


# define Python user-defined exceptions
class Error(Exception):
    """ Base class for other exceptions"""
    pass


class IncorrectNumberofQuestions(Error):
    """ the quiz contains unexpected number of questions"""
    pass


class IncorrectAnswerMarking(Error):
    """ the answers of a particular question should have
        only one good (!) marking (or total of 100 points) """
    pass


# from docx import Document  # package - python-docx !
# import docx2python as d2p
# from xdocmodel import iter_paragraphs

def normalize_size(text: str, size: int):
    parser = etree.XMLParser()
    try:  # can be html or not
        tree = etree.parse(io.StringIO(text), parser)
        # text could contain style attribute
        ele = tree.xpath('//span[starts-with(@style,"font-size:")]')
        if ele is not None and len(ele):
            ele[0].attrib['style'] = f"font-size:{size}pt"
            return etree.tostring(ele[0], encoding='unicode')
    except etree.XMLSyntaxError as e:
        # assume simple html string no surrounding tags
        return f'<span style="font-size:{size}pt">{text}</span>'


FULL_SCORE = 100
TITLE_SIZE = 24
QA_SIZE = 12

# the patterns
title_pattern = \
    re.compile(r"^<font size=\"(?P<fontsize>\d+)\"><u>(?P<text>.*)</u></font>")
title_style_pattern = \
    re.compile(r"^<span style=\"font-size:(?P<fontsize>[\dpt]+)\"><u>(?P<text>.*)</u>")

quiz_name_pattern = \
    re.compile(r"^<font size=\"(?P<fontsize>\d+[^\"]+)\"><b>(?P<text>.*)\s*</b></font>")
quiz_name_style_pattern = re.compile(
    r"^<span style=\"font-size:(?P<fontsize>[\dpt]+)"
    r"(;text-transform:uppercase)?\"><b>(?P<text>.*)\s*</b></span>")
# special match Sam
page_ref_style_pattern = \
    re.compile(r'(\(pp\.\s+[\d-]+)')

q_pattern_fontsize = \
    re.compile(r'^(?P<id>\d+)[).]\s+'
               r'(?P<prefix><font size="(?P<fontsize>\d+)">)(?P<text>.*</font>)')
q_pattern = \
    re.compile(r"^(?P<id>\d+)[).]\s+(?P<text>.*)")

# '!' before the text of answer marks it as the right answer
# idea: use [\d+]  for partially correct answer the sum must be FULL_SCORE
a_ok_pattern_fontsize = re.compile(
    r'^(?P<id>[a-d])\)\s+(?P<prefix><font size="(?P<fontsize>\d+)">.*)'
    r'(?P<fullscore>!)(?P<text>.*</font>)')
a_ok_pattern = \
    re.compile(r"^(?P<id>[a-d])\)\s+(?P<prefix>.*)(?P<fullscore>!)(?P<text>.*)")
# match a-d then ')' then skip whitespace and all chars up to '!' after answer skip </font>

a_wrong_pattern_fontsize = \
    re.compile(r'^(?P<id>[a-d])\)\s+'
               r'(?P<prefix><font size="(?P<fontsize>\d+)">)(?P<text>.*</font>)')
a_wrong_pattern = \
    re.compile(r"^(?P<id>[a-d])\)\s+(?P<text>.*)")


@dataclass()
class Rule:
    name: str
    pattern: re.Pattern
    type: str
    normalized_size: int = QA_SIZE


rules = [
    Rule(name='title', pattern=title_pattern, type='Title'),
    Rule(name='title_style', pattern=title_style_pattern, type='Title'),
    Rule(name='quiz_name', pattern=quiz_name_pattern, type='Quizname'),
    Rule(name='quiz_name_style', pattern=quiz_name_style_pattern, type='Quizname'),
    Rule(name='page_ref_style', pattern=page_ref_style_pattern, type='PageRefStyle'),
    Rule(name='question_fontsize', pattern=q_pattern_fontsize, type='Question'),
    Rule(name='question', pattern=q_pattern, type='Question'),
    Rule(name='ok_answer_fontsize', pattern=a_ok_pattern_fontsize, type='Answer'),
    Rule(name='ok_answer', pattern=a_ok_pattern, type='Answer'),
    Rule(name='wrong_answer_fontsize', pattern=a_wrong_pattern_fontsize, type='Answer'),
    Rule(name='wrong_answer', pattern=a_wrong_pattern, type='Answer'),
]


def get_document_html(filename: str, normalized_fontsize: int = 0):
    """
        :param normalized_fontsize: 0 is no normalization
        :param  filename: filename of the Word docx to parse
        :returns the first X paragraphs as HTML
    """
    #  from docx produce a text with minimal HTML formatting tags b,i, font size
    #  1) questiontitle
    #    a) wrong answer
    #    b) !right answer
    doc = d2p.docx2python(filename, html=True)
    # print(doc.body)
    section_nr = 0  # state machine

    #  the Word text contains one or more sections
    #  quiz_name (multiple)
    #    questions (5) starting with number 1
    #       answers (4)
    # we save the question list into the result list when we detect new question 1
    last_p_type = None
    par_list = []
    not_recognized = []
    # stop after the first question
    for par in d2p.iterators.iter_paragraphs(doc.body):
        par = par.strip()
        if not par:
            continue
        question_nr, weight, text, p_type = parse(par, normalized_fontsize)
        print(f"{par} = {p_type} {weight}")
        if p_type == 'Not recognized':
            not_recognized.append(par)
            continue

        if p_type == 'Quizname':
            quiz_name = text
            par_list.append((p_type, None, text, par))
        if last_p_type == 'Answer' and p_type in ('Question', 'Quizname'):  # last answer
            break
        if p_type == 'Answer':
            # answers.append(Answer(answer_html=text, answer_weight=weight))
            par_list.append((p_type, weight, text, par))
        if p_type == "Question":
            par_list.append((p_type, None, text, par))

        last_p_type = p_type

    return par_list, not_recognized


def parse(text: str, normalized_fontsize: int = 0):
    """
    :param text : text to parse
    :param normalized_fontsize: if non-zero change fontsizes in q & a
    :return determine the type and parsed values of a string by matching a ruleset and returning a
    tuple:
    - question number/answer: int/char,
    - score :int (if answer),
    - text: str,
    - type: str. One of ('Question','Answer', 'Title, 'Pageref', 'Quizname') or 'Not recognized'
    """

    def is_qa(rule):
        return rule.type in ('Question', 'Answer')

    for rule in rules:
        match = rule.pattern.match(text)
        if match:
            if rule.name in ('page_ref_style',):
                # just skip it
                continue
            id_str = match.group('id') if 'id' in match.groupdict() else ''
            id_norm = int(id_str) if id_str.isdigit() else id_str
            score = FULL_SCORE if 'fullscore' in match.groupdict() else 0
            prefix = match.group('prefix') if 'prefix' in match.groupdict() else ''
            text = prefix + match.group('text').strip()
            text = normalize_size(text, normalized_fontsize) if (normalized_fontsize > 0
                                                                 and is_qa(rule)) else text
            return id_norm, score, text, rule.type

    return None, 0, "", 'Not recognized'


def parse_document_d2p(filename: str, check_num_questions: int, normalize_fontsize=0):
    """ Pddarse the docx file, while checking number of questions and optionally normalize fontsize
        of the questions and answers (the fields in Canvas with HTML formatted texts)
        :param  filename: filename of the Word docx to parse
        :param check_num_questions: number of questions in a section
        :param normalize_fontsize: if > 0 change fontsizes Q&A
        :returns Tuple of [
          quiz_data: a List of Tuples[
            - quiz_names: str
            - questions: List[
                - question_name: str,
                - List[ Answers: List of Tuple[
                    name:str,
                    weight:int]]],
          not_recognized: List of not recognized lines]"""
    #  from docx produce a text with minimal HTML formatting tags b,i, font size
    #  1) question title
    #    a) wrong answer
    #    b) !right answer
    doc = d2p.docx2python(filename, html=True)
    # print(doc.body)
    section_nr = 0  # state machine
    last_p_type = None
    quiz_name = None
    last_quiz_name = None
    question_text = ''
    question_list = []
    not_recognized = []
    result = []
    answers = []

    #  the Word text contains one or more sections
    #  quiz_name (multiple)
    #    questions (5) starting with number 1
    #       answers (4)
    # we save the question list into the result list when we detect new question 1

    for par in d2p.iterators.iter_paragraphs(doc.body):
        par = par.strip()
        if not par:
            continue
        question_nr, weight, text, p_type = parse(par, normalize_fontsize)
        logging.debug(f"{par} = {p_type} {weight}")
        if p_type == 'Not recognized':
            not_recognized.append(par)
            continue

        if p_type == 'Quizname':
            last_quiz_name = quiz_name  # we need it, when saving question_list
            quiz_name = text
        if last_p_type == 'Answer' and p_type in ('Question', 'Quizname'):  # last answer
            question_list.append((question_text, answers))
            answers = []
        if p_type == 'Answer':
            answers.append(Answer(answer_html=text, answer_weight=weight))
        if p_type == "Question":
            question_text = text
            if question_nr == 1:
                logging.debug("New quiz is being parsed")
                if section_nr > 0:  # after first section add the quiz+questions
                    result.append((last_quiz_name, question_list))
                question_list = []
                section_nr += 1

        last_p_type = p_type
    # handle last question
    question_list.append((question_text, answers))
    # handle last section
    result.append((quiz_name, question_list))
    for question_list in result:
        nr_questions = len(question_list[1])
        if check_num_questions:
            if nr_questions != check_num_questions:
                raise IncorrectNumberofQuestions(f"Questionlist {question_list[0]} has "
                                                 f"{nr_questions} questions "
                                                 f"this should be {check_num_questions} questions")
        for questions in question_list[1]:
            assert len(questions[1]) == 4, f"{questions[0]} only {len(questions[1])} of 4 answers"
            tot_weight = 0
            for ans in questions[1]:
                tot_weight += ans.answer_weight
            if tot_weight != FULL_SCORE:
                raise IncorrectAnswerMarking(f"Check right/wrong marking and weights in "
                                             f"Q '{questions[0]}'\n Ans {questions[1]}")

    logging.debug('--- not recognized:' if not_recognized else
                  '--- all lines were recognized ---')
    for line in not_recognized:
        logging.debug(line)

    return result, not_recognized


def word2quiz(filename: str,
              course_id: int,
              check_num_questions,
              normalize_fontsize=False,
              testrun=False):
    """
    """
    os.chdir('../data')
    logging.debug(f"We are in folder {os.getcwd()}")
    quiz_data, not_recognized_lines = parse_document_d2p(filename=filename,
                                                         check_num_questions=check_num_questions,
                                                         normalize_fontsize=normalize_fontsize)
    if testrun:
        return quiz_data
    canvasrobot = CanvasRobot()
    result, stats = canvasrobot.create_quizzes_from_data(course_id=course_id,
                                                         data=quiz_data)

    return stats


if __name__ == '__main__':

    level = logging.INFO  # global logging level, this effects canvasapi too
    logging.basicConfig(filename=f'word2quiz.log', encoding='utf-8', level=level)

    if GUI:
        ctk.set_appearance_mode("Light")  # Modes: system (default/Mac), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
        # works on mac root = root or ctk.CTk()
        root = ctk.CTk()
        root.geometry("550x700")
        root.resizable(False, False)
        p = root.tk.eval('::msgcat::mcpackagelocale preferences')
        r = root.tk.eval('::msgcat::mcload [file join [file dirname [info script]] msgs]')

        #root.tk.eval('::msgcat::mclocale nl')
        app = Word2Quiz()
        app.init_ui()

        root.mainloop()
        exit()

    # start of CMD version
    _: Union[Callable[[str], str], Callable[[str], str]] = get_translator()
    console = Console(force_terminal=True)  # to trick Pycharm console into showing textattributes

    files = glob.glob('../data/*.docx')
    filename = Prompt.ask(_("Enter filename"),
                          choices=files,
                          show_choices=True)
    with console.status(_("Working..."), spinner="dots"):
        try:
            result = word2quiz(filename,
                               course_id=34,
                               check_num_questions=6,
                               testrun=False)
        except FileNotFoundError as e:
            console.print(f'\n[bold red]Error:[/] {e}')
        except (IncorrectNumberofQuestions, IncorrectAnswerMarking) as e:
            console.print(f'\n[bold red]Error:[/] {e}')
        else:
            rich_pprint(result)
