from pathlib import Path
from tkmsgcat import load, locale, get
import customtkinter as ctk
from customtkinter import CTkInputDialog, CTkFrame, CTkToplevel, CTKMessageBox
import tkinter as tk
from tkinter.messagebox import askyesno, askquestion
import time

tk_part = False

if tk_part:
    root = tk.Tk()
    root.geometry("600x800")
    root.resizable(False, False)

    r = root.tk.eval('::msgcat::mcpackagelocale preferences')
    print(r)

    # r = root.tk.eval('::msgcat::mcload [file join [file dirname [info script]] msgs]')
    msgsdir = Path(__file__).parent / "msgs"
    load(msgsdir)

    # locale("nl")  # not needed if the locale is set in envs LANG or others

    tr = get("Hello")  # Hallo

    print(tr)
    print(get("Goodbye"))
    print(get("Yes"))

app = ctk.CTk()
app.geometry("400x300")


class CTkInputDialogCustomButtonText(CTkInputDialog):

    def __init__(self, master=None, title="CTkDialog", text="CTkDialog", fg_color="default_theme",
                 hover_color="default_theme", border_color="default_theme",
                 ok_text='Ok', cancel_text='Annuleer'):
        super().__init__(master, title, text, fg_color, hover_color, border_color)
        self.top.after(11, lambda: self.change_buttons(ok_text, cancel_text))
        # needs after() otherwise the button to be changed are not there yet
        # see CTKInputdialog.__init__ which uses after() too.
        # 'lambda' is needed because after() callback command has parameters see
        # https://www.pythontutorial.net/tkinter/tkinter-after/

    def change_buttons(self, ok_text, cancel_text):
        self.ok_button.configure(text=ok_text)
        self.cancel_button.configure(text=cancel_text)


def open_input_dialog():
    # dialog = CTkInputDialogCustomButtonText(master=None, text="Type in a number:", title="Test")
    dialog = CTkInputDialog(master=None,
                            text="Type in a number:",
                            title="Test",
                            ok_text='Ok',
                            cancel_text='Annuleer')
    print("Number:", dialog.get_input())


def open_message_box():
    # dialog = CTkInputDialogCustomButtonText(master=None, text="Type in a number:", title="Test")
    dialog = CTkMessageBox(master=None,
                           message="A message",
                           title="a Title",
                           ok_text='Ok',
                           cancel_text='Annuleer')


def open_std_message_boxes():
    # dialog = CTkInputDialogCustomButtonText(master=None, text="Type in a number:", title="Test")
    dialog = tk.messagebox.showinfo(message="Information", icon="error")


button = ctk.CTkButton(app, text="Open Dialog", command=open_message_box)
button.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

app.mainloop()
