from tkinter import PhotoImage, Tk, messagebox
from tkinter import ttk
from functools import lru_cache
"""
On MacOS Big Sur using Python 3.9.3:
- messagebox.showwarning() shows yellow exclamationmark with small rocket icon
- showerror(),askretrycancel,askyesno,askquestion, askyesnocancel
On MacOS BigSur using Python 3.10.4 same but with a folder icon instead 
of the rocket item. Tcl/Tk version is 8.6.12 """
root = Tk()


# use a decorator with a parameter to add pre and postrocessing
# switching the iconphoto of the root/app see
# https://stackoverflow.com/questions/51530310/how-to-correct-picture-tkinter-messagebox-tkinter-on-macos
# https://realpython.com/primer-on-python-decorators/
def macos_set_icon(icon):
    @lru_cache(maxsize=10)
    def load_img(icon):
        return PhotoImage(file=f"images/{icon}.png")

    def set_icon(icon):
        img = load_img(icon)
        root.iconphoto(False, img)

    def decorator_func(original_func):
        def wrapper_func(*args, **kwargs):
            set_icon(icon)
            return original_func(*args, **kwargs)
            set_icon('app')  # restore app icon
        return wrapper_func
    return decorator_func


class MacOSMessagebox(object):

    @macos_set_icon('info')
    def showinfo(self, *args, **kwargs):
        return messagebox.showinfo(*args, **kwargs)

    @macos_set_icon('warning')
    def showwarning(self, *args, **kwargs):
        return messagebox.showwarning(*args, **kwargs)

    @macos_set_icon('error')
    def showerror(self, *args, **kwargs):
        return messagebox.showerror(*args, **kwargs)

    @macos_set_icon('question')
    def askquestion(self, *args, **kwargs):
        return messagebox.askquestion(*args, **kwargs)

    @macos_set_icon('question')
    def askyesno(self, *args, **kwargs):
        return messagebox.askyesno(*args, **kwargs)

    @macos_set_icon('question')
    def askokcancel(self, *args, **kwargs):
        return messagebox.askokcancel(*args, **kwargs)

    @macos_set_icon('question')
    def askretrycancel(self, *args, **kwargs):
        return messagebox.askretrycancel(*args, **kwargs)


macos_messagebox = MacOSMessagebox()
