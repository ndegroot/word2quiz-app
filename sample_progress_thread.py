from functools import partial
import threading
import tkinter as tk
from tkinter import ttk
import queue

import requests


def download(root, q):
    response = requests.get('https://www.python.org/ftp/python/3.10.4/Python-3.10.4.tgz', stream=True)
    dl_size = int(response.headers['Content-length'])
    
    with open('Python-3.10.4.tgz', 'wb') as fp:
        for chunk in response.iter_content(chunk_size=100_000):
            fp.write(chunk)
            q.put(len(chunk) / dl_size * 100)
            root.event_generate('<<Progress>>')
    
    root.event_generate('<<Done>>')


def updater(pb, q, event):
    pb['value'] += q.get()


q = queue.Queue()

root = tk.Tk()
root.title('Python Downloader')
root.geometry('400x100')

frame = ttk.Frame(root)
frame.pack(expand=True, fill=tk.BOTH)

progress = ttk.Progressbar(frame, length=250, orient='horizontal')
progress.pack(pady=20)

update_handler = partial(updater, progress, q)
root.bind('<<Progress>>', update_handler)

thread = threading.Thread(target=download, args=(root, q), daemon=True)

btn = ttk.Button(frame, text='Click', command=thread.start)
btn.pack()

root.bind('<<Done>>', lambda event: root.destroy())

root.mainloop()
