from datetime import date, datetime, timedelta
import os
import logging
import logging.config
import keyring
import keyring.util.platform_ as keyring_platform
from rich.prompt import Prompt
import tkinter as tk
from tkinter import simpledialog

# /home/username/.config/python_keyring  # Might be different for you

print(keyring.get_keyring())
# keyring.backends.SecretService.Keyring (priority: 5)

NAMESPACE = f"{__name__}"
ENTRY = "API_KEY"
URL = "URL"


def get_api_data(app_window=False):

    key = keyring.get_password(NAMESPACE, ENTRY)
    if key is None:
        msg = "Enter your Canvas APi Key"
        key = simpledialog.askstring("Input",
                                     msg,
                                     parent=app_window) if app_window \
            else Prompt.ask(msg)
        keyring.set_password(NAMESPACE, ENTRY, key)

    url = keyring.get_password(NAMESPACE, URL)
    if url is None:
        msg = "Enter your Canvas URL (like https://tilburguniversity.instructure.com)"
        url = simpledialog.askstring("Input",
                                     msg,
                                     parent=app_window) if app_window \
            else Prompt.ask(msg)
        keyring.set_password(NAMESPACE, URL, url)

    return url, key


ENROLLMENT_TYPES = {'student': 'StudentEnrollment',
                    'teacher': 'TeacherEnrollment',
                    'ta': 'TaEnrollment',
                    'observer': 'ObserverEnrollment',
                    'designer': 'DesignerEnrollment'}

now = datetime.now()
# July first is considered the end of season
AC_YEAR = now.year - 1 if now.month < 8 else now.year
LAST_YEAR = '-{0}-{1}'.format(AC_YEAR - 1, AC_YEAR)
THIS_YEAR = '-{0}-{1}'.format(AC_YEAR, AC_YEAR + 1)
NEXT_YEAR = '-{0}-{1}'.format(AC_YEAR + 1, AC_YEAR + 2)




