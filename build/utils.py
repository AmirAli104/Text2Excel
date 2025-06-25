from tkinter import messagebox
from typing import Union

LOG_MODE = ('Disable logging','Enable logging')
APP_TITLE = 'Text2Excel'
ENCODING = 'utf-8-sig'
LOG_DEFAULT = 'log ...'
FILE_TYPES = [
    ('All Files','*.*'),
    ('Excel Files','*.xlsx;*.xlsm;*.xltx;*.xltm'),
    ('CSV Files','*.csv')
]

class WithLogging:
    with_logging = True

MENU_COLOR_ARGS = {'activebackground' : '#00c8ff', 'activeforeground' : 'black'}
EXACT_CB_GRID_ARGS = {'row' : 0, 'column' : 1, 'sticky' : 's', 'pady' : (0,10), 'padx' : (0,20)}

def show_error(err : Union[Exception,str]) -> None:
    messagebox.showerror(title=APP_TITLE, message=err)