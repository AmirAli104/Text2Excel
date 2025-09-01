from tkinter import filedialog
from tkinter.simpledialog import askstring

from utils import *
from tkinter import Event, TclError

class PatternsMenuCommands:
    def __init__(self, patterns_list, window):
        self.patterns_list = patterns_list
        self.window = window

    @staticmethod
    def get_pattern() -> str:
        return askstring(title=APP_TITLE,prompt='Enter the pattern:')

    def add_pattern(self, event : Event=None) -> None:
        new_pattern = self.get_pattern()
        self.patterns_list.insert('end',new_pattern)

    def insert_pattern(self, event : Event=None) -> None:
        selected = self.patterns_list.curselection()
        if len(selected) == 1:
            new_pattern = self.get_pattern()
            self.patterns_list.insert(selected[0],new_pattern)
            
    def swap_up(self, selected_index : int):
        next_index = selected_index - 1

        next_item = self.patterns_list.get(selected_index)
        self.patterns_list.insert(next_index, next_item)
        self.patterns_list.delete(selected_index + 1)

    def move_selected(self, event : Event = None , up = True):
        selected_index = self.patterns_list.curselection()[0]

        if up:
            if selected_index == 0:
                return

            self.swap_up(selected_index)
            self.patterns_list.selection_set(selected_index - 1)

        else:
            self.swap_up(selected_index + 1)
        
    def edit_selected(self, event : Event=None) -> None:
        index = self.patterns_list.curselection()
        if len(index) == 1:
            value = self.patterns_list.get(index)
            new_value = askstring(title=APP_TITLE, prompt='Enter the pattern: ', initialvalue=value)
            if new_value:
                self.patterns_list.delete(index)
                self.patterns_list.insert(index, new_value)

    def delete_selected(self, event : Event=None) -> None:
        selected = self.patterns_list.curselection()
        if selected:
            self.patterns_list.delete(selected[0],selected[-1])
    
    def delete_all(self, event : Event=None) -> None:
        self.patterns_list.delete(0,'end')

    def copy_pattern(self, event : Event=None, all : bool=False) -> None:
        self.window.clipboard_clear()
        if all:
            self.window.clipboard_append('\n'.join(self.patterns_list.get(0,'end')))
        else:
            selected = self.patterns_list.curselection()
            if selected:
                self.window.clipboard_append('\n'.join(self.patterns_list.get(selected[0],selected[-1])))

    def import_from_file(self, event : Event=None) -> None:
        try:
            file_path = filedialog.askopenfilename(title='Import')
            if file_path:
                with open(file_path, encoding=ENCODING) as f:
                    for i in f.read().strip().splitlines():
                        self.patterns_list.insert('end', i)
        except UnicodeDecodeError:
            show_error('The patterns file cannot be a binary file')

    def export_to_file(self, event : Event=None) -> None:
        try:
            file_path = filedialog.asksaveasfilename(title='Export')
            if file_path:
                with open(file_path, 'w',encoding=ENCODING) as f:
                    for i in self.patterns_list.get(0,'end'):
                        f.write(i + '\n')
        except Exception as err:
            show_error(err)

class LogMenuCommands:
    def __init__(self, log_menu, window, log_text):
        self.log_menu = log_menu
        self.log_text = log_text
        self.window = window

    def toggle_log(self) -> None:
        WithLogging.with_logging = not WithLogging.with_logging
        
        mode = self.log_menu.entrycget(3,'label')

        if mode == LOG_MODE[0]:
            self.log_menu.entryconfig(3,label = LOG_MODE[1])
        else:
            self.log_menu.entryconfig(3,label = LOG_MODE[0])

    def copy_log(self, event : Event=None) -> None:
        self.window.clipboard_clear()
        try:
            data = self.log_text.selection_get()
        except TclError:
            data = self.log_text.get('1.0','end')
        self.window.clipboard_append(data)

    def clear_log(self, event : Event=None) -> None:
        self.log_text.config(state='normal')
        self.log_text.delete('1.0','end')
        self.log_text.insert('end',LOG_DEFAULT)
        self.log_text.config(state='disabled')

class CSVExcelSwitchFunctions:

    def __init__(self, exact_var, exact_cb, exact_cb_substitute_lbl, sheet_name_lbl, sheet_name_entry, col_var, excel_var):
        self.exact_var = exact_var
        self.exact_cb = exact_cb
        self.exact_cb_substitute_lbl = exact_cb_substitute_lbl
        self.sheet_name_lbl = sheet_name_lbl
        self.sheet_name_entry = sheet_name_entry
        self.col_var = col_var
        self.excel_var = excel_var

        self.exact_var_value = None

    def hide_exact_order_cb(self) -> None:
            self.exact_cb.grid_remove()
            self.exact_var_value = self.exact_var.get()
            self.exact_var.set(False)
            self.exact_cb_substitute_lbl.grid(**EXACT_CB_GRID_ARGS)

    def show_exact_order_cb(self) -> None:
        if self.excel_var.get():
            self.exact_cb_substitute_lbl.grid_remove()
            self.exact_var.set(self.exact_var_value)
            self.exact_cb.grid()

    def hide_only_excel_required_widgets(self) -> None: # sheet_name_entry, sheet_name_lbl, exact_cb
        if self.exact_cb.winfo_ismapped():
            self.hide_exact_order_cb()
        
        self.sheet_name_lbl.grid_remove()
        self.sheet_name_entry.grid_remove()

    def show_only_excel_required_widgets(self) -> None:
        if self.col_var.get():
            self.show_exact_order_cb()

        self.sheet_name_lbl.grid()
        self.sheet_name_entry.grid()
