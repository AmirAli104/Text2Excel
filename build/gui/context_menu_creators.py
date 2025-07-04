from tkinter.ttk import Entry
from tkinter import Menu
from utils import *
from gui.menu_commands import *

class CommandsObjects:
    log_menu_commands = None
    patterns_menu_commands = None
    csv_excel_switch_functions = None

def create_commands_objects(log_menu,log_text , window, patterns_list, exact_var,
                            exact_cb, exact_cb_substitute_lbl, sheet_name_lbl,
                            sheet_name_entry, col_var, excel_var):
    
    CommandsObjects.csv_excel_switch_functions = CSVExcelSwitchFunctions(exact_var, exact_cb, exact_cb_substitute_lbl,
                                                        sheet_name_lbl, sheet_name_entry, col_var, excel_var)
    CommandsObjects.log_menu_commands = LogMenuCommands(log_menu, window, log_text)
    CommandsObjects.patterns_menu_commands = PatternsMenuCommands(patterns_list, window)

def browse_files(widget : Entry, is_input_file_entry : bool) -> None:
    TITLE = 'Browse'

    if is_input_file_entry:
        file_path = filedialog.askopenfilename(title=TITLE)
    else:
        file_path = filedialog.askopenfilename(title=TITLE,filetypes=FILE_TYPES)

    if file_path:
        widget.delete(0,'end')
        widget.insert('end',file_path)

class MenuCreators:
    def create_patterns_menu() -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        menu.add_command(label='Add Pattern', command=CommandsObjects.patterns_menu_commands.add_pattern,accelerator='Ctrl+Shift+A')
        menu.add_command(label='Insert Pattern',command=CommandsObjects.patterns_menu_commands.insert_pattern,accelerator='Ctrl+I')
        menu.add_separator()
        menu.add_command(label='Edit selected', command=CommandsObjects.patterns_menu_commands.edit_selected,accelerator='F2')
        menu.add_command(label='Delete selected', command=CommandsObjects.patterns_menu_commands.delete_selected,accelerator='Delete')
        menu.add_command(label='Copy selected', command=CommandsObjects.patterns_menu_commands.copy_pattern,accelerator='Ctrl+C')
        menu.add_command(label='Delete All', command=CommandsObjects.patterns_menu_commands.delete_all,accelerator='Ctrl+Shift+D')
        menu.add_command(label='Copy All', command=lambda : CommandsObjects.patterns_menu_commands.copy_pattern(all=True),accelerator='Ctrl+Shift+C')
        menu.add_separator()
        menu.add_command(label='Import from file', command=CommandsObjects.patterns_menu_commands.import_from_file,accelerator='Ctrl+Shift+I')
        menu.add_command(label='Export to file', command=CommandsObjects.patterns_menu_commands.export_to_file,accelerator='Ctrl+E')
        return menu

    def create_log_menu() -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        menu.add_command(label='Copy log',command=CommandsObjects.log_menu_commands.copy_log,accelerator='Ctrl+C')
        menu.add_command(label='Clear log',command=CommandsObjects.log_menu_commands.clear_log,accelerator='Ctrl+D')
        menu.add_separator()
        menu.add_command(label=LOG_MODE[0],command=CommandsObjects.log_menu_commands.toggle_log)
        return menu

    def create_entry_menu(widget : Entry, excel_var, is_file_entry : bool=True,is_output_file_entry : bool=True) -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        menu.add_command(label='Select All', accelerator='Ctrl+A',command=lambda : widget.select_range(0,'end'))
        menu.add_command(label='Copy', accelerator='Ctrl+C',command=lambda : widget.event_generate('<<Copy>>'))
        menu.add_command(label='Paste', accelerator='Ctrl+V',command=lambda : widget.event_generate('<<Paste>>'))
        menu.add_command(label='Cut', accelerator='Ctrl+X',command=lambda : widget.event_generate('<<Cut>>'))
        menu.add_separator()
        menu.add_command(label='Clear',accelerator='Ctrl+Shift+C',command=lambda : widget.delete(0,'end'))
        if is_file_entry:
            menu.add_command(label='Browse',command=lambda : browse_files(widget, not is_output_file_entry),accelerator='Ctrl+B')
            if is_output_file_entry:
                menu.add_separator()
                menu.add_radiobutton(label='Excel',variable=excel_var,value=True,command=CommandsObjects.csv_excel_switch_functions.show_only_excel_required_widgets)
                menu.add_radiobutton(label='CSV',variable=excel_var,value=False,command=CommandsObjects.csv_excel_switch_functions.hide_only_excel_required_widgets)
        return menu