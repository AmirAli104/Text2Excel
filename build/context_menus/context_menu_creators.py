
from tkinter.ttk import Entry
from tkinter import Menu
from utils import *
from context_menus.context_menu_commands import *

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

    @staticmethod
    def create_patterns_menu() -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        # The numbers below the following lines are indices of context menu commands
        # which are used in context_menu_displayers.py module given to entryconfig method
        menu.add_command(label='Add Pattern', command=CommandsObjects.patterns_menu_commands.add_pattern,
                         accelerator='Ctrl+Shift+A', underline=0) # 0
        menu.add_command(label='Insert Pattern',command=CommandsObjects.patterns_menu_commands.insert_pattern,
                         accelerator='Ctrl+I', underline=2) # 1
        menu.add_separator() # 2
        menu.add_command(label='Move Up', command = CommandsObjects.patterns_menu_commands.move_selected, 
                         accelerator='K', underline=5) # 3
        menu.add_command(label='Move Down', command = lambda : CommandsObjects.patterns_menu_commands.move_selected(up = False), 
                         accelerator='J', underline=5) # 4
        menu.add_separator() # 5
        menu.add_command(label='Edit selected', command=CommandsObjects.patterns_menu_commands.edit_selected,
                         accelerator='F2', underline=0) # 6
        menu.add_command(label='Delete selected', command=CommandsObjects.patterns_menu_commands.delete_selected, accelerator='Delete') # 7
        menu.add_command(label='Copy selected', command=CommandsObjects.patterns_menu_commands.copy_pattern,
                         accelerator='Ctrl+C', underline=0) # 8
        menu.add_command(label='Delete All', command=CommandsObjects.patterns_menu_commands.delete_all, accelerator='Ctrl+Shift+D') # 9
        menu.add_command(label='Copy All', command=lambda : CommandsObjects.patterns_menu_commands.copy_pattern(all=True),
                         accelerator='Ctrl+Shift+C',underline=1) # 10
        menu.add_separator() # 11
        menu.add_command(label='Import from file', command=CommandsObjects.patterns_menu_commands.import_from_file,
                         accelerator='Ctrl+Shift+I', underline=0) # 12
        menu.add_command(label='Export to file', command=CommandsObjects.patterns_menu_commands.export_to_file,
                         accelerator='Ctrl+E', underline=1) # 13
        return menu

    @staticmethod
    def create_log_menu() -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        menu.add_command(label='Copy log',command=CommandsObjects.log_menu_commands.copy_log,accelerator='Ctrl+C', underline=0) # 0
        menu.add_command(label='Clear log',command=CommandsObjects.log_menu_commands.clear_log,accelerator='Ctrl+D', underline=1) # 1
        menu.add_separator() # 2
        menu.add_command(label=LOG_MODE[0],command=CommandsObjects.log_menu_commands.toggle_log, underline=0) # 3
        return menu

    @staticmethod
    def create_entry_menu(widget : Entry, excel_var, is_file_entry : bool=True,is_output_file_entry : bool=True) -> Menu:
        menu = Menu(tearoff=False,**MENU_COLOR_ARGS)
        menu.add_command(label='Select All', command=lambda : widget.select_range(0,'end'), accelerator='Ctrl+A', underline=7) # 0
        menu.add_command(label='Copy', command=lambda : widget.event_generate('<<Copy>>'), accelerator='Ctrl+C', underline=0) # 1
        menu.add_command(label='Paste', command=lambda : widget.event_generate('<<Paste>>'), accelerator='Ctrl+V', underline=0) # 2
        menu.add_command(label='Cut', command=lambda : widget.event_generate('<<Cut>>'), accelerator='Ctrl+X', underline=2) # 3
        menu.add_separator() # 4
        menu.add_command(label='Clear', command=lambda : widget.delete(0,'end'), accelerator='Ctrl+Shift+C', underline=1) # 5
        if is_file_entry:
            menu.add_command(label='Browse',command=lambda : browse_files(widget, not is_output_file_entry), accelerator='Ctrl+B', underline=0) # 6
            if is_output_file_entry:
                menu.add_separator() # 7
                menu.add_radiobutton(label='Excel',variable=excel_var,value=True,
                                     command=CommandsObjects.csv_excel_switch_functions.show_only_excel_required_widgets, underline=0) # 8

                menu.add_radiobutton(label='CSV',variable=excel_var,value=False,
                                     command=CommandsObjects.csv_excel_switch_functions.hide_only_excel_required_widgets, underline=1) # 9
        return menu

