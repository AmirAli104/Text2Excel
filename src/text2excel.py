"""
Text2Excel is a GUI desktop application that can extract data from a text file
and put them in an Excel or CSV file using regular expression (regex) patterns.

version 2.7.1

https://github.com/AmirAli104/Text2Excel

License: MIT License
"""

__version__ = '2.7.1'

import tkinter as tk
from tkinter import ttk
import os

from utils import *
from context_menus.context_menu_commands import *
from context_menus.context_menu_displayers import *
from context_menus.context_menu_creators import *
from extractors import DataExtractor

# ----- the root window -----

window = tk.Tk()
window.title(APP_TITLE)
window.resizable(False,False)

main_frm = tk.Frame()
main_frm.grid(row=0,column=0, rowspan=3,sticky='nsew')

frm = tk.Frame(main_frm,borderwidth=10)
frm.pack(fill='both',expand=1)

# ----- App icon -----

icon_path = os.path.join(os.path.dirname(__file__), f'resources{os.sep}icon.png')
icon = tk.PhotoImage(file=icon_path)
window.iconphoto(True,icon)

# ----- defining 'input file', 'output file' and 'sheet name' labels and entries -----

input_file_lbl = tk.Label(frm,text='Input file:')
output_file_lbl = tk.Label(frm,text='Output file:')
sheet_name_lbl = tk.Label(frm,text='Sheet name:')

input_file_entry = ttk.Entry(frm,width=30)
output_file_entry = ttk.Entry(frm,width=30)
sheet_name_entry = ttk.Entry(frm,width=15)

input_file_lbl.grid(row=2,column=0,sticky='w')
output_file_lbl.grid(row=2, column=3,sticky='w')
sheet_name_lbl.grid(row=4, column=3, sticky='w')

input_file_entry.grid(row=3,column=0,sticky='w')
output_file_entry.grid(row=3,column=3,sticky='w')
sheet_name_entry.grid(row=5, column=3, sticky='w')

# ----- 'log text' widget and its scrollbars -----

log_frm = tk.Frame(main_frm)
yscroll_log = tk.Scrollbar(log_frm)
xscroll_log = tk.Scrollbar(log_frm, orient='horizontal')
log_text = tk.Text(log_frm,width=23, height=10, font = 'TkTextFont', wrap = 'none',
                   yscrollcommand=yscroll_log.set, xscrollcommand=xscroll_log.set,takefocus=True,
                   highlightcolor='black',highlightthickness=1)
yscroll_log.config(command=log_text.yview)
xscroll_log.config(command=log_text.xview)
log_text.insert('end', LOG_DEFAULT)
log_text.config(state='disabled')
log_text.grid(row=0,column=0)
xscroll_log.grid(row=1,column=0, pady=(0,10), sticky='we')
yscroll_log.grid(row=0,column=1, sticky='ns')
log_frm.pack()

exact_cb_substitute_lbl = tk.Label()
exact_var = tk.IntVar()
exact_cb = ttk.Checkbutton(text='Exact order', variable=exact_var)
exact_cb.grid(**EXACT_CB_GRID_ARGS)

# ----- 'put in columns' and 'put in rows' buttons -----

excel_var = tk.IntVar(value=True)
col_var = tk.IntVar(value=True)

column_row_frm = tk.Frame()
col_rb = ttk.Radiobutton(column_row_frm,text='Put in columns', variable=col_var, value=True)
row_rb = ttk.Radiobutton(column_row_frm,text='Put in rows', variable=col_var, value=False)
col_rb.pack(anchor='w')
row_rb.pack(anchor='w')
column_row_frm.grid(row=1,column=1, sticky='s')

# ----- patterns menu -----

patterns_list_frm = tk.Frame(bd = 10)
yscroll_pl = tk.Scrollbar(patterns_list_frm)
xscroll_pl = tk.Scrollbar(patterns_list_frm, orient='horizontal')
pattern_lbl = tk.Label(patterns_list_frm,text='Patterns:')
patterns_list = tk.Listbox(patterns_list_frm,width=25,height=13, yscrollcommand=yscroll_pl.set,
                           xscrollcommand=xscroll_pl.set,selectmode='extended')
xscroll_pl.config(command=patterns_list.xview)
yscroll_pl.config(command=patterns_list.yview)
pattern_lbl.grid(row=0,column=0, sticky='w')
patterns_list.grid(row=1,column=0)
xscroll_pl.grid(row=2,column=0, sticky='we')
yscroll_pl.grid(row=1,column=1,sticky='ns')
patterns_list_frm.grid(row=2,column=1, sticky='s')

# ----- creating context menus for 'input file entry', 'output file entry',
# 'sheet name entry', 'log text' and 'patterns list' -----

create_commands_objects(None,log_text ,window, patterns_list, exact_var,
                            exact_cb, exact_cb_substitute_lbl, sheet_name_lbl,
                            sheet_name_entry, col_var, excel_var)

input_file_menu = MenuCreators.create_entry_menu(input_file_entry, excel_var, is_output_file_entry=False)
output_file_menu = MenuCreators.create_entry_menu(output_file_entry, excel_var)
sheet_name_menu = MenuCreators.create_entry_menu(sheet_name_entry, excel_var, False)
patterns_menu  = MenuCreators.create_patterns_menu()
log_menu = MenuCreators.create_log_menu()

CommandsObjects.log_menu_commands.log_menu = log_menu

# ----- context menu displayer functions for menus which defined in the last sections -----

context_menu_displayer = ContextMenuDisplayers(log_text, log_menu, patterns_menu, patterns_list, window,
                                               sheet_name_entry, input_file_entry, output_file_entry,
                                               sheet_name_menu, input_file_menu, output_file_menu)

input_file_entry.bind('<Button-3>',lambda event : context_menu_displayer.show_entry_menu(input_file_menu,event))
output_file_entry.bind('<Button-3>',lambda event : context_menu_displayer.show_entry_menu(output_file_menu,event))
sheet_name_entry.bind('<Button-3>', lambda event : context_menu_displayer.show_entry_menu(sheet_name_menu,event))
patterns_list.bind('<Button-3>', context_menu_displayer.show_patterns_menu)
log_text.bind('<Button-3>',lambda event : context_menu_displayer.show_log_menu(event))

try:
    context_menu_displayer.set_keysym('<Shift-F10>')

    if os.name == 'nt':
        context_menu_displayer.set_keysym('<App>')
    else:
        context_menu_displayer.set_keysym('<Menu>')

except tk.TclError:
    pass

# ----- keyboard shortcuts for context menus options -----

window.bind_class('TEntry','<Control-C>',lambda event : event.widget.delete(0,'end'))

log_text.bind('<Control-c>', CommandsObjects.log_menu_commands.copy_log)
log_text.bind('<Control-d>', CommandsObjects.log_menu_commands.clear_log)

input_file_entry.bind('<Control-b>',lambda event  : browse_files(event.widget, True))
output_file_entry.bind('<Control-b>',lambda event : browse_files(event.widget, False))

if os.name == "posix":
    window.bind_class("TEntry","<Control-a>", lambda event : event.widget.select_range(0,'end'))

patterns_list.bind('<Control-A>', CommandsObjects.patterns_menu_commands.add_pattern)
patterns_list.bind('<Control-i>', CommandsObjects.patterns_menu_commands.insert_pattern)
patterns_list.bind('<k>', CommandsObjects.patterns_menu_commands.move_selected)
patterns_list.bind('<j>', lambda event : CommandsObjects.patterns_menu_commands.move_selected(up = False))
patterns_list.bind('<F2>', CommandsObjects.patterns_menu_commands.edit_selected)
patterns_list.bind('<Delete>', CommandsObjects.patterns_menu_commands.delete_selected)
patterns_list.bind('<Control-d>', CommandsObjects.patterns_menu_commands.delete_selected)
patterns_list.bind('<Control-c>', CommandsObjects.patterns_menu_commands.copy_pattern)
patterns_list.bind('<Control-D>', CommandsObjects.patterns_menu_commands.delete_all)
patterns_list.bind('<Control-C>',lambda event : CommandsObjects.patterns_menu_commands.copy_pattern(all=True))
patterns_list.bind('<Control-I>', CommandsObjects.patterns_menu_commands.import_from_file)
patterns_list.bind('<Control-e>', CommandsObjects.patterns_menu_commands.export_to_file)

# ----- setting command for 'put in rows' and 'put in columns' radio buttons -----

col_rb.config(command=CommandsObjects.csv_excel_switch_functions.show_exact_order_cb)
row_rb.config(command=CommandsObjects.csv_excel_switch_functions.hide_exact_order_cb)

# ----- The convert button -----

data_extractor = DataExtractor(excel_var, log_text, col_var, exact_var)

btn_convert = tk.Button(frm,text='convert',width=10,height=5,background='#0080e5',
                        command=lambda : data_extractor.prepare_to_extract_data(output_file_entry.get(),
                        input_file_entry.get(), sheet_name_entry.get(), patterns_list.get(0,'end')), cursor='hand2')

btn_convert.bind('<Enter>', lambda event : btn_convert.config(bg = '#0092ff'))
btn_convert.bind('<Leave>', lambda event : btn_convert.config(bg = '#0080e5'))
btn_convert.grid(row=0,column=1)

# ----- preparing and launching the program -----

input_file_entry.focus_set()
log_text.bind('<FocusOut>',lambda event : log_text.tag_remove('sel','1.0','end'))

window.mainloop()
