import openpyxl, re
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from openpyxl.utils import get_column_letter
from os.path import isfile
from tkinter.simpledialog import askstring

APP_TITLE = 'Text2Excel'
ENCODING = 'utf-8-sig'
LOG_DEFAULT = 'log ...'

menu_color_args = {'activebackground' : '#00c8ff', 'activeforeground' : 'black'}

def ask():
    return askstring(title=APP_TITLE,prompt='Enter the pattern:')

def show_error(err):
    messagebox.showerror(title=APP_TITLE, message=err)

def delete_all(event=None):
    patterns_list.delete(0,'end')

def browse_files(widget):
    file_path = filedialog.askopenfilename(title='Browse')
    if file_path:
        widget.delete(0,'end')
        widget.insert('end',file_path)

def copy_log(event=None):
    window.clipboard_clear()
    try:
        data = log_text.selection_get()
    except tk.TclError:
        data = log_text.get('1.0','end')
    window.clipboard_append(data)

def clear_log(event=None):
    log_text.config(state='normal')
    log_text.delete('1.0','end')
    log_text.insert('end',LOG_DEFAULT)
    log_text.config(state='disabled')

def copy_pattern(event=None,all=False):
    window.clipboard_clear()
    if all:
        window.clipboard_append('\n'.join(patterns_list.get(0,'end')))
    else:
        selected = patterns_list.curselection()
        if selected:
            window.clipboard_append('\n'.join(patterns_list.get(selected[0],selected[-1])))

def import_from_file(event=None):
    try:
        file_path = filedialog.askopenfilename(title='Import')
        if file_path:
            with open(file_path, encoding=ENCODING) as f:
                for i in f.read().strip().splitlines():
                    patterns_list.insert('end', i)
    except Exception as err:
         show_error(err)

def export_to_file(event=None):
    try:
        file_path = filedialog.asksaveasfilename(title='Export')
        if file_path:
            with open(file_path, 'w') as f:
                for i in patterns_list.get(0,'end'):
                    f.write(i + '\n')
    except Exception as err:
        show_error(err)

def delete_selected(event=None):
    selected = patterns_list.curselection()
    if selected:
        patterns_list.delete(selected[0],selected[-1])

def edit_selected(event=None):
    index = patterns_list.curselection()
    if len(index) == 1:
        value = patterns_list.get(index)
        new_value = askstring(title=APP_TITLE, prompt='Enter the pattern: ', initialvalue=value)
        if new_value:
            patterns_list.delete(index)
            patterns_list.insert(index, new_value)
    
def add_pattern(event=None):
    new_pattern = ask()
    patterns_list.insert('end',new_pattern)

def insert_pattern(event=None):
    selected = patterns_list.curselection()
    if len(selected) == 1:
        new_pattern = ask()
        patterns_list.insert(selected[0],new_pattern)


def set_patterns(patterns):
    l=[]
    for x,y in enumerate(patterns, 1):
        if not '?P<item>' in y:
            y = '(?P<item>' + y + ')'
        if col_var.get():
            x = get_column_letter(x)
        l.append((str(x),y))
    return l

def find_max(wb, index, sheet_name):
    if col_var.get():
        row = 0
        for i in wb[sheet_name].iter_rows(min_col=index,max_col=index):
            if i[0].value is not None:
                row = i[0].row
        return row
    else:
        col = 0
        for i in wb[sheet_name].iter_cols(min_row=index,max_row=index):
            if i[0].value is not None:
                col = i[0].column
        return col

def create_excel_file(output_file,input_file,sheet_name, patterns):
        try:
            log_string = ''
            with open(input_file,encoding=ENCODING) as f:
                content = f.read()
            assert output_file, 'The name of output file is required.'
            if not isfile(output_file):
                wb = openpyxl.Workbook()
                wb.save(output_file)
                wb.close()
            sheet_name = sheet_name.title()

            wb = openpyxl.load_workbook(output_file)

            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
            else:
                sheet = wb.create_sheet(sheet_name)

            i=1

            if not exact_var.get():
                if col_var.get():
                    max_index = sheet.max_row
                else:
                    max_index = sheet.max_column
            else:
                max_index = find_max(wb,i, sheet_name)

            if col_var.get():
                get_index = lambda x, y : x + str(y)
            else:
                get_index = lambda x, y : get_column_letter(y) + x

            index=max_index+1
            for pattern in patterns:
                for item in re.finditer(pattern[1],content):
                    log_string += item.group('item') + '\n'
                    sheet[get_index(pattern[0], index)] = item.group('item')
                    index+=1
                i+=1
                if exact_var.get():
                    max_index = find_max(wb,i, sheet_name)
                index=max_index+1

            wb.save(output_file)
            log_string += f'\n{output_file!r} saved.' + '\n'
            log_text.config(state='normal')
            log_text.delete('1.0','end')
            log_text.insert('end', log_string)
            log_text.config(state='disabled')
            wb.close()
            log_text.see('end')

        except Exception as err:
             show_error(err)

def show_log_text_menu(event,app=False):
    if log_text.tag_ranges('sel'):
        text='Copy selected'
    else:
        text='Copy log'
    log_menu.entryconfig(0,label=text)

    if app:
        log_menu.tk_popup(log_text.winfo_rootx()+100,log_text.winfo_rooty()+100)
    else:
        log_menu.tk_popup(event.x_root,event.y_root)
        

def show_patterns_menu(event, app=False):
    selected = patterns_list.curselection()
    states = list()
    if selected:
        if len(selected)>1:
            for i in (1,3):
                patterns_menu.entryconfig(i,state='disabled')
            
            for i in range(4,6):
                patterns_menu.entryconfig(i,state='active')

        else:
            patterns_menu.entryconfig(1,state='active')
            for i in range(3,6):
                patterns_menu.entryconfig(i,state='active')
    else:
        patterns_menu.entryconfig(1,state='disabled')
        for i in range(3,6):
            patterns_menu.entryconfig(i,state='disabled')

    if app:
        patterns_menu.tk_popup(patterns_list.winfo_rootx()+100,patterns_list.winfo_rooty()+100)
    else:
        patterns_menu.tk_popup(event.x_root, event.y_root)

def show_entry_menu(menu,event,app=False):
    if window.focus_get() == event.widget:
        for i in range(4):
            menu.entryconfig(i,state='active')
    else:
        for i in range(4):
            menu.entryconfig(i,state='disabled')
    if app:
        if event.widget == sheet_name_entry:
            menu.tk_popup(event.widget.winfo_rootx()+50,event.widget.winfo_rooty()+25)
        else:
            menu.tk_popup(event.widget.winfo_rootx()+100,event.widget.winfo_rooty()+25)
    else:
        menu.tk_popup(event.x_root,event.y_root)

def create_patterns_menu():
     menu = tk.Menu(tearoff=False,**menu_color_args)
     menu.add_command(label='Add Pattern', command=add_pattern,accelerator='Ctrl+Shift+A')
     menu.add_command(label='Insert Pattern',command=insert_pattern,accelerator='Ctrl+I')
     menu.add_separator()
     menu.add_command(label='Edit selected', command=edit_selected,accelerator='F2')
     menu.add_command(label='Delete selected', command=delete_selected,accelerator='Delete')
     menu.add_command(label='Copy selected', command=copy_pattern,accelerator='Ctrl+C')
     menu.add_command(label='Delete All', command=delete_all,accelerator='Ctrl+Shift+D')
     menu.add_command(label='Copy All', command=lambda : copy_pattern(all=True),accelerator='Ctrl+Shift+C')
     menu.add_separator()
     menu.add_command(label='Import from file', command=import_from_file,accelerator='Ctrl+Shift+I')
     menu.add_command(label='Export to file', command=export_to_file,accelerator='Ctrl+E')
     return menu

def create_log_menu():
    menu = tk.Menu(tearoff=False,**menu_color_args)
    menu.add_command(label='Copy log',command=copy_log,accelerator='Ctrl+C')
    menu.add_command(label='Clear log',command=clear_log,accelerator='Ctrl+D')
    return menu

def create_entry_menu(widget,is_file_entry=True):
    menu = tk.Menu(tearoff=False,**menu_color_args)
    menu.add_command(label='Select All', accelerator='Ctrl+A',command=lambda : widget.select_range(0,'end'))
    menu.add_command(label='Copy', accelerator='Ctrl+C',command=lambda : widget.event_generate('<<Copy>>'))
    menu.add_command(label='Paste', accelerator='Ctrl+V',command=lambda : widget.event_generate('<<Paste>>'))
    menu.add_command(label='Cut', accelerator='Ctrl+X',command=lambda : widget.event_generate('<<Cut>>'))
    menu.add_separator()
    menu.add_command(label='Clear',accelerator='Ctrl+Shift+C',command=lambda : widget.delete(0,'end'))
    if is_file_entry:
        menu.add_command(label='Browse',command=lambda : browse_files(widget),accelerator='Ctrl+B')
    return menu

window =tk.Tk()
window.title(APP_TITLE)
window.resizable(False,False)

main_frm = tk.Frame()

frm = tk.Frame(main_frm,borderwidth=10)
frm.pack()

input_file_lbl = tk.Label(frm,text='Input file:')
output_file_lbl = tk.Label(frm,text='Output file:')
sheet_name_lbl = tk.Label(frm,text='Sheet name:')

input_file_entry = ttk.Entry(frm,width=30)
output_file_entry = ttk.Entry(frm,width=30)
sheet_name_entry = ttk.Entry(frm,width=10)

input_file_lbl.grid(row=2,column=0,sticky='w')
output_file_lbl.grid(row=2, column=3,sticky='w')
sheet_name_lbl.grid(row=4,column=3,sticky='w')

input_file_entry.grid(row=3,column=0,sticky='w')
output_file_entry.grid(row=3,column=3,sticky='w')
sheet_name_entry.grid(row=5,column=3,sticky='w')

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

main_frm.grid(row=0,column=0, rowspan=3)

exact_var = tk.IntVar()
exact_cb = ttk.Checkbutton(text='Exact order       ', variable=exact_var)
exact_cb.grid(row=0,column=1, sticky='s')

rbtn_frm = tk.Frame()
col_var = tk.IntVar(value=1)
col_rb = ttk.Radiobutton(rbtn_frm,text='Put in columns', variable=col_var, value=1)
row_rb = ttk.Radiobutton(rbtn_frm,text='Put in rows', variable=col_var, value=0)
col_rb.pack(anchor='w')
row_rb.pack(anchor='w')
rbtn_frm.grid(row=1,column=1, sticky='s')

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

input_file_menu = create_entry_menu(input_file_entry)
output_file_menu = create_entry_menu(output_file_entry)
sheet_name_menu = create_entry_menu(sheet_name_entry,False)
patterns_menu  = create_patterns_menu()
log_menu = create_log_menu()

patterns_list.bind('<Button-3>', show_patterns_menu)
patterns_list.bind('<App>', lambda event : show_patterns_menu(event, True))

log_text.bind('<Button-3>',lambda event : show_log_text_menu(event))
log_text.bind('<App>',lambda event : show_log_text_menu(event,True))

input_file_entry.bind('<Button-3>',lambda event : show_entry_menu(input_file_menu,event))
input_file_entry.bind('<App>',lambda event : show_entry_menu(input_file_menu,event,True))

output_file_entry.bind('<Button-3>',lambda event : show_entry_menu(output_file_menu,event))
output_file_entry.bind('<App>',lambda event : show_entry_menu(output_file_menu,event,True))

sheet_name_entry.bind('<Button-3>', lambda event : show_entry_menu(sheet_name_menu,event))
sheet_name_entry.bind('<App>',lambda event : show_entry_menu(sheet_name_menu,event,True))

window.bind_class('TEntry','<Control-C>',lambda event : event.widget.delete(0,'end'))

log_text.bind('<Control-c>',copy_log)
log_text.bind('<Control-d>',clear_log)

input_file_entry.bind('<Control-b>',lambda event  : browse_files(event.widget))
output_file_entry.bind('<Control-b>',lambda event : browse_files(event.widget))

input_file_entry.focus_set()

log_text.bind('<FocusOut>',lambda event : log_text.tag_remove('sel','1.0','end'))

patterns_list.bind('<Control-A>',add_pattern)
patterns_list.bind('<Control-i>',insert_pattern)
patterns_list.bind('<F2>',edit_selected)
patterns_list.bind('<Delete>',delete_selected)
patterns_list.bind('<Control-d>',delete_selected)
patterns_list.bind('<Control-c>',copy_pattern)
patterns_list.bind('<Control-D>',delete_all)
patterns_list.bind('<Control-C>',lambda event : copy_pattern(all=True))
patterns_list.bind('<Control-I>',import_from_file)
patterns_list.bind('<Control-e>',export_to_file)

btn_convert = tk.Button(frm,text='convert',width=10,height=5,background='#0080e5',
                        command=lambda : create_excel_file(output_file_entry.get(), input_file_entry.get(), sheet_name_entry.get(), 
                        set_patterns(patterns_list.get(0,'end')))
                        , cursor='hand2')

btn_convert.bind('<Enter>', lambda event : btn_convert.config(bg = '#0092ff'))
btn_convert.bind('<Leave>', lambda event : btn_convert.config(bg = '#0080e5'))
btn_convert.grid(row=0,column=1)

window.mainloop()
