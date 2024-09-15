import openpyxl, re
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from openpyxl.utils import get_column_letter
from os.path import isfile
from tkinter.simpledialog import askstring

APP_TITLE = 'Text2Excel'

show_error = lambda err : messagebox.showerror(title=APP_TITLE, message=err)

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

def copy_pattern(all=False):
    window.clipboard_clear()
    if all:
        window.clipboard_append('\n'.join(patterns_list.get(0,'end')))
    else:
        window.clipboard_append(patterns_list.get(patterns_list.curselection()))

def import_from_file():
    try:
        file_path = filedialog.askopenfilename(title='Import')
        if file_path:
            with open(file_path) as f:
                for i in f.read().strip().splitlines():
                    patterns_list.insert('end', i)
    except Exception as err:
         show_error(err)

def export_to_file():
    try:
        file_path = filedialog.asksaveasfilename(title='Export')
        if file_path:
            with open(file_path, 'w') as f:
                for i in patterns_list.get(0,'end'):
                    f.write(i + '\n')
    except Exception as err:
        show_error(err)

def edit_selected():
    index = patterns_list.curselection()
    value = patterns_list.get(index)
    new_value = askstring(title=APP_TITLE, prompt='Enter the pattern: ', initialvalue=value)
    if new_value:
        patterns_list.delete(index)
        patterns_list.insert(index, new_value)

def set_patterns(patterns):
    l=[]
    for x,y in enumerate(patterns, 1):
        if not '?P<item>' in y:
            y = '(?P<item>' + y + ')'
        if col_var.get():
            x = get_column_letter(x)
        l.append((str(x),y))
    return l

def show_menu(event, app=False):
    if patterns_list.curselection():
        menu.entryconfig(1, state='active')
        menu.entryconfig(2,state = 'active')
        menu.entryconfig(3,state = 'active')
    else:
        menu.entryconfig(1, state='disabled')
        menu.entryconfig(2,state = 'disabled')
        menu.entryconfig(3,state = 'disabled')
    if app:
        menu.tk_popup(patterns_list.winfo_rootx()+100,patterns_list.winfo_rooty()+100)
    else:
        menu.tk_popup(event.x_root, event.y_root)

def create_excel_file(output_file,input_file,sheet_name, patterns):
        try:
            log_string = ''
            with open(input_file,encoding='utf-8') as f:
                content = f.read()
            assert output_file, 'The name of output file is required.'
            if not isfile(output_file):
                wb = openpyxl.Workbook()
                wb.save(output_file)
                wb.close()
            sheet_name = sheet_name.title()

            wb = openpyxl.load_workbook(output_file)

            try:
                sheet=wb[sheet_name]
            except:
                 sheet = wb.create_sheet(sheet_name)

            i=1

            if not exact_var.get():
                if col_var.get():
                    max_index=sheet.max_row
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

def create_context_menu():
     menu = tk.Menu(tearoff=False)
     menu.add_command(label='Add Pattern', command=lambda : patterns_list.insert('end',askstring(title=APP_TITLE, prompt='Enter the pattern: ')))
     menu.add_command(label='Edit selected', command=edit_selected)
     menu.add_command(label='Delete selected', command=lambda : patterns_list.delete(patterns_list.curselection()))
     menu.add_command(label='Copy selected', command=lambda : copy_pattern())
     menu.add_command(label='Delete All', command=lambda : patterns_list.delete(0,'end'))
     menu.add_command(label='Copy All', command=lambda : copy_pattern(True))
     menu.add_separator()
     menu.add_command(label='Import from file', command=import_from_file)
     menu.add_command(label='Export to file', command=export_to_file)
     return menu

window =tk.Tk()
window.title(APP_TITLE)
window.resizable(False,False)
main_frm = tk.Frame()
frm = tk.Frame(main_frm,borderwidth=10)
input_file_lbl = tk.Label(frm,text='Input file:')
input_file_entry = ttk.Entry(frm,width=30)
output_file_lbl = tk.Label(frm,text='Output file:')
output_file_entry = ttk.Entry(frm,width=30)
sheet_name_lbl = tk.Label(frm,text='Sheet name:')
sheet_name_entry = ttk.Entry(frm,width=10)

input_file_lbl.grid(row=2,column=0,sticky='w')
input_file_entry.grid(row=3,column=0,sticky='w')
output_file_lbl.grid(row=2, column=3,sticky='w')
output_file_entry.grid(row=3,column=3,sticky='w')
sheet_name_lbl.grid(row=4,column=3,sticky='w')
sheet_name_entry.grid(row=5,column=3,sticky='w')

frm.pack()

log_frm = tk.Frame(main_frm)
yscroll_log = tk.Scrollbar(log_frm)
xscroll_log = tk.Scrollbar(log_frm, orient='horizontal')
log_text = tk.Text(log_frm,width=23, height=10, font = 'TkTextFont', wrap = 'none', 
                   yscrollcommand=yscroll_log.set, xscrollcommand=xscroll_log.set,takefocus=True,
                   highlightcolor='black',highlightthickness=1)
yscroll_log.config(command=log_text.yview)
xscroll_log.config(command=log_text.xview)
log_text.insert('end', 'log ...')
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
patterns_list = tk.Listbox(patterns_list_frm,width=25,height=13, yscrollcommand=yscroll_pl.set, xscrollcommand=xscroll_pl.set)
xscroll_pl.config(command=patterns_list.xview)
yscroll_pl.config(command=patterns_list.yview)
pattern_lbl.grid(row=0,column=0, sticky='w')
patterns_list.grid(row=1,column=0)
xscroll_pl.grid(row=2,column=0, sticky='we')
yscroll_pl.grid(row=1,column=1,sticky='ns')
patterns_list_frm.grid(row=2,column=1, sticky='s')

menu  = create_context_menu()
patterns_list.bind('<Button-3>', show_menu)
patterns_list.bind('<App>', lambda event : show_menu(event, True))

input_file_entry.focus_set()

btn_convert = tk.Button(frm,text='convert',width=10,height=5,background='#0080e5',
                        command=lambda : create_excel_file(output_file_entry.get(), input_file_entry.get(), sheet_name_entry.get(), 
                        set_patterns(patterns_list.get(0,'end')))
                        , cursor='hand2')
btn_convert.bind('<Enter>', lambda event : btn_convert.config(bg = '#0092ff'))
btn_convert.bind('<Leave>', lambda event : btn_convert.config(bg = '#0080e5'))
btn_convert.bind('<Return>', lambda event : btn_convert.invoke())
btn_convert.grid(row=0,column=1)

window.mainloop()
