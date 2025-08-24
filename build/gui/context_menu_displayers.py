from tkinter import Menu, Event
from os import name

def nt_get_label(self):
    if self.log_text.tag_ranges('sel'):
        return 'Copy selected'
    else:
        return 'Copy log'

def posix_get_label(self):
    return 'Copy log'


class ContextMenuDisplayers:

    def __init__(self,log_text, log_menu, patterns_menu, patterns_list, window,
                 sheet_name_entry,input_file_entry, output_file_entry,
                 sheet_name_menu, input_file_menu, output_file_menu):

        self.log_text = log_text
        self.log_menu = log_menu
        self.patterns_menu = patterns_menu
        self.patterns_list = patterns_list
        self.window = window

        self.sheet_name_entry = sheet_name_entry
        self.input_file_entry = input_file_entry
        self.output_file_entry = output_file_entry

        self.sheet_name_menu = sheet_name_menu
        self.input_file_menu = input_file_menu
        self.output_file_menu = output_file_menu

        if name == 'posix':
            self.get_label = posix_get_label
        else:
            self.get_label = nt_get_label

    def show_log_menu(self ,event : Event, app : bool = False) -> None:
        text = self.get_label(self)

        self.log_menu.entryconfig(0,label=text)

        if app:
            self.log_menu.tk_popup(self.log_text.winfo_rootx()+100, self.log_text.winfo_rooty()+100)
        else:
            self.log_menu.tk_popup(event.x_root,event.y_root)

    def disable_moveup_movedown(self):
        for i in (3,4):
            self.patterns_menu.entryconfig(i, state='disabled')

    def show_patterns_menu(self ,event : Event, app : bool=False) -> None:
        selected = self.patterns_list.curselection()
        if selected:
            if len(selected) > 1:
                for i in (1,6):
                    self.patterns_menu.entryconfig(i,state='disabled')

                for i in (7,8):
                    self.patterns_menu.entryconfig(i,state='active')

                self.disable_moveup_movedown()

            else:
                self.patterns_menu.entryconfig(1,state='active')

                for i in (6,7,8):
                    self.patterns_menu.entryconfig(i,state='active')

                for i in (3,4):
                    self.patterns_menu.entryconfig(i, state='active')
        else:
            self.patterns_menu.entryconfig(1,state='disabled')

            for i in (6,7,8):
                self.patterns_menu.entryconfig(i,state='disabled')

            self.disable_moveup_movedown()

        if app:
            self.patterns_menu.tk_popup(self.patterns_list.winfo_rootx()+100, self.patterns_list.winfo_rooty()+100)
        else:
            self.patterns_menu.tk_popup(event.x_root, event.y_root)

    def show_entry_menu(self, menu : Menu, event : Event, app : bool=False) -> None:
        if self.window.focus_get() == event.widget:
            for i in range(4):
                menu.entryconfig(i,state='active')
        else:
            for i in range(4):
                menu.entryconfig(i,state='disabled')
        if app:
            if event.widget == self.sheet_name_entry:
                menu.tk_popup(event.widget.winfo_rootx()+50,event.widget.winfo_rooty()+25)
            else:
                menu.tk_popup(event.widget.winfo_rootx()+100,event.widget.winfo_rooty()+25)
        else:
            menu.tk_popup(event.x_root,event.y_root)

    def set_keysym(self, menu_keysym):
        self.input_file_entry.bind(menu_keysym,lambda event : self.show_entry_menu(self.input_file_menu,event,True))
        self.output_file_entry.bind(menu_keysym,lambda event : self.show_entry_menu(self.output_file_menu,event,True))
        self.sheet_name_entry.bind(menu_keysym,lambda event : self.show_entry_menu(self.sheet_name_menu,event,True))
        self.log_text.bind(menu_keysym,lambda event : self.show_log_menu(event,True))
        self.patterns_list.bind(menu_keysym, lambda event : self.show_patterns_menu(event, True))
