from tkinter import Menu, Event

class ContextMenuDisplayers:

    def __init__(self,log_text, log_menu, patterns_menu, patterns_list, window, sheet_name_entry):
        self.log_text = log_text
        self.log_menu = log_menu
        self.patterns_menu = patterns_menu
        self.patterns_list = patterns_list
        self.window = window
        self.sheet_name_entry = sheet_name_entry

    def show_log_menu(self ,event : Event, app : bool = False) -> None:
        if self.log_text.tag_ranges('sel'):
            text='Copy selected'
        else:
            text='Copy log'
        self.log_menu.entryconfig(0,label=text)

        if app:
            self.log_menu.tk_popup(self.log_text.winfo_rootx()+100, self.log_text.winfo_rooty()+100)
        else:
            self.log_menu.tk_popup(event.x_root,event.y_root)
            

    def show_patterns_menu(self ,event : Event, app : bool=False) -> None:
        selected = self.patterns_list.curselection()
        if selected:
            if len(selected)>1:
                for i in (1,3):
                    self.patterns_menu.entryconfig(i,state='disabled')
                
                for i in range(4,6):
                    self.patterns_menu.entryconfig(i,state='active')

            else:
                self.patterns_menu.entryconfig(1,state='active')
                for i in range(3,6):
                    self.patterns_menu.entryconfig(i,state='active')
        else:
            self.patterns_menu.entryconfig(1,state='disabled')
            for i in range(3,6):
                self.patterns_menu.entryconfig(i,state='disabled')

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
