import tkinter


class FileOpenUI(tkinter.Frame):
    def __init__(self, master, file_types = [("All", "*.*")], base_root = "", title = "Open", label = "Open", **kw):
        tkinter.Frame.__init__(self, master =master)
        self._text_value, = tkinter.StringVar(value = ""),
        self._input = tkinter.Entry(master=self, textvariable = self._text_value)
        self._open_btn = tkinter.Button(master = self, text = label, command = self._open_click)
        self._input.pack(anchor = tkinter.W, expand = True, fill = tkinter.X, side = tkinter.LEFT)
        self._open_btn.pack(side = tkinter.LEFT)
        self.file_types = file_types
        self.base_root = base_root
        self.title = title 

    def _open_click(self):
        if type(self.base_root) == str:
            filename=tkinter.filedialog.askopenfilename(master = self.winfo_toplevel(), filetypes=self.file_types, title = self.title, initialdir = self.base_root)
        else:
            filename=tkinter.filedialog.askopenfilename(master = self.winfo_toplevel(), filetypes=self.file_types, title = self.title, initialdir = self.base_root())
        self._text_value.set(filename)
    
    def get_value(self):
        return self._text_value.get()

        