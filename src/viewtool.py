import tkinter.filedialog as Dialog
import tkinter
import os
import TkinterDnD2
def set_json_dir(root, file_content, filetypes = None, title = "Open", initialdir = None):
    filename=Dialog.askopenfilename(master = root, filetypes=filetypes, title = title, initialdir = initialdir)
    file_content.set(filename)

def get_screen_size(window):
    return window.winfo_screenwidth(),window.winfo_screenheight()
 
def get_window_size(window):
    return window.winfo_reqwidth(),window.winfo_reqheight()
 
def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)

def _drop(event):
    if event.data:
        widget = event.widget
        if isinstance(widget, tkinter.Entry):
            filename = widget.tk.splitlist(event.data)[0]
            if os.path.isfile(filename):
                widget.delete(0, tkinter.END)
                widget.insert(tkinter.END, filename)
    return event.action

def bind_drop(widget):
    widget.drop_target_register(TkinterDnD2.DND_FILES)
    widget.dnd_bind('<<Drop>>', _drop)
