""" 
" 顶层入口
" 
"""
import tkinter.filedialog as Dialog
import tkinter.messagebox as MsgBox
import openpyxl
import json
import re
import os.path as path
import tkinter as tk
import translateapi
import os
import views
import viewtool
import sytlog
import config
import TkinterDnD2.TkinterDnD as TkinterDnD

def main():
    root = TkinterDnD.Tk()
    root.withdraw()
    root.title(u"翻译工具")
    viewtool.center_window(root, 650, 680)
    # tv.pack(anchor = tk.W)
    # pv.pack(anchor = tk.W)
    tabframe = views.TabsView(master=root)
    tabframe.pack(anchor = tk.W)
    tabs = {
        u'翻译':views.TransView,
        u'提取':views.PickView,
        u'更新':views.UpdateView,
        u'繁体':views.TwView,
        u'工具':views.ToolsView
        }
    for tabname in tabs:
        tabframe.add(tabname, tabs[tabname])
    logpanel = views.LogPanel(root)
    logpanel.pack(expand = True, fill = tk.BOTH)
    sytlog.log_panel = logpanel
    # tk.Button(master=root, text="test",command = lambda:logpanel.print("hahaha")).pack()
    with open('config.json', 'r', encoding='UTF-8') as f:
        cfg = json.load(f)
        config.setcfg(cfg)
        menu = tk.Menu(root)
        item = tk.Menu(menu, tearoff = False)
        v = tk.StringVar(value = config.CN)
        config.choose(config.CN)
        for vsn, vsncfg in cfg.items():
            item.add_radiobutton(label=vsncfg['name'], value = vsn, variable =v, command = lambda : config.choose(v.get()))
        menu.add_cascade(label = u"版本", menu = item)
        root.config(menu = menu)
    # logpanel.dnd_bind()
    root.update_idletasks()
    root.deiconify()
    root.mainloop()


if __name__ == '__main__':
	main()