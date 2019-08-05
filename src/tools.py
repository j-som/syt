import tkinter as tk 
import tkinter.filedialog as Dialog
import os
import json 
import translateapi
import openpyxl
import re
import viewtool
import config
class DiffLan(tk.Frame):
    def __init__(self, root, title):
        tk.Frame.__init__(self, master=root)
        self.master = root
        root.withdraw()
        root.title(title)
        self.init()
        viewtool.center_window(root, 755, 212)
        # root.update_idletasks()
        # root.deiconify()


    def excel_to_dict(self, ws):
        s = {}
        for i in ws.iter_rows():
            src = i[0].value
            dest = i[1].value 
            s[src] = dest
        return s

    def diff_language(self, file1, file2):
        excel1 = file1.get()
        excel2 = file2.get()
        wb = openpyxl.load_workbook(excel1, read_only=True)
        ws = wb[wb.sheetnames[0]]
        dict1 = self.excel_to_dict(ws)
        wb.close()
        wb = openpyxl.load_workbook(excel2, read_only=True)
        ws = wb[wb.sheetnames[0]]
        dict2 = self.excel_to_dict(ws)
        wb.close()
        newdict = {}
        for k, v in dict1.items():
            if not k in dict2 or dict2[k] != v:
                newdict[k] = v 
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(1, 1).value = u"中文"
        ws.cell(1, 2).value = u"译文"
        row = 2
        for k, v in newdict.items():
            ws.cell(row, 1).value = k
            ws.cell(row, 2).value = v
            row += 1
        (_, file_name) = os.path.split(excel1)
        (short_name, _) = os.path.splitext(file_name)
        save_dir = Dialog.asksaveasfilename(master = self.master, filetypes=[("Excel File", "*.xlsx")], title = "请选择保存Excel文件的路径", initialdir = excel_root, initialfile = short_name)
        if save_dir == "":
            return
        (save_dir, file_name) = os.path.split(save_dir)
        (short_name, _) = os.path.splitext(file_name)
        save_full_name = os.path.join(save_dir, short_name + ".xlsx")
        wb.save(save_full_name)
        wb.close()

    def init(self):
        wnd = tk.Frame(self)
        title1_2 = tk.Label(master=wnd, text = "新Excel:")
        file1_2 = tk.StringVar(value = "")
        input1_2 = tk.Entry(master=wnd, textvariable=file1_2, width = 100)
        btn1_2 = tk.Button(master=wnd, text="打开", command=lambda : viewtool.set_json_dir(self.master, file1_2, filetypes=[("Excel File", "*.xlsx")], title="翻译文件", initialdir=config.excel_root()))
        title2_2 = tk.Label(master=wnd, text = "旧Excel:")
        file2_2 = tk.StringVar(value = "")
        input2_2 = tk.Entry(master=wnd, textvariable=file2_2, width = 100)
        btn2_2 = tk.Button(master=wnd, text="打开", command=lambda : viewtool.set_json_dir(self.master, file2_2, filetypes=[("Excel File", "*.xlsx")], title="翻译文件", initialdir=config.excel_root()))
        pickbtn = tk.Button(master=wnd, text="提取", command=lambda : self.diff_language(file1_2, file2_2))
        row = 0    
        title1_2.grid(row = row, column = 1, columnspan = 10, sticky = tk.W)
        row += 1
        input1_2.grid(row = row, column = 1, columnspan = 10)
        btn1_2.grid(row = row, column = 11)
        row += 1
        title2_2.grid(row = row, column = 1, columnspan = 10, sticky = tk.W)
        row += 1
        input2_2.grid(row = row, column = 1, columnspan = 10)
        btn2_2.grid(row = row, column = 11)
        row += 1
        pickbtn.grid(row = row, column = 1, sticky = tk.W)
        wnd.pack()

if __name__ == "__main__":
    DiffLan(TkinterDnD.Tk(), u"抽取未翻译的原文")