import tkinter as tk
from tkinter import filedialog,messagebox
import pandas as pd

def clean_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
    if not path:return
    df = pd.read_excel(path)
    old = len(df)
    df = df.drop_duplicates()
    df = df.fillna("")
    new_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if new_path:
        df.to_excel(new_path,index=False)
        messagebox.showinfo("完成",f"原数据{old}行\n清洗后{len(df)}行")

root = tk.Tk()
root.title("Excel一键清洗去重")
root.geometry("400x200")
tk.Button(root,text="选择Excel开始清洗",command=clean_excel,font=("微软雅黑",13),bg="#28a745",fg="white").pack(expand=True)
root.mainloop()