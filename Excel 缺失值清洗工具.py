import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox

def clean_null():
    f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
    if not f:
        return
    df = pd.read_excel(f)
    before = df.isnull().sum().sum()
    df = df.dropna()
    out = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if out:
        df.to_excel(out, index=False)
        messagebox.showinfo("完成", f"清理前空值：{before}\n已保存新文件")

root = tk.Tk()
root.title("Excel 缺失值清洗")
root.geometry("400x150")
tk.Button(root, text="选择Excel开始清洗", command=clean_null, font=("微软雅黑",12), bg="#28a745", fg="white").pack(expand=True)
root.mainloop()