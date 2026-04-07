import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def merge_excel():
    files = filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx")])
    if not files:
        return
    all_df = []
    for f in files:
        df = pd.read_excel(f)
        all_df.append(df)
    total = pd.concat(all_df, ignore_index=True)
    save_path = os.path.join(os.path.dirname(files[0]),"合并总表.xlsx")
    total.to_excel(save_path,index=False)
    messagebox.showinfo("完成",f"已合并保存：{save_path}")

root=tk.Tk()
root.title("多Excel合并工具")
root.geometry("400x200")
tk.Button(root,text="选择多个Excel并合并",command=merge_excel,font=("微软雅黑",12)).pack(pady=60)
root.mainloop()