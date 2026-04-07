import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

# 全局变量：数据
df = None

def import_data():
    """导入Excel数据"""
    global df
    path = filedialog.askopenfilename(
        filetypes=[("Excel文件", "*.xlsx;*.xls")],
        title="选择数据文件"
    )
    if not path:
        return
    try:
        df = pd.read_excel(path)
        # 加载列名到下拉框
        cols = df.columns.tolist()
        # 清空下拉框
        row_var.set("")
        col_var.set("")
        value_var.set("")
        agg_var.set("sum")
        # 更新下拉框选项
        row_menu["menu"].delete(0, "end")
        col_menu["menu"].delete(0, "end")
        value_menu["menu"].delete(0, "end")
        for col in cols:
            row_menu["menu"].add_command(label=col, command=lambda c=col: row_var.set(c))
            col_menu["menu"].add_command(label=col, command=lambda c=col: col_var.set(c))
            value_menu["menu"].add_command(label=col, command=lambda c=col: value_var.set(c))
        messagebox.showinfo("成功", f"数据导入成功！共{df.shape[0]}行，{df.shape[1]}列")
    except Exception as e:
        messagebox.showerror("导入失败", f"错误原因：{str(e)}")

def generate_pivot():
    """生成透视表"""
    global df
    if df is None:
        messagebox.showwarning("提示", "请先导入数据！")
        return
    # 获取用户选择
    row = row_var.get()
    col = col_var.get()
    value = value_var.get()
    agg = agg_var.get()

    if not all([row, value]):
        messagebox.showwarning("提示", "行字段和值字段为必填项！")
        return

    try:
        # 生成透视表
        pivot = pd.pivot_table(
            df,
            index=row,
            columns=col if col else None,
            values=value,
            aggfunc=agg,
            fill_value=0,
            margins=True,
            margins_name="总计"
        )
        # 显示透视表
        pivot_text.delete(1.0, tk.END)
        pivot_text.insert(1.0, pivot.to_string())
        # 保存透视表
        global pivot_df
        pivot_df = pivot.reset_index()
        messagebox.showinfo("成功", "透视表生成完成！")
    except Exception as e:
        messagebox.showerror("生成失败", f"错误原因：{str(e)}")

def export_pivot():
    """导出透视表到Excel"""
    if df is None or "pivot_df" not in globals():
        messagebox.showwarning("提示", "请先生成透视表！")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel文件", "*.xlsx")],
        initialfile="透视表结果.xlsx"
    )
    if not save_path:
        return
    try:
        pivot_df.to_excel(save_path, index=False)
        messagebox.showinfo("成功", f"透视表已导出到：{save_path}")
    except Exception as e:
        messagebox.showerror("导出失败", f"错误原因：{str(e)}")

# GUI界面（统一风格）
root = tk.Tk()
root.title("终极版Excel透视表自动生成工具")
root.geometry("800x600")
root.resizable(False, False)
root.config(bg="#f8f9fa")

# 标题
tk.Label(
    root, text="Excel透视表自动生成工具",
    font=("微软雅黑", 18, "bold"),
    bg="#f8f9fa", fg="#2c3e50"
).pack(pady=15)

# 操作按钮
frame_btn = tk.Frame(root, bg="#f8f9fa")
frame_btn.pack(pady=10)

tk.Button(
    frame_btn, text="导入数据", command=import_data,
    font=("微软雅黑", 12, "bold"), bg="#1677ff", fg="white", width=12, relief="flat"
).pack(side=tk.LEFT, padx=5)

tk.Button(
    frame_btn, text="生成透视表", command=generate_pivot,
    font=("微软雅黑", 12, "bold"), bg="#28a745", fg="white", width=12, relief="flat"
).pack(side=tk.LEFT, padx=5)

tk.Button(
    frame_btn, text="导出透视表", command=export_pivot,
    font=("微软雅黑", 12, "bold"), bg="#ff9800", fg="white", width=12, relief="flat"
).pack(side=tk.LEFT, padx=5)

# 透视表设置
frame_set = tk.Frame(root, bg="#f8f9fa")
frame_set.pack(pady=10)

# 行字段
tk.Label(frame_set, text="行字段：", font=("微软雅黑", 12), bg="#f8f9fa").grid(row=0, column=0, padx=10, pady=5)
row_var = tk.StringVar()
row_menu = ttk.OptionMenu(frame_set, row_var, "")
row_menu.grid(row=0, column=1, padx=5, pady=5)

# 列字段
tk.Label(frame_set, text="列字段：", font=("微软雅黑", 12), bg="#f8f9fa").grid(row=0, column=2, padx=10, pady=5)
col_var = tk.StringVar()
col_menu = ttk.OptionMenu(frame_set, col_var, "")
col_menu.grid(row=0, column=3, padx=5, pady=5)

# 值字段
tk.Label(frame_set, text="值字段：", font=("微软雅黑", 12), bg="#f8f9fa").grid(row=1, column=0, padx=10, pady=5)
value_var = tk.StringVar()
value_menu = ttk.OptionMenu(frame_set, value_var, "")
value_menu.grid(row=1, column=1, padx=5, pady=5)

# 聚合方式
tk.Label(frame_set, text="聚合方式：", font=("微软雅黑", 12), bg="#f8f9fa").grid(row=1, column=2, padx=10, pady=5)
agg_var = tk.StringVar(value="sum")
agg_menu = ttk.OptionMenu(frame_set, agg_var, "sum", "count", "mean", "max", "min")
agg_menu.grid(row=1, column=3, padx=5, pady=5)

# 透视表显示
tk.Label(root, text="透视表结果：", font=("微软雅黑", 12), bg="#f8f9fa").pack(pady=5, anchor=tk.W, padx=20)
pivot_text = tk.Text(root, font=("Consolas", 10), wrap=tk.WORD)
pivot_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

root.mainloop()