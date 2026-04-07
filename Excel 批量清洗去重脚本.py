import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


def clean_excel():
    """Excel去重清洗主函数"""
    # 获取选中的文件路径
    file_path = file_entry.get().strip()
    if not file_path:
        messagebox.showerror("错误", "请先选择要清洗的Excel文件！")
        return

    # 校验文件是否存在
    if not os.path.exists(file_path):
        messagebox.showerror("错误", "所选文件不存在，请重新选择！")
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        # 执行去重操作
        df_clean = df.drop_duplicates()
        # 拼接保存路径
        dir_name = os.path.dirname(file_path)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        save_path = os.path.join(dir_name, f"{file_name}_已去重.xlsx")
        # 保存清洗后的数据
        df_clean.to_excel(save_path, index=False)
        messagebox.showinfo("完成", f"Excel清洗去重完成！\n已保存至：\n{save_path}")
    except Exception as e:
        messagebox.showerror("处理失败", f"文件处理出错：{str(e)}")


def select_file():
    """选择Excel文件"""
    path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
    )
    if path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, path)


# 初始化主窗口
root = tk.Tk()
root.title("Excel批量清洗去重工具")
root.geometry("550x260")
root.resizable(False, False)
root.config(bg="#f8f9fa")

# 标题
title_label = tk.Label(
    root,
    text="Excel数据去重清洗",
    font=("微软雅黑", 18, "bold"),
    bg="#f8f9fa",
    fg="#2c3e50"
)
title_label.pack(pady=18)

# 文件选择区域
frame = tk.Frame(root, bg="#f8f9fa")
frame.pack(pady=8)

file_entry = tk.Entry(
    frame,
    width=45,
    font=("微软雅黑", 12),
    bd=1,
    relief="solid"
)
file_entry.pack(side=tk.LEFT, padx=6)

select_btn = tk.Button(
    frame,
    text="选择文件",
    command=select_file,
    font=("微软雅黑", 11),
    bg="#428252",
    fg="white",
    relief="flat",
    padx=8,
    pady=2,
    cursor="hand2"
)
select_btn.pack(side=tk.LEFT)

# 处理按钮
clean_btn = tk.Button(
    root,
    text="开始去重清洗",
    command=clean_excel,
    font=("微软雅黑", 13, "bold"),
    bg="#167252",
    fg="white",
    width=18,
    height=2,
    relief="flat",
    cursor="hand2"
)
clean_btn.pack(pady=20)

# 备注
tip_label = tk.Label(
    root,
    text="支持.xlsx和.xls格式，自动去重并保存新文件",
    font=("微软雅黑", 9),
    bg="#f8f9fa",
    fg="#666666"
)
tip_label.pack()

# 运行窗口
root.mainloop()
