import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime

# 全局变量：存储数据和统计结果
df = None
stat_result = None

# 新增：自定义绘图全局变量
x_var = None
y_var = None
agg_var = None
chart_type_var = None


def import_excel():
    """导入Excel文件，支持多sheet选择"""
    global df
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")],
        title="选择要分析的Excel文件"
    )
    if not file_path:
        return

    try:
        # 读取Excel所有sheet
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # 弹出sheet选择窗口
        sheet_root = tk.Toplevel()
        sheet_root.title("选择要分析的Sheet")
        sheet_root.geometry("300x200")

        tk.Label(sheet_root, text="请选择Sheet：", font=("微软雅黑", 12)).pack(pady=10)
        sheet_var = tk.StringVar(value=sheet_names[0])
        sheet_menu = ttk.OptionMenu(sheet_root, sheet_var, *sheet_names)
        sheet_menu.pack(pady=10, fill=tk.X, padx=20)

        def confirm_sheet():
            global df
            selected_sheet = sheet_var.get()
            df = pd.read_excel(file_path, sheet_name=selected_sheet)
            sheet_root.destroy()

            # 识别数值列，更新统计列选择框
            numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns.tolist()
            all_cols = df.columns.tolist()

            if not numeric_cols:
                messagebox.showwarning("提示", "该Sheet中无数值列，无法进行描述统计！")
                return

            # 清空并更新列选择框
            col_var.set("全部数值列")
            col_menu["menu"].delete(0, "end")
            col_menu["menu"].add_command(label="全部数值列", command=lambda: col_var.set("全部数值列"))
            for col in numeric_cols:
                col_menu["menu"].add_command(label=col, command=lambda c=col: col_var.set(c))

            # ====================== 新增：更新自定义绘图下拉框 ======================
            x_var.set("请选择X轴(分类)")
            y_var.set("请选择Y轴(数值)")

            x_menu["menu"].delete(0, "end")
            x_menu["menu"].add_command(label="请选择X轴(分类)", command=lambda: x_var.set("请选择X轴(分类)"))
            for c in all_cols:
                x_menu["menu"].add_command(label=c, command=lambda col=c: x_var.set(col))

            y_menu["menu"].delete(0, "end")
            y_menu["menu"].add_command(label="请选择Y轴(数值)", command=lambda: y_var.set("请选择Y轴(数值)"))
            for c in numeric_cols:
                y_menu["menu"].add_command(label=c, command=lambda col=c: y_var.set(col))

            messagebox.showinfo("成功",
                                f"✅ Excel导入成功！\n文件：{os.path.basename(file_path)}\nSheet：{selected_sheet}\n数值列：{', '.join(numeric_cols)}")

        tk.Button(sheet_root, text="确认", command=confirm_sheet, bg="#1677ff", fg="white", font=("微软雅黑", 11)).pack(
            pady=15)
        sheet_root.mainloop()

    except Exception as e:
        messagebox.showerror("导入失败", f"错误原因：{str(e)}")


def generate_statistics():
    """生成描述统计结果，支持单列/全部数值列"""
    global df, stat_result
    if df is None:
        messagebox.showwarning("提示", "请先导入Excel文件！")
        return

    selected_col = col_var.get()
    # 筛选数值列
    numeric_df = df.select_dtypes(include=["int64", "float64"])
    if selected_col != "全部数值列":
        if selected_col not in numeric_df.columns:
            messagebox.showwarning("提示", "所选列不是数值列，请重新选择！")
            return
        numeric_df = numeric_df[[selected_col]]

    # 生成完整描述统计（补充默认没有的统计项）
    basic_stats = numeric_df.describe()
    # 补充中位数（已包含在50%分位数，这里单独提取方便查看）
    median = numeric_df.median()
    # 补充众数
    mode = numeric_df.mode().iloc[0]  # 取第一个众数
    # 补充缺失值统计
    null_count = numeric_df.isnull().sum()
    null_rate = (null_count / len(numeric_df)) * 100
    # 补充方差
    variance = numeric_df.var()
    # 补充标准差（已包含，这里统一格式）
    std_dev = numeric_df.std()

    # 整合所有统计结果，生成DataFrame
    stat_result = pd.DataFrame({
        "样本数量": len(numeric_df),
        "缺失值数量": null_count,
        "缺失率(%)": null_rate.round(2),
        "平均值": basic_stats.loc["mean"].round(4),
        "中位数": median.round(4),
        "众数": mode.round(4),
        "最小值": basic_stats.loc["min"].round(4),
        "最大值": basic_stats.loc["max"].round(4),
        "标准差": std_dev.round(4),
        "方差": variance.round(4),
        "25%分位数": basic_stats.loc["25%"].round(4),
        "75%分位数": basic_stats.loc["75%"].round(4),
        "极差": (basic_stats.loc["max"] - basic_stats.loc["min"]).round(4)
    })

    # 显示结果到文本框
    stat_text.delete(1.0, tk.END)
    stat_text.insert(1.0, "=" * 80 + "\n")
    stat_text.insert(tk.END, f"📊 Excel数据描述统计结果\n")
    stat_text.insert(tk.END, f"📅 统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    stat_text.insert(tk.END, f"📋 统计列：{selected_col if selected_col != '全部数值列' else '所有数值列'}\n")
    stat_text.insert(tk.END, f"📈 样本总量：{len(numeric_df)} 行\n")
    stat_text.insert(tk.END, "=" * 80 + "\n\n")
    stat_text.insert(tk.END, stat_result.to_string())

    # 启用按钮
    btn_plot.config(state=tk.NORMAL)
    btn_export_excel.config(state=tk.NORMAL)
    btn_export_pdf.config(state=tk.NORMAL)
    btn_custom_plot.config(state=tk.NORMAL)

    messagebox.showinfo("统计完成", "✅ 描述统计已生成！")


# ====================== 原有图表功能（保留不动） ======================
def draw_charts():
    """绘制统计可视化图表（直方图+箱线图）"""
    global df, stat_result
    if stat_result is None:
        messagebox.showwarning("提示", "请先生成描述统计结果！")
        return

    selected_col = col_var.get()
    numeric_df = df.select_dtypes(include=["int64", "float64"])
    if selected_col != "全部数值列":
        numeric_df = numeric_df[[selected_col]]

    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    fig.suptitle(f"数据分布可视化（{selected_col if selected_col != '全部数值列' else '所有数值列'}）", fontsize=14,
                 fontweight="bold")

    for col in numeric_df.columns:
        ax1.hist(numeric_df[col].dropna(), bins=20, alpha=0.7, label=col, edgecolor="black")
    ax1.set_title("数据分布直方图", fontsize=12)
    ax1.set_xlabel("数值")
    ax1.set_ylabel("频数")
    ax1.legend()
    ax1.grid(alpha=0.3)

    numeric_df.boxplot(ax=ax2, grid=False)
    ax2.set_title("数据箱线图（异常值检测）", fontsize=12)
    ax2.set_ylabel("数值")
    ax2.tick_params(axis="x", rotation=45)

    plt.tight_layout()
    plt.show()


# ====================== ✅ 新增：自定义汇总图表（你要的功能） ======================
def draw_custom_chart():
    global df, x_var, y_var, agg_var, chart_type_var
    if df is None:
        messagebox.showwarning("提示", "请先导入Excel！")
        return

    x_col = x_var.get()
    y_col = y_var.get()
    agg = agg_var.get()
    chart_type = chart_type_var.get()

    if x_col == "请选择X轴(分类)" or y_col == "请选择Y轴(数值)":
        messagebox.showwarning("提示", "请选择X轴、Y轴字段！")
        return

    # 数据清洗
    df_plot = df[[x_col, y_col]].copy()
    df_plot[y_col] = pd.to_numeric(df_plot[y_col], errors="coerce")
    df_plot = df_plot.dropna()

    # 汇总计算
    if agg == "求和":
        res = df_plot.groupby(x_col)[y_col].sum()
    elif agg == "均值":
        res = df_plot.groupby(x_col)[y_col].mean()
    elif agg == "计数":
        res = df_plot.groupby(x_col)[y_col].count()
    else:
        res = df_plot.groupby(x_col)[y_col].sum()

    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.figure(figsize=(12, 6))

    if chart_type == "柱状图":
        res.plot(kind="bar", color="#1677ff")
    elif chart_type == "折线图":
        res.plot(kind="line", marker="o", color="#ff9800")
    elif chart_type == "饼图":
        res.plot(kind="pie", autopct="%.1f%%")
    else:
        res.plot(kind="bar", color="#1677ff")

    plt.title(f"{x_col} - {y_col} ({agg})", fontsize=14, fontweight="bold")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# ====================== 导出功能（保留） ======================
def export_result(file_type):
    """导出统计结果（Excel/PDF）"""
    global stat_result
    if stat_result is None:
        messagebox.showwarning("提示", "请先生成描述统计结果！")
        return

    if file_type == "excel":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"Excel描述统计结果_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        )
        if not save_path:
            return
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
            stat_result.to_excel(writer, sheet_name="描述统计结果", index=True)
            numeric_df = df.select_dtypes(include=["int64", "float64"])
            selected_col = col_var.get()
            if selected_col != "全部数值列":
                numeric_df = numeric_df[[selected_col]]
            numeric_df.to_excel(writer, sheet_name="原始数值数据", index=False)
        messagebox.showinfo("导出成功", f"✅ Excel统计结果已保存至：\n{save_path}")

    elif file_type == "pdf":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")],
            initialfile=f"Excel描述统计结果_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        )
        if not save_path:
            return
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.axis('tight')
        ax.axis('off')
        table_data = []
        headers = ["统计项"] + list(stat_result.columns)
        table_data.append(headers)
        for idx, row in stat_result.iterrows():
            table_row = [idx] + [f"{val:.4f}" if isinstance(val, float) else str(val) for val in row.values]
            table_data.append(table_row)
        table = ax.table(cellText=table_data[1:], colLabels=table_data[0], cellLoc='center', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 2)
        plt.title(f"Excel数据描述统计结果\n统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", fontsize=14,
                  fontweight="bold", pad=20)
        plt.savefig(save_path, bbox_inches='tight', dpi=300)
        plt.close()
        messagebox.showinfo("导出成功", f"✅ PDF统计结果已保存至：\n{save_path}")


# ====================== GUI界面（完全保留你的布局 + 新增一行自定义绘图） ======================
root = tk.Tk()
root.title("Excel数据描述统计")
root.geometry("900x650")
root.resizable(False, False)
root.config(bg="#f8f9fa")

# 标题
tk.Label(
    root, text="Excel数据描述统计",
    font=("微软雅黑", 18, "bold"),
    bg="#f8f9fa", fg="#2c3e50"
).pack(pady=15)


frame_custom = tk.Frame(root, bg="#f8f9fa")
frame_custom.pack(pady=5)

tk.Label(frame_custom, text="X轴(分类)：", font=("微软雅黑", 11), bg="#f8f9fa").grid(row=0, column=0, padx=5)
x_var = tk.StringVar(value="请选择X轴(分类)")
x_menu = ttk.OptionMenu(frame_custom, x_var, "请选择X轴(分类)")
x_menu.grid(row=0, column=1, padx=5)

tk.Label(frame_custom, text="Y轴(数值)：", font=("微软雅黑", 11), bg="#f8f9fa").grid(row=0, column=2, padx=5)
y_var = tk.StringVar(value="请选择Y轴(数值)")
y_menu = ttk.OptionMenu(frame_custom, y_var, "请选择Y轴(数值)")
y_menu.grid(row=0, column=3, padx=5)

tk.Label(frame_custom, text="汇总：", font=("微软雅黑", 11), bg="#f8f9fa").grid(row=0, column=4, padx=5)
agg_var = tk.StringVar(value="求和")
agg_menu = ttk.OptionMenu(frame_custom, agg_var, "求和", "求和", "均值", "计数")
agg_menu.grid(row=0, column=5, padx=5)

tk.Label(frame_custom, text="图表：", font=("微软雅黑", 11), bg="#f8f9fa").grid(row=0, column=6, padx=5)
chart_type_var = tk.StringVar(value="柱状图")
chart_menu = ttk.OptionMenu(frame_custom, chart_type_var, "柱状图", "柱状图", "折线图", "饼图")
chart_menu.grid(row=0, column=7, padx=5)

btn_custom_plot = tk.Button(frame_custom, text="生成汇总图表", command=draw_custom_chart, font=("微软雅黑", 11),
                            bg="#28a745", fg="white", state=tk.DISABLED)
btn_custom_plot.grid(row=0, column=8, padx=5)

# 操作区域
frame_operate = tk.Frame(root, bg="#f8f9fa")
frame_operate.pack(pady=10)

# 导入按钮
tk.Button(
    frame_operate, text="导入Excel文件", command=import_excel,
    font=("微软雅黑", 12, "bold"), bg="#1677ff", fg="white", width=15, relief="flat"
).pack(side=tk.LEFT, padx=5)

# 列选择
tk.Label(frame_operate, text="统计列：", font=("微软雅黑", 12), bg="#f8f9fa").pack(side=tk.LEFT, padx=10)
col_var = tk.StringVar(value="全部数值列")
col_menu = ttk.OptionMenu(frame_operate, col_var, "全部数值列")
col_menu.pack(side=tk.LEFT, padx=5)

# 统计按钮
tk.Button(
    frame_operate, text="生成描述统计", command=generate_statistics,
    font=("微软雅黑", 12, "bold"), bg="#28a745", fg="white", width=15, relief="flat"
).pack(side=tk.LEFT, padx=5)

# 绘图按钮
btn_plot = tk.Button(
    frame_operate, text="绘制可视化图表", command=draw_charts,
    font=("微软雅黑", 12, "bold"), bg="#ff9800", fg="white", width=15, relief="flat", state=tk.DISABLED
)
btn_plot.pack(side=tk.LEFT, padx=5)

# 导出按钮
btn_export_excel = tk.Button(
    frame_operate, text="导出Excel", command=lambda: export_result("excel"),
    font=("微软雅黑", 12, "bold"), bg="#9c27b0", fg="white", width=12, relief="flat", state=tk.DISABLED
)
btn_export_excel.pack(side=tk.LEFT, padx=5)

btn_export_pdf = tk.Button(
    frame_operate, text="导出PDF", command=lambda: export_result("pdf"),
    font=("微软雅黑", 12, "bold"), bg="#dc3545", fg="white", width=12, relief="flat", state=tk.DISABLED
)
btn_export_pdf.pack(side=tk.LEFT, padx=5)

# 统计结果显示区域
tk.Label(root, text="统计结果展示：", font=("微软雅黑", 12), bg="#f8f9fa").pack(pady=5, anchor=tk.W, padx=20)
stat_text = tk.Text(root, font=("Consolas", 10), wrap=tk.WORD)
stat_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

# 底部说明
tk.Label(
    root, text="💡 升级：支持【自定义X轴/Y轴/汇总方式/图表类型】一键生成业务汇总图表！",
    font=("微软雅黑", 9), bg="#f8f9fa", fg="#666"
).pack(pady=10)

root.mainloop()