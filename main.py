import pandas as pd
import tkinter as tk
import os
from tkinter import ttk
from tkinter import messagebox



# 建立主視窗
root = tk.Tk()
root.title("Excel 資料搜尋")
root.geometry("950x500")

# 欄位名稱
columns = ("Membership ID", "Full Name", "Mobile Number", "Membership Email Address",
           "Membership Type", "Student Number", "University", "Degree")

# 讀取 CSV 檔案
filename = "C:/Users/Peter/Desktop/Python GUTS/membership.csv"
try:
    df = pd.read_csv(filename, encoding="utf-8")
except UnicodeDecodeError:
    try:
        df = pd.read_csv(filename, encoding="cp950")
    except Exception as e:
        messagebox.showerror("讀取錯誤", f"無法正確讀取 CSV 檔案：\n{e}")
        root.destroy()

# 搜尋功能
def search():
    query = entry_query.get().lower()
    selected_col = combo_column.get()

    if selected_col not in columns:
        messagebox.showerror("錯誤", "請選擇有效的搜尋欄位。")
        return

    result = df[df[selected_col].astype(str).str.lower().str.contains(query, na=False)]

    for row in tree.get_children():
        tree.delete(row)

    if not result.empty:
        for _, row in result.iterrows():
            tree.insert("", "end", values=tuple(row[col] for col in columns))
    else:
        messagebox.showinfo("查無資料", f"在「{selected_col}」中找不到包含「{query}」的資料。")

# 搜尋欄與欄位選擇區域
frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="輸入查詢關鍵字:").grid(row=0, column=0, padx=5)
entry_query = tk.Entry(frame)
entry_query.grid(row=0, column=1, padx=5)

tk.Label(frame, text="選擇欄位:").grid(row=0, column=2, padx=5)
combo_column = ttk.Combobox(frame, values=columns, state="readonly", width=30)
combo_column.set("Full Name")
combo_column.grid(row=0, column=3, padx=5)

btn_search = tk.Button(frame, text="搜尋", command=search)
btn_search.grid(row=0, column=4, padx=5)

# 表格與滾動軸區域
table_frame = tk.Frame(root)
table_frame.pack(expand=True, fill="both", padx=10, pady=10)

# 水平與垂直滾軸
x_scrollbar = tk.Scrollbar(table_frame, orient="horizontal")
y_scrollbar = tk.Scrollbar(table_frame, orient="vertical")

# Treeview 表格
tree = ttk.Treeview(table_frame, columns=columns, show="headings",
                    xscrollcommand=x_scrollbar.set,
                    yscrollcommand=y_scrollbar.set)

x_scrollbar.config(command=tree.xview)
y_scrollbar.config(command=tree.yview)

# 建立表頭與欄位格式
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor="center")

# 表格與滾軸佈局
tree.grid(row=0, column=0, sticky="nsew")
y_scrollbar.grid(row=0, column=1, sticky="ns")
x_scrollbar.grid(row=1, column=0, sticky="ew")
table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

# 讓觸控板或滑鼠滾輪也可以直接控制左右滑動
def on_mousewheel(event):
    if event.state == 0:
        if event.delta < 0:
            tree.xview_scroll(1, "pages")
        elif event.delta > 0:
            tree.xview_scroll(-1, "pages")
        return "break"

# Mac 和 Windows 的差異處理
tree.bind("<MouseWheel>", on_mousewheel)          # Windows 垂直滾輪事件（改作水平用）
tree.bind("<Button-4>", lambda e: tree.xview_scroll(-3, "units"))  # Linux 上左滑
tree.bind("<Button-5>", lambda e: tree.xview_scroll(3, "units"))   # Linux 上右滑

# 啟動主迴圈
root.mainloop()
