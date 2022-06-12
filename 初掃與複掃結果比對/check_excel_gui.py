import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter

import pandas as pd

win = tk.Tk()
win.title("excel檔案讀取")
win.geometry("1200x750")
# 視窗最大化
# w, h = win.maxsize()
# win.geometry(f"{w}x{h}")

#禁止子控件重置其尺寸
win.pack_propagate(0)
#不可以變更視窗大小
win.resizable(0, 1)

#===============================初掃==================================


def open_file1():
    filename1 = filedialog.askopenfilename(initialdir="c:\\tmp",
                                           title="請選擇一個檔案",
                                           filetypes=(("xlsx file", "*.xlsx"), ("other", "*.*")))
    #需將檔案路徑儲存到tk label，才有辦法將值給予其他涵式使用
    label_file1["text"] = filename1
    return None


def load_file1():
    try:
        file_path1 = label_file1["text"]
        df1 = pd.read_excel(file_path1)
        print(df1)
    except ValueError:
        # tk.Message.showerre("訊息:", "選擇了無效的檔案")
        messagebox.askokcancel(title="訊息", message="選擇了無效的檔案，重新選擇")
        return None
    except FileNotFoundError:
        messagebox.askokcancel(title="訊息", message=f"{file_path1}檔案不存在")
        return None

    #清除遺留在treeview的舊資料
    clear_data1()

    #設定將excel資料呈現於treeview資料框
    tv1["column"] = list(df1.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df1.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None


def clear_data1():
    tv1.delete(*tv1.get_children())


#設定Treeview的標籤框
frame1 = tk.LabelFrame(win, text="初掃Excel Data")
frame1.place(height=500, width=500, rely=0.01, relx=0.55)

#設定開啟檔的標籤框
file1_frame = tk.LabelFrame(win, text="Open File")
file1_frame.place(height=200, width=500, rely=0.70, relx=0.55)

#設定檔案標籤
label_file1 = tk.Label(text="尚未選擇初掃檔案")
label_file1.place(rely=0.75, relx=0.57)

#設定按鈕
btn1 = tk.Button(text="1.開啟弱掃初掃檔", command=lambda: open_file1())
btn1.place(rely=0.80, relx=0.85)  # 0,0代表為視窗左上解
btn2 = tk.Button(text="2.載入弱掃初掃檔", command=lambda: load_file1())
btn2.place(rely=0.85, relx=0.85)

#建立Treeview，顯示excel內容，並設定滑桿
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1)
treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)
tv1.config(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")

#===============================複掃==================================


def open_file2():
    filename2 = filedialog.askopenfilename(initialdir="c:\\tmp",
                                           title="請選擇一個檔案",
                                           filetypes=(("xlsx file", "*.xlsx"), ("other", "*.*")))
    #需將檔案路徑儲存到tk label，才有辦法將值給予其他涵式使用
    label_file2["text"] = filename2


def load_file2():
    try:
        file_path2 = label_file2["text"]
        df2 = pd.read_excel(file_path2)
        print(df2)
    except ValueError:
        # tk.Message.showerre("訊息:", "選擇了無效的檔案")
        messagebox.askokcancel(title="訊息", message="選擇了無效的檔案，重新選擇")
        return None
    except FileNotFoundError:
        messagebox.askokcancel(title="訊息", message=f"{file_path2}檔案不存在")
        return None

    #清除遺留在treeview的舊資料
    clear_data2()

    #設定將excel資料呈現於treeview資料框
    tv2["column"] = list(df2.columns)
    tv2["show"] = "headings"
    for column in tv2["columns"]:
        tv2.heading(column, text=column)

    df_rows = df2.to_numpy().tolist()
    for row in df_rows:
        tv2.insert("", "end", values=row)
    return None


def clear_data2():
    tv2.delete(*tv2.get_children())


#設定Treeview的標籤框
frame2 = tk.LabelFrame(win, text="複掃Excel Data")
frame2.place(height=500, width=500, rely=0.01, relx=0.05)

file1_frame = tk.LabelFrame(win, text="Open File")
file1_frame.place(height=200, width=500, rely=0.70, relx=0.05)

label_file2 = tk.Label(text="尚未選擇複掃檔案")
label_file2.place(rely=0.75, relx=0.07)

btn3 = tk.Button(text="3.開啟弱掃複掃檔", command=lambda: open_file2())
btn3.place(rely=0.80, relx=0.35)  # 0,0代表為視窗左上解
btn4 = tk.Button(text="4.載入弱掃複掃檔", command=lambda: load_file2())
btn4.place(rely=0.85, relx=0.35)

#建立Treeview，顯示excel內容，並設定滑桿
tv2 = ttk.Treeview(frame2)
tv2.place(relheight=1, relwidth=1)
treescrolly2 = tk.Scrollbar(frame2, orient="vertical", command=tv2.yview)
treescrollx2 = tk.Scrollbar(frame2, orient="horizontal", command=tv2.xview)
tv2.config(xscrollcommand=treescrollx2.set, yscrollcommand=treescrolly2.set)
treescrollx2.pack(side="bottom", fill="x")
treescrolly2.pack(side="right", fill="y")

#===============================比對==================================

btn5 = tk.Button(text="5.資料比對")
btn5.place(rely=0.80, relx=0.48)  # 0,0代表為視窗左上解

win.mainloop()
