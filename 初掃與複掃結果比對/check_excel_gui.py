from fileinput import filename
from datetime import date
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
import pandas as pd
import os



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

#提前宣告涵室內的變數為空值
file_path1 = ""
file_path2 = ""
save_path = ""
df_new = ""
#===============================初掃==================================
def open_file1():
    filename1 = filedialog.askopenfilename(initialdir="c:\\tmp",
                                           title="請選擇一個檔案",
                                           filetypes=(("xlsx file", "*.xlsx"), ("other", "*.*")))
    #需將檔案路徑儲存到tk label，才有辦法將值給予其他涵式使用
    label_file1["text"] = filename1
    return None

def load_file1():
    #將local變數宣告成global變數，以便引用給其他涵式使用
    global file_path1
    try:
        file_path1 = label_file1["text"]
        df1 = pd.read_excel(file_path1)
        # print(df1)
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
    #將local變數宣告成global變數，以便引用給其他涵式使用
    global file_path2
    try:
        file_path2 = label_file2["text"]
        df2 = pd.read_excel(file_path2)
        # print(df2)
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
def diff_file():
    global df_new
    print(file_path1)
    print(file_path2)
    file1name = file_path1.split('/')
    file1name = file1name[2]
    file2name = file_path2.split('/')
    file2name = file1name[2]
   
    #================================變數設定==================================================
    today = date.today()
    today = today.strftime("%Y%m%d")
    new_excel = "弱掃複掃結果清單"
    old_excel = "弱掃初掃結果清單"
    #================================1. 將利用read excel產生txt檔==============================

    '''
    這邊必須將excel轉成txt方便進行欄位內容比對，且需將"弱點解決方法"欄位移除，否則會因為內容有換行而導致
    資料的欄位數(columns)無法固定，進而導致部分比對會異常。
    '''
    #取得弱掃初掃結果清單
    df_old = pd.read_excel(
        rf'{file_path1}', sheet_name='High-level-ALL')
    df_old.drop("弱點解決方法", axis=1, inplace=True)  # 刪除欄位
    df_old.to_csv(rf'c:\tmp\{file1name}.txt', sep=';',
                index=False)  # 依據分號;進行欄位切割，然後以csv形式存成txt檔案

    #弱掃複掃結果清單
    df_new = pd.read_excel(
        rf'{file_path2}', sheet_name='High-level-ALL')
    df_new.drop("弱點解決方法", axis=1, inplace=True)  # 刪除欄位
    df_new.to_csv(rf'c:\tmp\{file2name}.txt', sep=';',
                index=False)  # 依據分號;進行欄位切割，然後以csv形式存成txt檔案


    #=====================================2. 讀取txt檔案=======================================

    #將取得弱掃初掃.txt的內容依序存入lista列表中
    lista = []
    old = open(rf'c:\tmp\{file1name}.txt', encoding='utf-8',)
    for line in old.readlines():
        #print(line)
        lista.append(line.replace('\n', ''))  # 將每筆資料存入lista中
    #移除不需要的欄位
    lista.remove("任務名稱;任務季度;漏洞類型;弱點描述;任務執行時間;IP;Hostname;系統管理者;OS版本;Port;Port Protocol;CVSS分數;風險等級;CVEs編號;處理步驟;是否可修補(Y/N);是否為新弱點(Y/N);完成確認日期")

    # print("lista=",lista)

    #將取得弱掃複掃.txt的內容依序存入listb列表中
    listb = []
    new = open(rf'c:\tmp\{file2name}.txt', encoding='utf-8')
    for line in new.readlines():
        #print(line)
        listb.append(line.replace('\n', ''))  # 將每筆資料存入listb中
    #移除不需要的欄位
    listb.remove("任務名稱;任務季度;漏洞類型;弱點描述;任務執行時間;IP;Hostname;系統管理者;OS版本;Port;Port Protocol;CVSS分數;風險等級;CVEs編號;處理步驟;是否可修補(Y/N);是否為新弱點(Y/N);完成確認日期")
    # print("listb=",listb)

    # #======================================取得未修補弱點的名稱與IP============================
    #將list轉換為集合，初掃(old)-複掃(new)會取得未修補的漏洞資料
    x_old = set(lista)
    y_old = set(listb)
    z_old = x_old - y_old
    # print("x1 - y1=\n",z1)

    #再將結果由集合轉為list，以便產生有序的列表
    listc_old = list(z_old)
    list_count_old = len(listc_old)  # 取得未修漏洞筆數
    # print((list_count_old))

    #建立空的序列，以便將所有未修補漏洞填入
    listc_name_old = []
    listc_ip_old = []
    listc_hostname_old = []
    listc_admin_old = []
    listc_os_old = []
    listc_status_old = []
    listc_result_old = []
    listc_confirmdata_old = []

    #取得所有未修補弱點的特定欄位，並且依序存入序列中(0~16,弱點解決方法欄位被移除)
    for i in range(list_count_old):
        v1 = listc_old[i].split(';')
        #print(v1)
        listc_name_old.append(v1[3])  # 弱點描述
        listc_ip_old.append(v1[5])  # IP
        listc_hostname_old.append(v1[6])  # hostname
        listc_admin_old.append(v1[7])  # 系統管理者
        listc_os_old.append(v1[8])  # 作業系統版本
        listc_status_old.append(v1[14])  # 處理步驟
        listc_result_old.append(v1[15])  # 是否可修補(Y/N)
        listc_confirmdata_old.append(v1[17])  # 完成確認日期

    #======================================取得新弱點的弱點名稱與IP============================
    #轉換為集合取得新的弱點
    x_new = set(lista)
    y_new = set(listb)
    z_new = y_new - x_new
    # print("y-x=\n",z_new)
    #將集合結果轉成list，並將弱點名稱與ip切割
    listc_new = list(z_new)
    list_count_new = len(listc_new)
    listc_name_new = []  # 弱點描述
    listc_ip_new = []  # IP

    n = 0
    for i in range(list_count_new):
        v2 = listc_new[i].split(';')
        # print(len(v2))
        listc_name_new.append(v2[3])
        listc_ip_new.append(v2[5])
        # breakpoint()
    # print("listc_name=", listc_name_new)
    # print("listc_ip", listc_ip_new)

    # #=====================================進行資料修改=========================================
    #讀取原始excel檔案內容
    df_old = pd.read_excel(rf'{file_path1}', index_col=False)
    df_new = pd.read_excel(rf'{file_path2}', index_col=False)

    #將上次弱點修復結果的處理步驟、是否可修補(Y/N)欄位填入本次弱點掃描又掃到的項目中
    for i in range(list_count_old):
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), '系統管理者'] = listc_admin_old[i]
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), 'Hostname'] = listc_hostname_old[i]
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), 'OS版本'] = listc_os_old[i]
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), '是否可修補(Y/N)'] = listc_result_old[i]
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), '處理步驟'] = listc_status_old[i]
        df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (
            df_new['IP'] == listc_ip_old[i])), '完成確認日期'] = listc_confirmdata_old[i]
    #print(df_new)

    #將新弱點標示為Y，運行的前提: 當df_new裡和df_old有相同的筆的資料時(代表弱點沒有被解決所以又出現於df_new的掃描結果中)，
    #必須將df_old的"處理步驟"和""是否可修補(Y/N)"的內容填入df_new對應的欄位中，否則會因為資料不相同而被視為新的弱點("z2= y2 - x2"，會過濾出新的弱點的描述與IP)
    for i in range(list_count_new):
        # df_new.loc[~df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'N'
        # df_new.loc[((df_new['弱點描述'] == listc_name_new[i]) & (df_new['IP'] == listc_ip_new[i])), '是否為新弱點(Y/N)'] = 'Y'
        df_new.loc[df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'Y'
        df_new.loc[~df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'N'
    # print(df_new)
    print(type(df_new))
    # df_new.to_excel(rf'c:\tmp\{today}-弱掃初掃與複掃合併結果清單.xlsx',
    #                 sheet_name=f'High-level-ALL', index=False)

def save_file():
    save_path = filedialog.asksaveasfilename(initialdir="c:\\tmp",
                                           title="請選擇儲存位置",
                                           defaultextension=".xlsx",
                                           filetypes=(("xlsx file", "*.xlsx"),))
    print(save_path)
    df_new.to_excel(rf'{save_path}',
                    sheet_name=f'High-level-ALL', index=False)

btn5 = tk.Button(text="5.資料比對", command=lambda: diff_file())
btn5.place(rely=0.80, relx=0.48)  # 0,0代表為視窗左上解

btn6 = tk.Button(text="6.資料另存", command=lambda: save_file())
btn6.place(rely=0.85, relx=0.48)  # 0,0代表為視窗左上解

win.mainloop()
