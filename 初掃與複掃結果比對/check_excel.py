import pandas as pd
import numpy as np
from datetime import date
import os
"""
此程式目的為將"弱點掃描結果清單_new"與"弱點掃描結果清單_old"做差異比較(以下簡稱new和old):
1. 找出無法修復的弱點，將old的"處理步驟"與"是否可修補(Y/N)"填入new中
2. 找出所有新弱點的資料，並於"是否為新弱點(Y/N)"欄位上填入"Y"
當一筆資料在new有，但old卻沒有，這代表該筆資料是新的弱點，利用程式取出該筆資料的"弱點描述"與"IP"
與new做比對，當同時滿足以上的兩個條件，則在new滿足條件的該筆資料上的"是否為新弱點(Y/N)"欄位上填入"Y"
"""

#================================變數設定==================================================
today = date.today()
today = today.strftime("%Y%m%d")
new_excel = "弱掃複掃結果清單"
old_excel = "弱掃初掃結果清單"
#================================1. 將利用read excel產生txt檔==============================
#建立tmp目錄
try:
    os.mkdir(r"c:\tmp")
except Exception as error:
    print("c:\tmp目錄已存在")

'''
這邊必須將excel轉成txt方便進行欄位內容比對，且需將"弱點解決方法"欄位移除，否則會因為內容有換行而導致
資料的欄位數(columns)無法固定，進而導致部分比對會異常。
'''
#取得弱掃初掃結果清單
df_old = pd.read_excel(rf'c:\tmp\{old_excel}.xlsx', sheet_name='High-level-ALL')
df_old.drop("弱點解決方法",axis=1,inplace=True) #刪除欄位
df_old.to_csv(rf'c:\tmp\{old_excel}.txt', sep=';', index=False) #依據分號;進行欄位切割，然後以csv形式存成txt檔案

#弱掃複掃結果清單
df_new = pd.read_excel(rf'c:\tmp\{new_excel}.xlsx', sheet_name='High-level-ALL')
df_new.drop("弱點解決方法",axis=1,inplace=True) #刪除欄位
df_new.to_csv(rf'c:\tmp\{new_excel}.txt', sep=';', index=False) #依據分號;進行欄位切割，然後以csv形式存成txt檔案


#=====================================2. 讀取txt檔案=======================================

#將取得弱掃初掃.txt的內容依序存入lista列表中
lista =[]
old = open(rf'c:\tmp\{old_excel}.txt', encoding='utf-8',)   
for line in old.readlines():
    #print(line)
    lista.append(line.replace('\n',''))#將每筆資料存入lista中
#移除不需要的欄位
lista.remove("任務名稱;任務季度;漏洞類型;弱點描述;任務執行時間;IP;Hostname;系統管理者;OS版本;Port;Port Protocol;CVSS分數;風險等級;CVEs編號;處理步驟;是否可修補(Y/N);是否為新弱點(Y/N);完成確認日期")

# print("lista=",lista)

#將取得弱掃複掃.txt的內容依序存入listb列表中
listb =[]
new = open(rf'c:\tmp\{new_excel}.txt', encoding='utf-8')
for line in new.readlines():
    #print(line)
    listb.append(line.replace('\n','')) #將每筆資料存入listb中
#移除不需要的欄位
listb.remove("任務名稱;任務季度;漏洞類型;弱點描述;任務執行時間;IP;Hostname;系統管理者;OS版本;Port;Port Protocol;CVSS分數;風險等級;CVEs編號;處理步驟;是否可修補(Y/N);是否為新弱點(Y/N);完成確認日期")
# print("listb=",listb)

# #======================================取得未修補弱點的名稱與IP============================
#將list轉換為集合，初掃(old)-複掃(new)會取得未修補的漏洞資料
x_old = set(lista)
y_old = set(listb)
z_old= x_old - y_old
# print("x1 - y1=\n",z1)

#再將結果由集合轉為list，以便產生有序的列表
listc_old = list(z_old)
# print("listc_old===\n",listc_old)
# print(listc_old[0])
# print(listc_old[1])
# print(listc_old[2])
list_count_old = len(listc_old) #取得未修漏洞筆數
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
    listc_name_old.append(v1[3]) #弱點描述
    listc_ip_old.append(v1[5]) #IP
    listc_hostname_old.append(v1[6]) #hostname
    listc_admin_old.append(v1[7]) #系統管理者
    listc_os_old.append(v1[8]) #作業系統版本
    listc_status_old.append(v1[14]) #處理步驟
    listc_result_old.append(v1[15]) #是否可修補(Y/N)
    listc_confirmdata_old.append(v1[17]) #完成確認日期

# print("listc_name_old=",listc_name_old)
# print("listc_ip_old",listc_ip_old)
# print("listc_admin_old",listc_admin_old)
# print("listc_os_old",listc_os_old)
# print("listc_status_old",listc_status_old) 
# print("listc_result_old",listc_result_old) 

#======================================取得新弱點的弱點名稱與IP============================
#轉換為集合取得新的弱點
x_new = set(lista)
y_new = set(listb)
z_new= y_new - x_new
# print("y-x=\n",z_new)
#將集合結果轉成list，並將弱點名稱與ip切割
listc_new = list(z_new)
list_count_new = len(listc_new)
listc_name_new = [] #弱點描述
listc_ip_new = [] #IP

#====test====
# v2 = listc_new[i].split(';')
# listc_name_new.append(v2[2])
# print(listc_name_new)
# listc_ip_new.append(v2[4])
# print(listc_ip_new)
#============

n = 0
for i in range(list_count_new):
    v2 = listc_new[i].split(';')
    print(len(v2))
    listc_name_new.append(v2[3])
    listc_ip_new.append(v2[5])
    # breakpoint()
print("listc_name=",listc_name_new)
print("listc_ip",listc_ip_new)

# #=====================================進行資料修改=========================================
#讀取原始excel檔案內容
df_old = pd.read_excel(rf'c:\tmp\{old_excel}.xlsx', index_col=False)
df_new = pd.read_excel(rf'c:\tmp\{new_excel}.xlsx', index_col=False)

#將上次弱點修復結果的處理步驟、是否可修補(Y/N)欄位填入本次弱點掃描又掃到的項目中
for i in range(list_count_old):
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), '系統管理者'] = listc_admin_old[i]
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), 'Hostname'] = listc_hostname_old[i]
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), 'OS版本'] = listc_os_old[i]
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), '是否可修補(Y/N)'] = listc_result_old[i]
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), '處理步驟'] = listc_status_old[i]
    df_new.loc[((df_new['弱點描述'] == listc_name_old[i]) & (df_new['IP'] == listc_ip_old[i])), '完成確認日期'] = listc_confirmdata_old[i]  
#print(df_new)

#將新弱點標示為Y，運行的前提: 當df_new裡和df_old有相同的筆的資料時(代表弱點沒有被解決所以又出現於df_new的掃描結果中)，
#必須將df_old的"處理步驟"和""是否可修補(Y/N)"的內容填入df_new對應的欄位中，否則會因為資料不相同而被視為新的弱點("z2= y2 - x2"，會過濾出新的弱點的描述與IP)
for i in range(list_count_new):
    # df_new.loc[~df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'N'
    # df_new.loc[((df_new['弱點描述'] == listc_name_new[i]) & (df_new['IP'] == listc_ip_new[i])), '是否為新弱點(Y/N)'] = 'Y'
    df_new.loc[df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'Y'
    df_new.loc[~df_new['處理步驟'].isnull(), '是否為新弱點(Y/N)'] = 'N'
print(df_new)

#將經過篩選比對的high分頁內容寫入弱掃初掃與複掃合併結果清單，而Medium、Low分頁的資料則是直接從複掃清單複製而未經過過濾比對
result_file = pd.ExcelWriter(rf"c:\tmp\{today}-弱掃初掃與複掃合併結果清單.xlsx", engine='openpyxl')
df_new.to_excel(result_file, sheet_name= f'High-level-ALL', index=False)

df_new_Medium = pd.read_excel(rf'c:\tmp\{new_excel}.xlsx', sheet_name='Medium-level')
df_new_Medium.to_excel(result_file, sheet_name= f'Medium-level', index=False)

df_new_Low = pd.read_excel(rf'c:\tmp\{new_excel}.xlsx', sheet_name='Low-level')
df_new_Low.to_excel(result_file, sheet_name= f'Low-level', index=False)

result_file.save()