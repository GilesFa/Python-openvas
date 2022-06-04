import pandas as pd
import os
import glob
import numpy as np
import os
from datetime import date, datetime

'''
此程式的目標功能:
1. 利用scripts掃描伺服器網段(nmap -sL -i serverIP.txt  | grep "(" | awk 'BEGIN{FS=" "} {print $6 $5}')，並探測解析主機名稱、作業系統版本()nmap -O 10.0.5.35 |grep "OS details" |awk 'BEGIN{FS=":"} {print $2}'
2. 將OpenVAS各網段弱掃清單合併與過濾特定欄位成EXCEL檔
3. 將第一步驟收集到的資訊，以IP作為mearge的依據，將主機名稱與OS版本帶入
4. 最後依據最後版本的excel檔，將其資料寫入mysql資料庫中，後續復掃的結果從資料庫中提取來比對
'''
#================================變數設定==================================================
#抓取當前日期與季
today = str(datetime.today())
today_date = today.split(' ')
today_date_all = today_date[0]
print(today_date_all)

today_date = today_date_all.split('-')
year = today_date[0]
month = today_date[1]
day = today_date[2]

if int(month) >= 1 and int(month)<= 3:
    quarter = "Q1"
elif int(month) >= 4 and int(month)<= 6:
    quarter = "Q2"
elif int(month) >= 7 and int(month)<= 9:
    quarter = "Q3"
elif int(month) >= 10 and int(month)<= 12:
    quarter = "Q4"

#設定檔案名稱格式為:2022 Q2 弱掃初掃結果清單
yearquarter = year + " " +quarter
filename = f"{yearquarter}" + " " + "弱掃結果清單"

#建立tmp目錄
try:
    os.mkdir(r"c:\tmp")
except Exception as error:
    print("c:\tmp目錄已存在")

#請設定讀取路徑，取得當前目錄下的特定檔案名稱，並且以list的方式存入變數file_list
# file_list = (glob.glob(r'W:\1100\1120\資訊安全\資安防範措施\系統弱點掃描\弱點掃描\2022\上半年\初掃\*.csv'))
file_list = (glob.glob(r'C:\tmp\*.csv'))

print(file_list)

appended_data = []
for item in file_list:
  #取得當前絕對路徑 + file_list的值 = 當前路徑下的檔案完整絕對路徑
  file_path = os.path.abspath(f'{item}')
  #讀取內容並存入變數data
  data = pd.read_csv(rf"{file_path}")
  #將變數data的內容依序存入appended_data列表中
  appended_data.append(data)

print(appended_data[0])

# pd.concat([df1,df2,df3])，將多個列表內容合併，axis=0代表以x軸方向合併，
# ignore_index=True代表忽略原本的index，重新由0開始編號
result = pd.concat(appended_data, axis=0, ignore_index=True)

#篩選需要的欄位
#result = result[['IP','Hostname','Port','Port Protocol','CVSS','Severity','CVEs','Task Name','Timestamp','Impact','Solution','Affected Software/OS']]
result = result[['Task Name','NVT Name', 'Timestamp','IP','Hostname','Port','Port Protocol','CVSS','Severity','CVEs','Solution']]

#變更欄位名稱
result.columns = ['任務名稱','弱點描述', '任務執行時間','IP','Hostname','Port','Port Protocol','CVSS分數','風險等級','CVEs編號','弱點解決方法']

###################欄位新增############################
#新增欄位於最後的欄位，並且value為NaN
result["處理步驟"] = np.nan
result["是否可修補(Y/N)"] = np.nan
result["是否為新弱點(Y/N)"] = "N"
result["完成確認日期"] = np.nan #此index為16
#result["OS版本"] = np.nan，然後插入到特定順序欄位的後面，index從0開始
result.insert(5, "系統管理者", np.nan)
result.insert(6, "OS版本", np.nan)
result.insert(1, "漏洞類型", "未定義")
#result.insert(1, "漏洞類型", np.nan)
###################欄位新增############################

#調整欄位順序
#rresult = result[['任務名稱','任務執行時間','IP','Hostname',"OS版本",'Port','Port Protocol','CVSS分數','風險等級','CVEs編號','弱點說明','弱點解決方法','影響範圍(軟體/系統)','處理步驟']]

#建立excel存放路徑，以便存放後面分頁的資料
# result_file = pd.ExcelWriter(rf"W:\1100\1120\資訊安全\資安防範措施\系統弱點掃描\弱點掃描\python\{outputfile}.xlsx", engine='openpyxl')
result_file = pd.ExcelWriter(rf"c:\tmp\{filename}.xlsx", engine='openpyxl')

#過濾出Severity欄位為High & 弱點名稱不是"Report outdated / end-of-life Scan Engine / Environment (local)"
    #範例1: df[df[“column_name”] == value]，多字段篩選 : df[(df[“column_name1”] <= value) & (df[“column_name2”] == str)]
        #result_filter_High = result[result["弱點名稱"] != "Report outdated / end-of-life Scan Engine / Environment (local)"]

    #範例2：df[df[“column_name”].isin(li)] (# li = [20, 25, 27] 或 li = np.arange(20, 30))　
        #result_filter_High = result[result['風險等級'].isin(["High"])]

    #範例3 : 複合式，範例1與範例2搭配使用

#過濾出風險為High的資料    
result_filter_High = result[result['風險等級'].isin(["High"])]

#當弱點描述欄位的value不包含(EOL、End Of Life)且同時包含其他內容時，修改漏洞類型欄位的值為特定名稱
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("HTTP"))) , '漏洞類型'] = 'HTTP'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("Tomcat"))) , '漏洞類型'] = 'Tomcat'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("PHP"))) , '漏洞類型'] = 'PHP'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("IIS"))) , '漏洞類型'] = 'IIS'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("Remote Desktop"))) , '漏洞類型'] = 'RDP'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("Oracle"))) , '漏洞類型'] = 'Oracle'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & ((result_filter_High['弱點描述'].str.contains("Password")) | (result_filter_High['弱點描述'].str.contains("password")))) , '漏洞類型'] = 'password'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("SMB"))) , '漏洞類型'] = 'SMB'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("iLO"))) , '漏洞類型'] = 'HP iLO'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("VMware"))) , '漏洞類型'] = 'VMware'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("Avaya"))) , '漏洞類型'] = 'Avaya'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("MongoDB"))) , '漏洞類型'] = 'MongoDB'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("MariaDB"))) , '漏洞類型'] = 'MariaDB'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("Brute Force"))) , '漏洞類型'] = 'Brute Force'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("OpenSSL"))) , '漏洞類型'] = 'OpenSSL'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & ((result_filter_High['弱點描述'].str.contains("rexec")) | (result_filter_High['弱點描述'].str.contains("rsh")))) , '漏洞類型'] = '遠端執行'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("rlogin"))) , '漏洞類型'] = '遠端登入'
result_filter_High.loc[(((~result_filter_High['弱點描述'].str.contains("EOL")) | (~result_filter_High['弱點描述'].str.contains("End Of Life"))) & (result_filter_High['弱點描述'].str.contains("XAMPP"))) , '漏洞類型'] = 'web'
#當某個欄位的value包含特定value時，修改另一個欄位的值
result_filter_High.loc[((result_filter_High['弱點描述'].str.contains("EOL")) | (result_filter_High['弱點描述'].str.contains("End Of Life"))), '漏洞類型'] = 'EOL'


#取得總rows與columns
total_all_high = result_filter_High.shape
#取得總rows筆數
total_all_high_rows = total_all_high[0]
#寫入excel分頁，index=Flase代表不添加顯示筆數順序的欄位
result_filter_High.to_excel(result_file, sheet_name= f'High-level-ALL', index=False)

#====================================================================================================================
#過濾出Severity欄位為High & 弱點名稱包含"SMB""
    #範例4:contains函數， df[df[“column_name”].str.contains(“str”)]
# result_filter_High_SMB = result_filter_High[result_filter_High["弱點描述"].str.contains("SMB")]
# #取得總rows與columns
# total_SMB_high = result_filter_High_SMB.shape
# #取得總rows筆數
# total_SMB_high_rows = total_SMB_high[0]
# result_filter_High_SMB.to_excel(result_file, sheet_name= f'High-level-SMB({total_SMB_high_rows})')

# #過濾出Severity欄位為High & 弱點名稱包含"HTTP""
# result_filter_High_HTTP = result_filter_High[result_filter_High["弱點描述"].str.contains("HTTP")]
# #取得總rows與columns
# total_HTTP_high = result_filter_High_HTTP.shape
# #取得總rows筆數
# total_HTTP_high_rows = total_HTTP_high[0]
# result_filter_High_HTTP.to_excel(result_file, sheet_name= f'High-level-HTTP({total_HTTP_high_rows})')

# #過濾出Severity欄位為High & 弱點名稱包含"Tomcat""
# result_filter_High_Tomcat = result_filter_High[result_filter_High["弱點描述"].str.contains("Tomcat")]
# #取得總rows與columns
# total_Tomcat_high = result_filter_High_Tomcat.shape
# #取得總rows筆數
# total_Tomcat_high_rows = total_Tomcat_high[0]
# result_filter_High_Tomcat.to_excel(result_file, sheet_name= f'High-level-Tomcat({total_Tomcat_high_rows})')

# #過濾出Severity欄位為High & 弱點名稱包含"DB""
# result_filter_High_DB = result_filter_High[result_filter_High["弱點描述"].str.contains("DB")]
# #取得總rows與columns
# total_DB_high = result_filter_High_DB.shape
# #取得總rows筆數
# total_DB_high_rows = total_DB_high[0]
# result_filter_High_DB.to_excel(result_file, sheet_name= f'High-level-DB({total_DB_high_rows})')

#====================================================================================================================

#過濾出Severity為Medium
result_filter_Medium = result[result['風險等級'].isin(["Medium"])]
#當某個欄位的value包含特定value時，修改另一個欄位的值
result_filter_Medium.loc[result_filter_Medium['弱點描述'].str.contains("EOL"), '漏洞類型'] = 'EOL'
result_filter_Medium.loc[result_filter_Medium['弱點描述'].str.contains("End Of Life"), '漏洞類型'] = 'End Of Life'
#當弱點描述欄位的value不包含(EOL、End Of Life)且同時包含其他內容時，修改漏洞類型欄位的值為特定名稱
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("HTTP"))) , '漏洞類型'] = 'HTTP'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("Tomcat"))) , '漏洞類型'] = 'Tomcat'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("PHP"))) , '漏洞類型'] = 'PHP'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("IIS"))) , '漏洞類型'] = 'IIS'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("Remote Desktop"))) , '漏洞類型'] = 'RDP'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("Oracle"))) , '漏洞類型'] = 'Oracle'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & ((result_filter_Medium['弱點描述'].str.contains("Password")) | (result_filter_Medium['弱點描述'].str.contains("password")))) , '漏洞類型'] = 'password'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("SMB"))) , '漏洞類型'] = 'SMB'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("iLO"))) , '漏洞類型'] = 'HP iLO'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("VMware"))) , '漏洞類型'] = 'VMware'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("Avaya"))) , '漏洞類型'] = 'Avaya'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("MongoDB"))) , '漏洞類型'] = 'MongoDB'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("MariaDB"))) , '漏洞類型'] = 'MariaDB'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("Brute Force"))) , '漏洞類型'] = 'Brute Force'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("OpenSSL"))) , '漏洞類型'] = 'OpenSSL'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & ((result_filter_Medium['弱點描述'].str.contains("rexec")) | (result_filter_Medium['弱點描述'].str.contains("rsh")))) , '漏洞類型'] = '遠端執行'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("rlogin"))) , '漏洞類型'] = '遠端登入'
result_filter_Medium.loc[(((~result_filter_Medium['弱點描述'].str.contains("EOL")) | (~result_filter_Medium['弱點描述'].str.contains("End Of Life"))) & (result_filter_Medium['弱點描述'].str.contains("XAMPP"))) , '漏洞類型'] = 'web'

#取得總rows與columns
total_all_Medium = result_filter_Medium.shape
#取得總rows筆數
total_all_Medium_rows = total_all_Medium[0]
#寫入excel分頁
result_filter_Medium.to_excel(result_file, sheet_name= f'Medium-level')

#=========================================================================================================================


#過濾出Severity為Low
result_filter_Low = result[result['風險等級'].isin(["Low"])]
#當某個欄位的value包含特定value時，修改另一個欄位的值
result_filter_Low.loc[result_filter_Low['弱點描述'].str.contains("EOL"), '漏洞類型'] = 'EOL'
result_filter_Low.loc[result_filter_Low['弱點描述'].str.contains("End Of Life"), '漏洞類型'] = 'End Of Life'
#當弱點描述欄位的value不包含(EOL、End Of Life)且同時包含其他內容時，修改漏洞類型欄位的值為特定名稱
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("HTTP"))) , '漏洞類型'] = 'HTTP'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("Tomcat"))) , '漏洞類型'] = 'Tomcat'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("PHP"))) , '漏洞類型'] = 'PHP'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("IIS"))) , '漏洞類型'] = 'IIS'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("Remote Desktop"))) , '漏洞類型'] = 'RDP'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("Oracle"))) , '漏洞類型'] = 'Oracle'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & ((result_filter_Low['弱點描述'].str.contains("Password")) | (result_filter_Low['弱點描述'].str.contains("password")))) , '漏洞類型'] = 'password'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("SMB"))) , '漏洞類型'] = 'SMB'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("iLO"))) , '漏洞類型'] = 'HP iLO'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("VMware"))) , '漏洞類型'] = 'VMware'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("Avaya"))) , '漏洞類型'] = 'Avaya'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("MongoDB"))) , '漏洞類型'] = 'MongoDB'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("MariaDB"))) , '漏洞類型'] = 'MariaDB'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("Brute Force"))) , '漏洞類型'] = 'Brute Force'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("OpenSSL"))) , '漏洞類型'] = 'OpenSSL'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & ((result_filter_Low['弱點描述'].str.contains("rexec")) | (result_filter_Low['弱點描述'].str.contains("rsh")))) , '漏洞類型'] = '遠端執行'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("rlogin"))) , '漏洞類型'] = '遠端登入'
result_filter_Low.loc[(((~result_filter_Low['弱點描述'].str.contains("EOL")) | (~result_filter_Low['弱點描述'].str.contains("End Of Life"))) & (result_filter_Low['弱點描述'].str.contains("XAMPP"))) , '漏洞類型'] = 'web'

#取得總rows與columns
total_all_Low = result_filter_Low.shape
#取得總rows筆數
total_all_Low_rows = total_all_Low[0]
#寫入excel分頁
result_filter_Low.to_excel(result_file, sheet_name= f'Low-level')

#=========================================================================================================================

#將所有分頁保存成excel檔
result_file.save()

print("資料合併與過濾完成!")
# print(rf"W:\1100\1120\資訊安全\資安防範措施\系統弱點掃描\弱點掃描\python\{outputfile}.xlsx")
