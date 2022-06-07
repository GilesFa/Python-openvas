from datetime import datetime
from pydoc import doc
from re import M
from elasticsearch import Elasticsearch
'''
必須存在已經過merge_excel.py匯出的excel檔案，並且檔名為202X QX 弱掃結果清單.xlsx
然後存放於C:\tmp\目錄下
'''
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

#設定來源excel檔案路徑
excelFilePath = rf'C:\tmp\{filename}.xlsx'

#設定elasticsearch所需資訊
# index_name = f"openvas-{today_date_all}"
index_name = f"openvas-" + f"{year}" + "-" + f"{month}"
es_user = "elastic"
es_pwd = "umec@123"
es_url = "http://10.0.99.100:9200"

#1. 建立連線=====================================================================
es = Elasticsearch(f"{es_url}", http_auth=(f'{es_user}', f'{es_pwd}'))

#2. 建立index====================================================================
index_body = {
    "settings": {
        "index": { "number_of_shards": 1,  "number_of_replicas": 1 }
    },
    "mappings": {
        "properties": {
            "ip" : {"type" : "ip"},
            "ip" : {"type" : "keyword"},
            "任務季度" : {"type" : "text"},
            "任務季度" : {"type" : "keyword"},
            "任務名稱" : {"type" : "text"},
            "任務名稱" : {"type" : "keyword"},
            "漏洞類型" : {"type" : "text"},
            "漏洞類型" : {"type" : "keyword"},
            "弱點描述" : {"type" : "text"},
            "弱點描述" : {"type" : "keyword"},
            "任務執行時間" : {"type" : "date"},
            "Hostname" : {"type" : "text"},
            "Hostname" : {"type" : "keyword"},
            "系統管理者" : {"type" : "text"},
            "系統管理者" : {"type" : "keyword"},
            "OS版本" : {"type" : "text"},
            "OS版本" : {"type" : "keyword"},
            "Port" : {"type" : "text"},
            "Port" : {"type" : "keyword"},
            "Port Protocol" : {"type" : "text"},
            "Port Protocol" : {"type" : "keyword"},
            "CVSS分數" : {"type" : "integer"},
            "CVSS分數" : {"type" : "keyword"},
            "風險等級" : {"type" : "text"},
            "風險等級" : {"type" : "keyword"},
            "CVEs編號" : {"type" : "text"},
            "CVEs編號" : {"type" : "keyword"},
            "弱點解決方法" : {"type" : "text"},
            "弱點解決方法" : {"type" : "keyword"},
            "處理步驟" : {"type" : "text"},
            "處理步驟" : {"type" : "keyword"},
            "是否可修補(Y/N)" : {"type" : "text"},
            "是否可修補(Y/N)" : {"type" : "keyword"},
            "是否為新弱點(Y/N)" : {"type" : "text"},
            "是否為新弱點(Y/N)" : {"type" : "keyword"},
            "完成確認日期" : {"type" : "date"},
            "完成確認日期" : {"type" : "keyword"}
        }
  }
}
print("index_name",index_name)

try : 
    result = es.indices.create(index=f'{index_name}', body=index_body)
except Exception as error:
    print("警告index已有建立:",error)

# 3.將excel轉為csv檔=============================================================
import pandas as pd

sheet = ['High-level-ALL','Medium-level','Low-level']
High_csv = rf'C:\tmp\{filename}_High-level.csv'
Medium_csv = rf'C:\tmp\{filename}_Medium-level.csv'
Low_csv = rf'C:\tmp\{filename}_Low-level.csv'

for i in sheet:
    read_file = pd.read_excel (f"{excelFilePath}", f'{i}')
    if i == "High-level-ALL":
        read_file.to_csv (f"{High_csv}", 
            index = None,
            header=True,
            encoding='utf_8_sig')
        high_rows = read_file.shape[0] #計算high風險的筆數
    elif i == "Medium-level":
        read_file.to_csv (f"{Medium_csv}", 
            index = None,
            header=True,
            encoding='utf_8_sig')
        Medium_rows = read_file.shape[0] #計算Medium風險的筆數
    else:
        read_file.to_csv (f"{Low_csv}", 
            index = None,
            header=True,
            encoding='utf_8_sig')
        Low_rows = read_file.shape[0] #計算Low風險的筆數

# 4.讀取csv檔案並轉換成特定json格式================================================
import csv 
import json 
import collections
from collections import OrderedDict
orderedDict = collections.OrderedDict()

csvFilePath = [High_csv, Medium_csv, Low_csv]
jsonFilePath = [rf'C:\tmp\{filename}_High.json' ,
                rf'C:\tmp\{filename}_Medium.json',
                rf'C:\tmp\{filename}_Low.json' ]

def csv_to_json(csvFilePath, jsonFilePath, doc_id):
    # global doc_id
    with open(csvFilePath, 'r',  encoding='utf-8') as csvf: 
        with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
            csvReader = csv.DictReader(csvf) 
            #  #為了讓每筆資料不會重複新增，因此必須加入"_id"來指定index內置id編號
            for row in csvReader: 
                x = OrderedDict([('index', {"_index": f"{index_name}", "_id": doc_id})])      
                jsonString = json.dumps(x)  
                doc_id += 1
                # print(row)
                jsonf.write(jsonString)
                jsonf.write("\n")
                y = json.dumps(row, ensure_ascii=False)
                # print(y)
                jsonf.write(y)
                jsonf.write("\n")

#依據風險等級計算筆數的起始與結束編號: high -> Medium -> Low
total_rows = high_rows + Medium_rows + Low_rows
high_rows_start_id = 1
high_rows_end_id = high_rows
Medium_rows_start_id = high_rows_end_id + 1
Medium_rows_end_id = Medium_rows_start_id + Medium_rows -1
Low_rows_start_id = Medium_rows_end_id + 1
Low_rows_end_id = Low_rows_start_id + Low_rows -1

# 依據風險等級將csv轉換成對應的json，然後讀取json檔案並批量將資料insert到elasticsearch
for i in csvFilePath:
    if "High" in i:
        csv_to_json(i, jsonFilePath[0], high_rows_start_id)
        with open(f'{jsonFilePath[0]}', 'r', encoding='utf-8') as json_file_high:
            documents = json_file_high.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')

    elif "Medium" in i:
        csv_to_json(i, jsonFilePath[1], Medium_rows_start_id)
        with open(f'{jsonFilePath[1]}', 'r', encoding='utf-8') as json_file_Medium:
            documents = json_file_Medium.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')

    elif "Low" in i:
        csv_to_json(i, jsonFilePath[2], Low_rows_start_id)
        with open(f'{jsonFilePath[2]}', 'r', encoding='utf-8') as json_file_Low:
            documents = json_file_Low.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')

print(f"高風險id從{high_rows_start_id}開始，總共{high_rows}筆，最後一碼id為{high_rows_end_id}")
print(f"中風險id從{Medium_rows_start_id}開始，總共{Medium_rows}筆，最後一碼id為{Medium_rows_end_id}")
print(f"低風險id從{Low_rows_start_id}開始，總共{Low_rows}筆，最後一碼id為{Low_rows_end_id}")
print(f"總弱點數量為{total_rows}")