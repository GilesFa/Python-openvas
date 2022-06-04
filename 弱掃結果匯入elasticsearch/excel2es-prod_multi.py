from datetime import datetime
from pydoc import doc
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
index_name = f"openvas-{today_date_all}"
es_user = "elastic"
es_pwd = "password"
es_url = "http://192.168.0.10:9200"

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
            "完成確認日期" : {"type" : "text"},
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
    elif i == "Medium-level":
        read_file.to_csv (f"{Medium_csv}", 
            index = None,
            header=True,
            encoding='utf_8_sig')
    else:
        read_file.to_csv (f"{Low_csv}", 
            index = None,
            header=True,
            encoding='utf_8_sig')

# 4.讀取csv檔案並轉換成特定json格式================================================
import csv 
import json 
import collections
orderedDict = collections.OrderedDict()
from collections import OrderedDict

csvFilePath = [High_csv, Medium_csv, Low_csv]
jsonFilePath = [rf'C:\tmp\{filename}_High.json' ,
                rf'C:\tmp\{filename}_Medium.json',
                rf'C:\tmp\{filename}_Low.json' ]

def csv_to_json(csvFilePath, jsonFilePath):
    i = 1
    jsonArray = []
    with open(csvFilePath, 'r',  encoding='utf-8') as csvf: 
        with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
            csvReader = csv.DictReader(csvf) 
            for row in csvReader: 
                x = OrderedDict([('index', {"_index": f"{index_name}"})])      
                jsonString = json.dumps(x)  
                i += 1
                # print(row)
                jsonf.write(jsonString)
                jsonf.write("\n")
                y = json.dumps(row, ensure_ascii=False)
                # print(y)
                jsonf.write(y)
                jsonf.write("\n")

# 依據風險等級將csv轉換成對應的json，然後讀取json檔案並批量將資料insert到elasticsearch
for i in csvFilePath:
    if "High" in i:
        csv_to_json(i, jsonFilePath[0])
        with open(f'{jsonFilePath[0]}', 'r', encoding='utf-8') as json_file_high:
            documents = json_file_high.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')
    elif "Medium" in i:
        csv_to_json(i, jsonFilePath[1])
        with open(f'{jsonFilePath[1]}', 'r', encoding='utf-8') as json_file_Medium:
            documents = json_file_Medium.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')
    elif "Low" in i:
        csv_to_json(i, jsonFilePath[2])
        with open(f'{jsonFilePath[2]}', 'r', encoding='utf-8') as json_file_Low:
            documents = json_file_Low.readlines() #將會一行一行進行讀取
        result = es.bulk(body=documents, index=f'{index_name}')