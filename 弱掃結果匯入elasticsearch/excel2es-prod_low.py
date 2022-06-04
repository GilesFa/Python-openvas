from datetime import datetime
from pydoc import doc
from elasticsearch import Elasticsearch

today = str(datetime.today())
today_date = today.split(' ')
today_date = today_date[0]

excelFilePath = r'C:\temp\2022 Q2 弱掃初掃結果清單.xlsx'
csvFilePath = r'C:\temp\2022 Q2 弱掃初掃結果清單_Low-level.csv'
jsonFilePath =r'C:\temp\2022 Q2 弱掃初掃結果清單_Low-level.json'
index_name = f"openvas-{today_date}"
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

try : 
    result = es.indices.create(index=f'{index_name}', body=index_body)
except :
    print("index已有建立")

# 3.將excel轉為csv檔=============================================================
import pandas as pd
  
read_file = pd.read_excel (f"{excelFilePath}", 'Low-level')

print(read_file)

read_file.to_csv (f"{csvFilePath}", 
                  index = None,
                  header=True,
                  encoding='utf_8_sig')
    

# 4.讀取csv檔案並轉換成特定json格式================================================
import csv 
import json 
import collections
orderedDict = collections.OrderedDict()
from collections import OrderedDict

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
                      
csv_to_json(csvFilePath, jsonFilePath)

#5. 讀取json檔案並批量將資料insert到elasticsearch
with open(f'{jsonFilePath}', 'r', encoding='utf-8') as json_file:
    documents = json_file.readlines() #將會一行一行進行讀取

result = es.bulk(body=documents, index=f'{index_name}')