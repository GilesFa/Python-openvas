from datetime import datetime
from elasticsearch import Elasticsearch
from collections import OrderedDict
import json
import csv 
import pandas as pd
import collections

#1.變數設定
#抓取西元年-月-日
today = str(datetime.today())
today_date = today.split(' ')
today_date_all = today_date[0]

#設定elasticsearch所需資訊
index_name = f"demo-{today_date_all}"
es_user = "elastic"
es_pwd = "password"
es_url = "http://192.168.0.10:9200"

#1. 建立連線
#1.1 舊版連線方式:http_auth
# es = Elasticsearch(es_url, http_auth=(f'{es_user}', f'{es_pwd}'))
#1.2 新版連線方式:basic_auth
es = Elasticsearch(es_url, basic_auth=(es_user, es_pwd))

#1.3 印出當前elasticsearch叢集資訊
# print(es.info())
#將資訊另存成json檔案
js = dict(es.info())
with open("esinfo.json", "w") as jsonfile:
    json.dump(js, jsonfile)

#2. 建立index
index_body = {
    "settings": {
        "index": { "number_of_shards": 1,  "number_of_replicas": 1 }
    },
    "mappings": {
        "properties": {
            "ip" : {"type" : "ip"},
            "ip" : {"type" : "keyword"},
            "漏洞類型" : {"type" : "text"},
            "漏洞類型" : {"type" : "keyword"},
            "弱點描述" : {"type" : "text"},
            "弱點描述" : {"type" : "keyword"},
            "Hostname" : {"type" : "text"},
            "Hostname" : {"type" : "keyword"},
        }
  }
}

try : 
    result = es.indices.create(index=f'{index_name}', body=index_body)
except Exception as error:
    pass
    print(error)

#3. excel轉為csv檔
read_file = pd.read_excel(r"c:\temp\demo.xlsx")
read_file.to_csv (r'c:\temp\demo_csv', 
            index = None,
            header=True,
            encoding='utf_8_sig')

#4. csv檔案轉換成特定json格式
orderedDict = collections.OrderedDict()

def csv_to_json(csvFilePath, jsonFilePath):
    with open(csvFilePath, 'r',  encoding='utf-8') as csvf: 
        with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
            csvReader = csv.DictReader(csvf) 
            i =1
            for row in csvReader: 
                #為了讓每筆資料不會重複新增，因此必須加入"_id"來指定index內置id編號
                x = OrderedDict([('index', {"_index": f"{index_name}", "_id": i})])    
                i +=1   
                jsonString = json.dumps(x)  
                print(row)
                jsonf.write(jsonString)
                jsonf.write("\n")
                y = json.dumps(row, ensure_ascii=False)
                print(y)
                jsonf.write(y)
                jsonf.write("\n")

#5. 讀取json檔案並批量將資料insert到elasticsearch
csv_to_json(r'c:\temp\demo_csv', r'c:\temp\demo_json')

with open(r'c:\temp\demo_json', 'r', encoding='utf-8') as jsonfile:
    documents = jsonfile.readlines()
result = es.bulk(body=documents, index=f'{index_name}')
