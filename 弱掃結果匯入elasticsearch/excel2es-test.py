from datetime import datetime
from elasticsearch import Elasticsearch

#1. 建立連線
es = Elasticsearch("http://192.168.0.10:9200", http_auth=('elastic', 'password'))

#2. 建立index
index_body = {
    "settings": {
        "index": { "number_of_shards": 1,  "number_of_replicas": 1 }
    },
    "mappings": {
        "properties": {
            "close" : {"type" : "float"},
            "date" : {"type" : "date"},
            "high" : {"type" : "float"},
            "low" : {"type" : "float"},
            "open" : {"type" : "float"},
            "stock_id" : {"type" : "keyword"},
            "volume" : {"type" : "integer"}
        }
  }
}

result = es.indices.create(index='python', body=index_body)

#3. 批量新增資料
documents = [
    {"index":{"_id" : "0003"}},
    { "stock_id":"0050", "date":"2020-09-11", "volume":2905291,"open":103.20,"high":105.35, "low":103.80, "close":104.25 },
    {"index":{"_id" : "0004"}},
    { "stock_id":"0050", "date":"2020-09-12", "volume":2232343, "open":104.20, "high":105.35, "low":102.80, "close":104.00 },
]

result = es.bulk(body=documents, index='python')