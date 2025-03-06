import requests
import json

API_ENDPOINT = "https://ews-bcn.api.bosch.com/knowledge/insight-and-analytics/nlu/s/v2/openie/open-facts"
API_KEY = "<SUA-CHAVE-DE-API>"

proxies = {
    "http": "http://rb-proxy-de.bosch.com:8080",
    "https": "http://rb-proxy-de.bosch.com:8080",
}

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

data = {
    "document": {
        "doc_id": "ma-alibaba-investments",
        "language": "en",
        "title": "Ma founded Alibaba",
        "body": "Ma founded Alibaba in Hangzhou with $2M investment from SoftBank and Goldman."
    },
    "openie_config": {
        "keep_longest": False,
        "include_optional_adverbs": False,
        "reverb_relation_style": False,
        "process_possessives": False,
        "process_appositions": True,
        "process_partmods": False,
        "process_cc_non_verbs": True,
        "process_all_verbs": True,
        "lemmatize": False
    }
}

response = requests.post(API_ENDPOINT, headers=headers, data=json.dumps(data), proxies=proxies)

print(response.status_code, response.text)
