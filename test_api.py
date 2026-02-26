import requests
import json

# 测试保存股票数据
url = 'http://localhost:5000/api/notes'
data = {
    'code': '600000',
    'name': '浦发银行',
    'concept': '银行',
    'industry': '银行III',
    'notes': '测试股票',
    'timestamp': '2026-02-24T15:00:00Z'
}

response = requests.post(url, json=data)
print('保存响应:', response.status_code, response.json())

# 测试获取股票数据
response = requests.get(url)
print('获取响应:', response.status_code)
print('股票列表:', list(response.json().keys()))
