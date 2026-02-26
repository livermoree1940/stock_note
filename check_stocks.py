import json
import os

if os.path.exists('data/stock_notes.json'):
    with open('data/stock_notes.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
        print('JSON文件中的股票代码:', list(data.keys()))
        print('600519是否存在:', '600519' in data)
        if '600519' in data:
            print('600519的详细信息:', data['600519'])
else:
    print('JSON文件不存在')
