import urllib.request, json, sys
urls=[
    'http://127.0.0.1:5000/filtros',
    'http://127.0.0.1:5000/dados?nivel=linha',
    'http://127.0.0.1:5000/dados?nivel=linha&sku=1001',
    'http://127.0.0.1:5000/dados?nivel=linha&linha=Coco%20Ralado',
    'http://127.0.0.1:5000/dados?nivel=linha&produto=COCO%20RALADO%20SACHE%20CDV%20100G'
]
for url in urls:
    try:
        with urllib.request.urlopen(url, timeout=5) as r:
            data = json.load(r)
            print(url)
            if isinstance(data, dict):
                print('keys', sorted(data.keys()))
                if 'linhas' in data: print('linhas', data['linhas'][:5])
                if 'skus' in data: print('skus', data['skus'][:5])
                if 'rows' in data: print('rows', len(data['rows']))
            else:
                print(data)
    except Exception as e:
        print('ERROR', url, e, file=sys.stderr)
