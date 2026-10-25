import zipfile
import json
import re

with open('config/plantillas.json', 'r', encoding='utf-8') as f:
    plantillas = json.load(f)

for p in plantillas:
    cod = p['codigo']
    if 'POS' in cod:
        path = 'plantillas/' + p['archivo_plantilla']
        with zipfile.ZipFile(path, 'r') as z:
            xml = z.read('word/document.xml').decode('utf-8')
            for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL):
                t_str = tbl.group(0)
                txt = ' '.join(re.sub(r'<[^>]+>', ' ', t_str).split())
                if 'FOLIO' in txt.upper() and not 'EMPAQUE' in txt.upper():
                    print('**************************************************')
                    print(f'{cod} ({p["archivo_plantilla"]})')
                    print('TEXT:', txt)
                    print('FULL XML:')
                    print(t_str)
