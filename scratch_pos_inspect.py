import zipfile
import json
import re

with open('config/plantillas.json', 'r', encoding='utf-8') as f:
    plantillas = json.load(f)

for p in plantillas:
    if 'POS' in p['codigo']:
        print('-------------------------------------------')
        print(p['codigo'], ':', p['archivo_plantilla'])
        path = 'plantillas/' + p['archivo_plantilla']
        with zipfile.ZipFile(path, 'r') as z:
            xml = z.read('word/document.xml').decode('utf-8')
            # Buscar tablas o párrafos que contengan FOLIO o 340 o FECHA en los primeros 10000 caracteres
            for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL):
                tbl_xml = tbl.group(0)
                txt = re.sub(r'<[^>]+>', ' ', tbl_xml)
                txt = ' '.join(txt.split())
                if 'FOLIO' in txt.upper() or '340' in txt:
                    print('TABLA ENCONTRADA:')
                    print(txt[:300])
                    print('TABLA XML:')
                    print(tbl_xml)
                    print('===========================================')
