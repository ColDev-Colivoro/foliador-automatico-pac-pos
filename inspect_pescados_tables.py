import zipfile
import re
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

pescados_dir = os.path.abspath(r"../Programa Aseguramiento de calidad/Docuemntos Pos paC prerrequisitos/PAC VERSION 15/registros pac pescados")

for fn in ["Pc2 Almacenamiento Pescados.docx", "pc3 Temperatura en Linea de proceso.docx", "Pc7 Control de Agua y Glaseo pescados.docx"]:
    path = os.path.join(pescados_dir, fn)
    print("=" * 60)
    print("ARCHIVO:", fn)
    with zipfile.ZipFile(path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL):
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', tbl.group(0)).split())
        if 'FOLIO' in txt.upper() or 'FECHA' in txt.upper():
            print("TBL:", tbl.group(0)[:600])
            print("TXT:", txt)
            print("-----------------------------------")
