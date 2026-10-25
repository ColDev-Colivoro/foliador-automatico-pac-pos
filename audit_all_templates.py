import zipfile
import json
import re
import os
import sys

# Forzar utf-8 en stdout
sys.stdout.reconfigure(encoding='utf-8')

with open('config/plantillas.json', 'r', encoding='utf-8') as f:
    plantillas = json.load(f)

print(f"Total plantillas a auditar: {len(plantillas)}\n")

for p in plantillas:
    cod = p['codigo']
    fn = p['archivo_plantilla']
    path = os.path.join('plantillas', fn)
    print("=" * 60)
    print(f"AUDITORIA: {cod} ({fn})")
    print("=" * 60)
    if not os.path.exists(path):
        print(f"ERROR: Archivo no existe: {path}")
        continue

    with zipfile.ZipFile(path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')

    # Buscar tablas con folio o flotantes de encabezado
    tables_found = []
    for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL):
        tbl_xml = tbl.group(0)
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', tbl_xml).split())
        # ignorar la tabla grande de datos si no es la del folio
        if any(k in txt.upper() for k in ['FOLIO', '340', '000104']) and not 'EMPAQUE' in txt.upper():
            tables_found.append((txt[:200], tbl_xml))

    # Buscar párrafos con FOLIO
    paragraphs_found = []
    for m in re.finditer(r'<w:p(?:\s+[^>]*)?>.*?</w:p>', xml, re.DOTALL):
        p_xml = m.group(0)
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', p_xml).split())
        if 'FOLIO' in txt.upper() and not 'EMPAQUE' in txt.upper():
            para_id = re.findall(r'paraId="([^"]+)"', p_xml)
            paragraphs_found.append((txt, para_id, p_xml))

    print(f"Tablas de folio encontradas: {len(tables_found)}")
    for txt, t_xml in tables_found:
        print(f"  -> TABLA: '{txt}'")
        tblp = re.findall(r'<w:tblpPr[^>]*?>', t_xml)
        if tblp:
            print(f"     tblpPr: {tblp[0]}")
        cols = re.findall(r'<w:gridCol[^>]*?>', t_xml)
        print(f"     Columnas: {len(cols)} {cols}")

    print(f"Parrafos con 'FOLIO': {len(paragraphs_found)}")
    for txt, pid, p_xml in paragraphs_found:
        print(f"  -> PARRAFO: '{txt}' | paraId: {pid}")

print("\nFIN AUDITORIA.")
