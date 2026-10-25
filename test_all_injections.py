import os
import json
import zipfile
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

from core.folio_engine import FolioEngine

engine = FolioEngine()
salida_test = os.path.join(engine.salida_dir, "test_audit_all")
os.makedirs(salida_test, exist_ok=True)

with open('config/plantillas.json', 'r', encoding='utf-8') as f:
    plantillas = json.load(f)

print(f"Probando generacion para {len(plantillas)} plantillas...\n")

for p in plantillas:
    cod = p['codigo']
    fn = p['archivo_plantilla']
    src = os.path.join('plantillas', fn)
    dst = os.path.join(salida_test, f"{cod}_Folio_00555.docx")

    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    xml_str = data.decode("utf-8")
                    xml_str = engine._inyectar_folio_quirurgico(xml_str, cod, "00555")
                    data = xml_str.encode("utf-8")
                zout.writestr(item, data)

    # Ahora analizar el resultado generado
    with zipfile.ZipFile(dst, 'r') as z:
        gen_xml = z.read('word/document.xml').decode('utf-8')

    # Buscar todas las ocurrencias de FOLIO y números sospechosos
    print("=" * 60)
    print(f"RESULTADO GENERADO: {cod}")
    print("=" * 60)

    # Buscar párrafos con folio
    found_p = []
    for m in re.finditer(r'<w:p(?:\s+[^>]*)?>.*?</w:p>', gen_xml, re.DOTALL):
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', m.group(0)).split())
        if any(k in txt.upper() for k in ['FOLIO', '340', '000104', '00555']) and not 'EMPAQUE' in txt.upper():
            found_p.append(txt)

    # Buscar tablas flotantes
    found_tbl = []
    for m in re.finditer(r'<w:tbl>.*?</w:tbl>', gen_xml, re.DOTALL):
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', m.group(0)).split())
        if any(k in txt.upper() for k in ['FOLIO', '340', '000104', '00555']) and not 'EMPAQUE' in txt.upper():
            found_tbl.append(txt)

    print("Tablas con folio en documento generado:")
    for t in found_tbl:
        print("  ->", t)

    print("Párrafos con folio en documento generado:")
    for fp in found_p:
        print("  ->", fp)

    # Verificar si quedaron números viejos sin reemplazar
    numeros_viejos = [n for n in ['340', '000 10 4', '000104', '000 4 1 4', '000 38 5', '00 3 79', '1 30', '000 3 8', '003 7'] if n in gen_xml or n.replace(' ', '') in gen_xml.replace(' ', '')]
    if numeros_viejos:
        print(f"⚠️ ALERTA: Se detectaron números viejos residuales: {numeros_viejos}")
    else:
        print("✅ Limpio de números viejos.")
