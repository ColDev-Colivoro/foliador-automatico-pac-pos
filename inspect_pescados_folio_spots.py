import zipfile
import re
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

pescados_dir = os.path.abspath(r"../Programa Aseguramiento de calidad/Docuemntos Pos paC prerrequisitos/PAC VERSION 15/registros pac pescados")

files = [
    ("PAC_PCC1_PESC", "PCC1 Recepcion MP Pescados.docx"),
    ("PAC_PC2_PESC", "Pc2 Almacenamiento Pescados.docx"),
    ("PAC_PC3_PESC", "pc3 Temperatura en Linea de proceso.docx"),
    ("PAC_PC4_PESC", "pc4 Congelacion y enfriado pescados.docx"),
    ("PAC_PC5_PESC", "pcc5 Empaque y Reempaque pescados.docx"),
    ("PAC_PC6_PESC", "Pc 6 Despacho Pescados.docx"),
    ("PAC_PC7_PESC", "Pc7 Control de Agua y Glaseo pescados.docx")
]

for cod, fn in files:
    path = os.path.join(pescados_dir, fn)
    print("=" * 70)
    print(f"ANÁLISIS DE FOLIO: {cod} ({fn})")
    print("=" * 70)
    with zipfile.ZipFile(path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')

    # Buscar ocurrencias de FOLIO (case-insensitive) o números de muestra
    for m in re.finditer(r'<w:p(?:\s+[^>]*)?>.*?</w:p>', xml, re.DOTALL):
        p_xml = m.group(0)
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', p_xml).split())
        if any(k in txt.upper() for k in ['FOLIO', '00 0 1', '00 4', '003']) and not 'EMPAQUE' in txt.upper():
            pid = re.findall(r'paraId="([^"]+)"', p_xml)
            print(f"Párrafo [{pid}]: '{txt}'")
            print(f"XML: {p_xml}\n")

    # Si hay tablas de folio
    for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.DOTALL):
        t_xml = tbl.group(0)
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', t_xml).split())
        if any(k in txt.upper() for k in ['FOLIO', '00 0 1', '00 4', '003']) and not 'EMPAQUE' in txt.upper():
            props = re.findall(r'<w:tblpPr[^>]*?>', t_xml)
            if props:
                print(f"Tabla flotante: {props}")
                print(f"Texto tabla: '{txt[:150]}'\n")
