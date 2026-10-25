import zipfile
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

for cod, fn in [
    ('PAC_PC2', 'PAC_PC2_Alm_Cefalopodos_V15.docx'),
    ('PAC_PC4', 'PAC_PC4_Control_Temperatura_Linea_V15.docx'),
    ('PAC_PC5', 'PAC_PC5_Congelacion_Enfriado_V15.docx'),
    ('PAC_PC6', 'PAC_PC6_Empaque_Reempaque_V15.docx'),
    ('PAC_DIAS', 'PAC_Dias_Productivos_V15.docx')
]:
    print('====================================================')
    print(f'INSPECCION DETALLADA: {cod} ({fn})')
    print('====================================================')
    with zipfile.ZipFile('plantillas/' + fn, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')

    # Buscar todo texto en los primeros 15 párrafos o tablas
    for m in re.finditer(r'<w:p(?:\s+[^>]*)?>.*?</w:p>', xml, re.DOTALL):
        txt = ' '.join(re.sub(r'<[^>]+>', ' ', m.group(0)).split())
        if any(w in txt.upper() for w in ['FOLIO', 'LOTE', '38', '37', '000', 'DIAS']):
            if not any(ign in txt.lower() for ign in ['empaque', 'guía', 'haccp', 'lota']):
                pid = re.findall(r'paraId="([^"]+)"', m.group(0))
                print(f'Párrafo ({pid}): {txt}')
                print(f'XML snippet: {m.group(0)[:300]}')
                print('---')
