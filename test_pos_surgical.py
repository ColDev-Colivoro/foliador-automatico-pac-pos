import re
import zipfile
import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

def test_inject(xml_str, codigo, folio_str):
    # 1. POS_R1 -> Tabla flotante convertida a única tarjeta rectangular
    if codigo == "POS_R1":
        tbl_regex = re.compile(r'<w:tbl>(?:(?!</w:tbl>).)*?tblpX="8263".*?</w:tbl>', re.DOTALL)
        single_rect_tbl = (
            '<w:tbl><w:tblPr><w:tblStyle w:val="Tablaconcuadrcula"/>'
            '<w:tblpPr w:leftFromText="141" w:rightFromText="141" w:vertAnchor="text" w:horzAnchor="page" w:tblpX="8263" w:tblpY="-150"/>'
            '<w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3256"/></w:tblGrid>'
            '<w:tr w:rsidR="001136DA" w14:paraId="1DEEED86" w14:textId="77777777" w:rsidTr="000758C0">'
            '<w:trPr><w:trHeight w:val="365"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="3256" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="E8E8E8" w:themeFill="background2"/></w:tcPr>'
            '<w:p w14:paraId="322661FE" w:rsidR="001136DA" w:rsidRDefault="001136DA">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr><w:t>N° FOLIO: {folio_str}</w:t></w:r>'
            '</w:p></w:tc></w:tr></w:tbl>'
        )
        return tbl_regex.sub(single_rect_tbl, xml_str, count=1)

    # 2. POS_R3 -> Tabla flotante convertida a única tarjeta rectangular
    if codigo == "POS_R3":
        tbl_regex = re.compile(r'<w:tbl>(?:(?!</w:tbl>).)*?w:tblpY="7".*?</w:tbl>', re.DOTALL)
        single_rect_tbl = (
            '<w:tbl><w:tblPr><w:tblStyle w:val="Tablaconcuadrcula"/>'
            '<w:tblpPr w:leftFromText="141" w:rightFromText="141" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="7"/>'
            '<w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2879"/></w:tblGrid>'
            '<w:tr w:rsidR="008F2E9E" w14:paraId="0E4FF744" w14:textId="77777777" w:rsidTr="00437BEB">'
            '<w:trPr><w:trHeight w:val="440"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="2879" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="E8E8E8" w:themeFill="background2"/></w:tcPr>'
            '<w:p w14:paraId="16F685F4" w:rsidR="008F2E9E" w:rsidRDefault="008F2E9E">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr><w:t>N° FOLIO: {folio_str}</w:t></w:r>'
            '</w:p></w:tc></w:tr></w:tbl>'
        )
        return tbl_regex.sub(single_rect_tbl, xml_str, count=1)

    # 3. POS_R4 -> Tabla flotante convertida a única tarjeta rectangular
    if codigo == "POS_R4":
        tbl_regex = re.compile(r'<w:tbl>(?:(?!</w:tbl>).)*?w:tblpY="172".*?</w:tbl>', re.DOTALL)
        single_rect_tbl = (
            '<w:tbl><w:tblPr><w:tblStyle w:val="Tablaconcuadrcula"/>'
            '<w:tblpPr w:leftFromText="141" w:rightFromText="141" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="172"/>'
            '<w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2760"/></w:tblGrid>'
            '<w:tr w:rsidR="006C4E56" w14:paraId="4228FD9F" w14:textId="77777777" w:rsidTr="00645E02">'
            '<w:trPr><w:trHeight w:val="365"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="2760" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="E8E8E8" w:themeFill="background2"/></w:tcPr>'
            '<w:p w14:paraId="17065DD4" w:rsidR="006C4E56" w:rsidRDefault="006C4E56">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr><w:t>N° FOLIO: {folio_str}</w:t></w:r>'
            '</w:p></w:tc></w:tr></w:tbl>'
        )
        return tbl_regex.sub(single_rect_tbl, xml_str, count=1)

    # 4. POS_R2 -> Párrafo exacto con id 4ABEE3F6
    if codigo == "POS_R2":
        target = re.compile(r'<w:p [^>]*?paraId="4ABEE3F6"[^>]*?>.*?</w:p>', re.DOTALL)
        new_p = (
            '<w:p w14:paraId="4ABEE3F6" w14:textId="0E3D8556" w:rsidR="00AE0822" w:rsidRDefault="00AE0822" w:rsidP="004C1A34">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr><w:t>FOLIO: {folio_str}</w:t></w:r>'
            '</w:p>'
        )
        return target.sub(new_p, xml_str, count=1)

    # 5. PAC_PC5 -> Párrafo 76FBC3D6 (FOLIO) y 3394D40A (número)
    if codigo == "PAC_PC5":
        target_lbl = re.compile(r'<w:p [^>]*?paraId="76FBC3D6"[^>]*?>.*?</w:p>', re.DOTALL)
        new_lbl = (
            '<w:p w14:paraId="76FBC3D6" w:rsidR="00B44498" w:rsidRDefault="00B44498" w:rsidP="00463808">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="Calibri" w:cs="Calibri"/><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="Calibri" w:cs="Calibri"/><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>FOLIO</w:t></w:r>'
            '</w:p>'
        )
        xml_str = target_lbl.sub(new_lbl, xml_str, count=1)

        target_num = re.compile(r'<w:p [^>]*?paraId="3394D40A"[^>]*?>.*?</w:p>', re.DOTALL)
        new_num = (
            '<w:p w14:paraId="3394D40A" w:rsidR="00B44498" w:rsidRDefault="00EE1DDC" w:rsidP="00463808">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="Calibri" w:cs="Calibri"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:eastAsia="Calibri" w:cs="Calibri"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-ES"/></w:rPr><w:t>{folio_str}</w:t></w:r>'
            '</w:p>'
        )
        return target_num.sub(new_num, xml_str, count=1)

    # 6. PAC_PC6 -> Párrafo 44F446EC (FOLIO) y 6A5193E3 (número)
    if codigo == "PAC_PC6":
        target_lbl = re.compile(r'<w:p [^>]*?paraId="44F446EC"[^>]*?>.*?</w:p>', re.DOTALL)
        new_lbl = (
            '<w:p w14:paraId="44F446EC" w:rsidR="002A61BE" w:rsidRDefault="002A61BE" w:rsidP="002A61BE">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>FOLIO</w:t></w:r>'
            '</w:p>'
        )
        xml_str = target_lbl.sub(new_lbl, xml_str, count=1)

        target_num = re.compile(r'<w:p [^>]*?paraId="6A5193E3"[^>]*?>.*?</w:p>', re.DOTALL)
        new_num = (
            '<w:p w14:paraId="6A5193E3" w:rsidR="00E70877" w:rsidRDefault="007A2A70" w:rsidP="00E70877">'
            '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-ES"/></w:rPr><w:t>{folio_str}</w:t></w:r>'
            '</w:p>'
        )
        return target_num.sub(new_num, xml_str, count=1)

    return xml_str

# Ejecutar test sobre POS_R1, POS_R2, POS_R3, POS_R4, PAC_PC5, PAC_PC6
with open('config/plantillas.json', 'r', encoding='utf-8') as f:
    plantillas = json.load(f)

for p in plantillas:
    cod = p['codigo']
    if cod in ['POS_R1', 'POS_R2', 'POS_R3', 'POS_R4', 'PAC_PC5', 'PAC_PC6']:
        path = 'plantillas/' + p['archivo_plantilla']
        with zipfile.ZipFile(path) as z:
            xml = z.read('word/document.xml').decode('utf-8')
        res = test_inject(xml, cod, "00398")
        
        # Buscar en el resultado cómo quedó la tabla o párrafos
        for tbl in re.finditer(r'<w:tbl>.*?</w:tbl>', res, re.DOTALL):
            t = ' '.join(re.sub(r'<[^>]+>', ' ', tbl.group(0)).split())
            if '00398' in t:
                print(f"=== {cod} RESULTADO TABLA ===")
                print("Texto:", t)
                print("Cols:", re.findall(r'<w:gridCol[^>]*?>', tbl.group(0)))
                print("340 presente?", "340" in tbl.group(0))
                print("000 4 1 4 presente?", "000 4 1 4" in tbl.group(0))
                print("000 38 5 presente?", "000 38 5" in tbl.group(0))
                print()
