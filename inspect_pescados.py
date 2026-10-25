import zipfile
import re
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

pescados_dir = os.path.abspath(r"../Programa Aseguramiento de calidad/Docuemntos Pos paC prerrequisitos/PAC VERSION 15/registros pac pescados")
path = os.path.join(pescados_dir, "pcc5 Empaque y Reempaque pescados.docx")

with zipfile.ZipFile(path, 'r') as z:
    xml = z.read('word/document.xml').decode('utf-8')

print("DOCUMENT XML:")
print(xml)
