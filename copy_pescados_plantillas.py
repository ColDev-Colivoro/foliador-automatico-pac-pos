import shutil
import os

src_dir = os.path.abspath(r"../Programa Aseguramiento de calidad/Docuemntos Pos paC prerrequisitos/PAC VERSION 15/registros pac pescados")
dst_dir = os.path.abspath("plantillas")

file_map = {
    "PCC1 Recepcion MP Pescados.docx": "PAC_PCC1_Recepcion_Pescados_V15.docx",
    "Pc2 Almacenamiento Pescados.docx": "PAC_PC2_Alm_Pescados_V15.docx",
    "pc3 Temperatura en Linea de proceso.docx": "PAC_PC3_Temp_Linea_Pescados_V15.docx",
    "pc4 Congelacion y enfriado pescados.docx": "PAC_PC4_Congelacion_Enfriado_Pescados_V15.docx",
    "pcc5 Empaque y Reempaque pescados.docx": "PAC_PC5_Empaque_Pescados_V15.docx",
    "Pc 6 Despacho Pescados.docx": "PAC_PC6_Despacho_Pescados_V15.docx",
    "Pc7 Control de Agua y Glaseo pescados.docx": "PAC_PC7_Agua_Glaseo_Pescados_V15.docx",
}

print(f"Copiando {len(file_map)} plantillas a {dst_dir}...\n")
for src_name, dst_name in file_map.items():
    src_path = os.path.join(src_dir, src_name)
    dst_path = os.path.join(dst_dir, dst_name)
    shutil.copy2(src_path, dst_path)
    print(f"OK: {src_name} -> {dst_name} ({os.path.getsize(dst_path)} bytes)")

print("\nCopia completada con éxito.")
