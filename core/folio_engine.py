"""
Motor central de foliación para planillas PAC y POS.
Garantiza la preservación del 100% de la estructura original de Word:
logos, tablas, bordes y fuentes intactos.
Inyecta el número de folio centrado quirúrgicamente sin alterar namespaces de OpenXML.
Soporta ejecución en desarrollo y compilado como .EXE independiente con PyInstaller.
"""

import os
import json
import re
import datetime
import zipfile
import subprocess
import sys
import time

try:
    import win32api
    import win32print
    import win32con
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


class FolioEngine:
    def __init__(self, base_dir=None):
        if base_dir is None:
            if getattr(sys, 'frozen', False):
                self.app_dir = os.path.dirname(sys.executable)
                self.bundle_dir = getattr(sys, '_MEIPASS', self.app_dir)
            else:
                self.app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                self.bundle_dir = self.app_dir
        else:
            self.app_dir = os.path.abspath(base_dir)
            self.bundle_dir = self.app_dir

        # Rutas de configuración y plantillas (prioridad directorio de app, luego bundle)
        self.config_path = os.path.join(self.app_dir, "config", "plantillas.json")
        if not os.path.exists(self.config_path):
            self.config_path = os.path.join(self.bundle_dir, "config", "plantillas.json")

        self.plantillas_dir = os.path.join(self.app_dir, "plantillas")
        if not os.path.exists(self.plantillas_dir):
            self.plantillas_dir = os.path.join(self.bundle_dir, "plantillas")

        # Salida y persistencia de datos (siempre en el directorio visible de la app)
        self.data_dir = os.path.join(self.app_dir, "data")
        self.state_path = os.path.join(self.data_dir, "estado_folios.json")
        self.salida_dir = os.path.join(self.app_dir, "salida")

        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.salida_dir, exist_ok=True)

        self.plantillas = self._cargar_config()
        self.estado = self._cargar_estado()

    def _cargar_config(self):
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"No se encontró el archivo de configuración: {self.config_path}")
        with open(self.config_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _cargar_estado(self):
        estado = {}
        if os.path.exists(self.state_path):
            try:
                with open(self.state_path, "r", encoding="utf-8") as f:
                    estado = json.load(f)
            except Exception:
                pass

        modificado = False
        for p in self.plantillas:
            codigo = p["codigo"]
            if codigo not in estado:
                ultimo = p.get("ultimo_folio_defecto", 0)
                estado[codigo] = {
                    "ultimo_folio": ultimo,
                    "fecha_actualizacion": datetime.datetime.now().isoformat(),
                    "historial": [
                        {
                            "fecha": datetime.datetime.now().isoformat(),
                            "folio_inicial": ultimo,
                            "folio_final": ultimo,
                            "cantidad": 0,
                            "accion": "Inicialización base"
                        }
                    ]
                }
                modificado = True

        if modificado or not os.path.exists(self.state_path):
            self._guardar_estado(estado)

        return estado

    def _guardar_estado(self, estado=None):
        if estado is not None:
            self.estado = estado
        with open(self.state_path, "w", encoding="utf-8") as f:
            json.dump(self.estado, f, indent=2, ensure_ascii=False)

    def get_plantilla_info(self, codigo):
        for p in self.plantillas:
            if p["codigo"] == codigo:
                return p
        return None

    def get_ultimo_folio(self, codigo):
        if codigo in self.estado:
            return self.estado[codigo]["ultimo_folio"]
        info = self.get_plantilla_info(codigo)
        return info.get("ultimo_folio_defecto", 0) if info else 0

    def set_ultimo_folio(self, codigo, nuevo_folio, nota="Ajuste manual"):
        if codigo not in self.estado:
            self.estado[codigo] = {"ultimo_folio": 0, "historial": []}

        self.estado[codigo]["ultimo_folio"] = int(nuevo_folio)
        self.estado[codigo]["fecha_actualizacion"] = datetime.datetime.now().isoformat()
        self.estado[codigo]["historial"].append({
            "fecha": datetime.datetime.now().isoformat(),
            "folio_inicial": int(nuevo_folio),
            "folio_final": int(nuevo_folio),
            "cantidad": 0,
            "accion": nota
        })
        self._guardar_estado()

    def get_preferencias_usuario(self):
        """Retorna las preferencias guardadas del usuario."""
        ruta = os.path.join(self.data_dir, "preferencias_usuario.json")
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def guardar_preferencia(self, clave, valor):
        """Guarda una clave-valor en las preferencias del usuario."""
        prefs = self.get_preferencias_usuario()
        prefs[clave] = valor
        ruta = os.path.join(self.data_dir, "preferencias_usuario.json")
        try:
            with open(ruta, "w", encoding="utf-8") as f:
                json.dump(prefs, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def get_impresora_guardada(self):
        """Retorna la última impresora seleccionada por el usuario, si existe."""
        prefs = self.get_preferencias_usuario()
        return prefs.get("impresora_seleccionada")

    def guardar_impresora_guardada(self, nombre):
        """Persiste la impresora seleccionada para recordarla en próximas sesiones."""
        if nombre:
            self.guardar_preferencia("impresora_seleccionada", nombre)

    def get_impresoras_disponibles(self):
        """Retorna la lista de nombres de impresoras instaladas/disponibles en Windows."""
        if not WIN32_AVAILABLE:
            return []
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            nombres = [p[2] for p in printers if p and len(p) > 2 and p[2]]
            vistos = set()
            resultado = []
            for n in nombres:
                if n not in vistos:
                    vistos.add(n)
                    resultado.append(n)
            return resultado
        except Exception:
            return []

    def get_impresora_predeterminada(self):
        """Retorna el nombre de la impresora predeterminada de Windows si está disponible, o None."""
        if WIN32_AVAILABLE:
            try:
                pname = win32print.GetDefaultPrinter()
                if pname and pname.strip():
                    return pname.strip()
            except Exception:
                pass
        return None

    def set_impresora_predeterminada_sistema(self, nombre_impresora):
        """Establece la impresora indicada como la predeterminada a nivel del sistema Windows."""
        if not WIN32_AVAILABLE or not nombre_impresora:
            return False
        try:
            win32print.SetDefaultPrinter(nombre_impresora)
            return True
        except Exception:
            return False

    def hay_impresora_disponible(self):
        """Verifica si existe al menos una impresora accesible en Windows."""
        if not WIN32_AVAILABLE:
            return False
        try:
            disponibles = self.get_impresoras_disponibles()
            return len(disponibles) > 0
        except Exception:
            return False

    def set_impresora_duplex(self, activar=True, nombre_impresora=None):
        """
        Configura el modo dúplex (ambos lados) en la impresora especificada o predeterminada.
        DMDUP_VERTICAL (2) = voltear en borde largo (vertical estándar).
        """
        if not WIN32_AVAILABLE:
            return False
        try:
            pname = nombre_impresora or self.get_impresora_guardada() or self.get_impresora_predeterminada()
            if not pname:
                return False
            hprinter = win32print.OpenPrinter(pname, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
            try:
                info = win32print.GetPrinter(hprinter, 2)
                devmode = info["pDevMode"]
                devmode.Duplex = win32con.DMDUP_VERTICAL if activar else win32con.DMDUP_SIMPLEX
                devmode.Fields |= win32con.DM_DUPLEX
                win32print.SetPrinter(hprinter, 2, info, 0)
                return True
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception:
            return False

    def imprimir_lote(self, lista_rutas, doble_cara=True, callback_progreso=None, nombre_impresora=None):
        """
        Imprime una lista de documentos Word en lote asegurando que CADA uno
        conserve la configuración dúplex (ambos lados) sin interrupciones ni cierres prematuros.
        Envía a la impresora indicada o a la preferida/predeterminada.
        """
        if not lista_rutas:
            return True

        if not self.hay_impresora_disponible():
            raise RuntimeError("No se detectó ninguna impresora activa o configurada en Windows.")

        pname = nombre_impresora or self.get_impresora_guardada() or self.get_impresora_predeterminada()
        if not pname:
            raise RuntimeError("No hay ninguna impresora seleccionada ni configurada en Windows.")

        for r in lista_rutas:
            if not os.path.exists(r):
                raise FileNotFoundError(f"Archivo no encontrado para imprimir: {r}")

        if WIN32_AVAILABLE:
            if doble_cara:
                self.set_impresora_duplex(True, nombre_impresora=pname)

            word = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0  # wdAlertsNone
                word.Options.PrintBackground = False

                # Asignar la impresora elegida en Word
                try:
                    word.ActivePrinter = pname
                except Exception:
                    # Intento con sufijo de puerto si Word COM lo requiere según idioma de Office
                    asignado = False
                    for i in range(15):
                        for prep in ["en Ne", "on Ne"]:
                            try:
                                word.ActivePrinter = f"{pname} {prep}{i:02d}:"
                                asignado = True
                                break
                            except Exception:
                                pass
                        if asignado:
                            break

                total = len(lista_rutas)
                for i, ruta in enumerate(lista_rutas):
                    ruta_abs = os.path.abspath(ruta)
                    doc = word.Documents.Open(ruta_abs, ReadOnly=True)
                    try:
                        doc.PrintOut(Background=False)
                    finally:
                        doc.Close(False)

                    if callback_progreso:
                        try:
                            callback_progreso(i + 1, total, os.path.basename(ruta))
                        except Exception:
                            pass

                    time.sleep(0.5)

                return True
            except Exception as e:
                # Falla en Word COM: no iterar ciegamente abriendo ventanas modales de error por cada archivo
                raise RuntimeError(f"Error al enviar documentos a la impresora '{pname}': {e}")
            finally:
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
        else:
            raise RuntimeError("La impresión directa requiere controladores de impresora disponibles en Windows.")

    def imprimir_documento(self, ruta_archivo, doble_cara=True, nombre_impresora=None):
        """
        Envía un documento Word directamente a la impresora.
        """
        return self.imprimir_lote([ruta_archivo], doble_cara=doble_cara, nombre_impresora=nombre_impresora)

    def _inyectar_folio_quirurgico(self, xml_str, codigo, folio_str):
        """
        Reemplaza quirúrgicamente el folio con alineación centrada (<w:jc w:val="center"/>)
        sin alterar ni reserializar namespaces de OpenXML para conservar intactos todos los logos y gráficos.
        """
        # Caso 1: PAC_PCC1 -> Párrafo exacto con id 594CB2BD
        if codigo == "PAC_PCC1":
            target = re.compile(r'<w:p [^>]*?paraId="594CB2BD"[^>]*?>.*?</w:p>', re.DOTALL)
            new_p = (
                '<w:p w14:paraId="594CB2BD" w:rsidR="009B78B8" w:rsidRDefault="009B78B8">'
                '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr><w:t>Folio {folio_str}</w:t></w:r>'
                '</w:p>'
            )
            return target.sub(new_p, xml_str)

        # Caso 2: PAC_PC7 -> Párrafo exacto con id 095F0636
        if codigo == "PAC_PC7":
            target = re.compile(r'<w:p [^>]*?paraId="095F0636"[^>]*?>.*?</w:p>', re.DOTALL)
            new_p = (
                '<w:p w14:paraId="095F0636" w:rsidR="005A7800" w:rsidRDefault="005A7800">'
                '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr><w:t>Folio {folio_str}</w:t></w:r>'
                '</w:p>'
            )
            return target.sub(new_p, xml_str)

        # Caso 3: PAC_PC2 -> Celda superior derecha de la primera tabla
        if codigo == "PAC_PC2":
            cell_replacement = (
                '<w:tc><w:tcPr><w:tcW w:w="2848" w:type="dxa"/></w:tcPr>'
                '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr><w:t>FOLIO: {folio_str}</w:t></w:r></w:p>'
                '</w:tc>'
            )
            pattern_pc2 = re.compile(r'<w:tc[^>]*>(?:(?!</w:tc>).)*?(?:<w:t>Lote</w:t>|FOLIO:\s*\d+).*?</w:tc>', re.DOTALL)
            return pattern_pc2.sub(cell_replacement, xml_str, count=1)

        # Caso 4: PAC_PC3 -> Tabla de folio flotante (convertida a un solo rectángulo centrado sin celda extra 000104)
        if codigo == "PAC_PC3":
            tbl_regex = re.compile(r'<w:tbl>(?:(?!</w:tbl>).)*?tblpX="8271".*?</w:tbl>', re.DOTALL)
            single_rect_tbl = (
                '<w:tbl><w:tblPr><w:tblStyle w:val="Tablaconcuadrcula"/>'
                '<w:tblpPr w:leftFromText="141" w:rightFromText="141" w:vertAnchor="text" w:horzAnchor="page" w:tblpX="8271" w:tblpY="394"/>'
                '<w:tblOverlap w:val="never"/><w:tblW w:w="0" w:type="auto"/>'
                '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
                '<w:tblGrid><w:gridCol w:w="3114"/></w:tblGrid>'
                '<w:tr w:rsidR="00725527" w:rsidRPr="00661D19" w14:paraId="178AD15E" w14:textId="77777777">'
                '<w:trPr><w:trHeight w:val="246"/></w:trPr>'
                '<w:tc><w:tcPr><w:tcW w:w="3114" w:type="dxa"/></w:tcPr>'
                '<w:p w:rsidR="00725527" w:rsidRDefault="00725527">'
                '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>FOLIO {folio_str}</w:t></w:r>'
                '</w:p></w:tc></w:tr></w:tbl>'
            )
            return tbl_regex.sub(single_rect_tbl, xml_str, count=1)

        # Caso 5: Resto de plantillas (PC4, PC5, PC6, Dias, POS 1..4)
        def repl_p(match):
            p_full = match.group(0)
            inner_text = re.sub(r'<[^>]+>', '', p_full).strip()

            if "folios empaque" in inner_text.lower():
                return p_full

            is_folio = (
                re.match(r"^(?:N[°º\s]*FOLIO|FOLIO)[\s:]*\d*$", inner_text, re.IGNORECASE) or
                re.match(r"^Folio\s*\d+$", inner_text, re.IGNORECASE)
            )

            if not is_folio:
                return p_full

            if "n°" in inner_text.lower() or "n " in inner_text.lower():
                texto_folio = f"N° FOLIO: {folio_str}"
            elif "folio:" in inner_text.lower():
                texto_folio = f"FOLIO: {folio_str}"
            elif inner_text.isupper():
                texto_folio = f"FOLIO {folio_str}"
            else:
                texto_folio = f"Folio {folio_str}"

            p_tag_match = re.match(r'<w:p(?:\s+[^>]*)?>', p_full)
            p_tag = p_tag_match.group(0) if p_tag_match else '<w:p>'

            return (
                f'{p_tag}'
                '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:bCs/><w:lang w:val="es-ES"/></w:rPr><w:t>{texto_folio}</w:t></w:r>'
                '</w:p>'
            )

        p_general = re.compile(r'<w:p(?:\s+[^>]*)?>.*?</w:p>', re.DOTALL)
        return p_general.sub(repl_p, xml_str)

    def generar_planillas(self, codigo, cantidad=1, carpeta_sublote=None, imprimir_al_generar=False, doble_cara=True, nombre_impresora=None):
        info = self.get_plantilla_info(codigo)
        if not info:
            raise ValueError(f"Plantilla no reconocida: {codigo}")

        cantidad = int(cantidad)
        if cantidad < 1:
            raise ValueError("La cantidad de planillas debe ser al menos 1.")

        ultimo_folio = self.get_ultimo_folio(codigo)
        folio_inicial = ultimo_folio + 1
        folio_final = ultimo_folio + cantidad
        digitos = info.get("digitos", 5)

        plantilla_src = os.path.join(self.plantillas_dir, info["archivo_plantilla"])
        if not os.path.exists(plantilla_src):
            raise FileNotFoundError(f"Archivo de plantilla no encontrado: {plantilla_src}")

        if carpeta_sublote:
            dest_dir = os.path.join(self.salida_dir, carpeta_sublote)
        else:
            fecha_hoy = datetime.date.today().strftime("%Y-%m-%d")
            dest_dir = os.path.join(self.salida_dir, f"{fecha_hoy}_{codigo}")

        os.makedirs(dest_dir, exist_ok=True)

        archivos_generados = []

        for num in range(folio_inicial, folio_final + 1):
            folio_str = f"{num:0{digitos}d}"
            nombre_salida = f"{codigo}_Folio_{folio_str}.docx"
            ruta_salida = os.path.join(dest_dir, nombre_salida)

            with zipfile.ZipFile(plantilla_src, "r") as zin:
                with zipfile.ZipFile(ruta_salida, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        if item.filename == "word/document.xml":
                            xml_str = data.decode("utf-8")
                            xml_str = self._inyectar_folio_quirurgico(xml_str, codigo, folio_str)
                            data = xml_str.encode("utf-8")
                        zout.writestr(item, data)

            archivos_generados.append({
                "folio": folio_str,
                "numero": num,
                "archivo": ruta_salida,
            })

        impreso_exitoso = False
        error_impresion = None
        if imprimir_al_generar:
            if not self.hay_impresora_disponible():
                error_impresion = "No se detectó ninguna impresora activa en Windows."
            else:
                try:
                    rutas_a_imprimir = [a["archivo"] for a in archivos_generados]
                    self.imprimir_lote(rutas_a_imprimir, doble_cara=doble_cara, nombre_impresora=nombre_impresora)
                    impreso_exitoso = True
                except Exception as e:
                    error_impresion = str(e)

        self.estado[codigo]["ultimo_folio"] = folio_final
        self.estado[codigo]["fecha_actualizacion"] = datetime.datetime.now().isoformat()
        self.estado[codigo]["historial"].append({
            "fecha": datetime.datetime.now().isoformat(),
            "folio_inicial": folio_inicial,
            "folio_final": folio_final,
            "cantidad": cantidad,
            "accion": f"Emisión de {cantidad} planilla(s) [{folio_inicial:0{digitos}d} al {folio_final:0{digitos}d}]",
            "carpeta": dest_dir,
            "impreso": impreso_exitoso
        })
        self._guardar_estado()

        return {
            "codigo": codigo,
            "nombre": info["nombre"],
            "folio_inicial": f"{folio_inicial:0{digitos}d}",
            "folio_final": f"{folio_final:0{digitos}d}",
            "cantidad": cantidad,
            "carpeta_salida": dest_dir,
            "archivos": archivos_generados,
            "impreso": impreso_exitoso,
            "error_impresion": error_impresion
        }
