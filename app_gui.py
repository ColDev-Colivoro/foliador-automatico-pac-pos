"""
Interfaz Gráfica de Escritorio (GUI) para el Foliador Automático PAC / POS.
Usa Tkinter estándar (sin dependencias externas pesadas).
Preserva 100% la estructura original de los documentos Word (logos, tablas, bordes).
Centrado exacto de los números de folio.
Envío directo a la impresora predeterminada (EPSON).
Selector de alta visibilidad y contraste.
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.folio_engine import FolioEngine


class FoliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.engine = FolioEngine()

        self.title("ColDev · Foliador Oficial PAC / POS (Preservación Exacta Word)")
        self.geometry("880x820")
        self.minsize(820, 750)
        self.configure(bg="#0f172a")

        # Configurar estilos de ttk
        self._configurar_estilos()

        # Variables de estado
        self.var_cantidad = tk.IntVar(value=1)
        self.var_ultimo_folio = tk.StringVar()
        self.var_resumen_rango = tk.StringVar()
        self.var_auto_imprimir = tk.BooleanVar(value=False)
        self.var_doble_cara = tk.BooleanVar(value=True)
        self.ultimos_archivos_generados = []

        # Mapa de variables de checkbox y widgets
        self.check_vars = {}
        self.check_widgets = {}

        # Construir UI
        self._construir_ui()
        self._actualizar_estado_impresora()

        # Seleccionar primera plantilla por defecto
        if self.engine.plantillas:
            primer_cod = self.engine.plantillas[0]["codigo"]
            if primer_cod in self.check_vars:
                self.check_vars[primer_cod].set(True)
        self._al_cambiar_seleccion_checkboxes()

    def _configurar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Colores generales
        style.configure(".", background="#0f172a", foreground="#f8fafc", font=("Segoe UI", 10))
        style.configure("TFrame", background="#0f172a")
        style.configure("Card.TFrame", background="#1e293b", relief="flat")

        style.configure("TLabel", background="#0f172a", foreground="#f8fafc")
        style.configure("Card.TLabel", background="#1e293b", foreground="#f8fafc")
        style.configure("Muted.TLabel", background="#1e293b", foreground="#94a3b8", font=("Segoe UI", 9))
        style.configure("Header.TLabel", background="#0f172a", foreground="#38bdf8", font=("Segoe UI", 16, "bold"))
        style.configure("SubHeader.TLabel", background="#0f172a", foreground="#94a3b8", font=("Segoe UI", 9))

        style.configure("Highlight.TLabel", background="#1e293b", foreground="#10b981", font=("Segoe UI", 13, "bold"))
        style.configure("Warning.TLabel", background="#1e293b", foreground="#f59e0b", font=("Segoe UI", 12, "bold"))
        style.configure("Danger.TLabel", background="#1e293b", foreground="#f87171", font=("Segoe UI", 10, "bold"))

        style.configure("Primary.TButton", background="#059669", foreground="#ffffff", font=("Segoe UI", 11, "bold"), padding=10)
        style.map("Primary.TButton",
                  background=[("active", "#10b981"), ("pressed", "#047857")],
                  foreground=[("active", "#ffffff")])

        style.configure("Print.TButton", background="#0284c7", foreground="#ffffff", font=("Segoe UI", 10, "bold"), padding=8)
        style.map("Print.TButton",
                  background=[("active", "#38bdf8"), ("pressed", "#0369a1")])

        style.configure("Secondary.TButton", background="#334155", foreground="#f8fafc", font=("Segoe UI", 10), padding=8)
        style.map("Secondary.TButton",
                  background=[("active", "#475569"), ("pressed", "#1e293b")])

        style.configure("Small.TButton", background="#3b82f6", foreground="#ffffff", font=("Segoe UI", 8, "bold"), padding=4)
        style.map("Small.TButton",
                  background=[("active", "#60a5fa"), ("pressed", "#2563eb")])

        style.configure("Category.TLabel", background="#1e293b", foreground="#38bdf8", font=("Segoe UI", 9, "bold"))
        style.configure("ActionText.TButton", background="#334155", foreground="#f8fafc", font=("Segoe UI", 8, "bold"), padding=(6, 2))
        style.map("ActionText.TButton",
                  background=[("active", "#475569"), ("pressed", "#1e293b")])

        style.configure("TSpinbox", fieldbackground="#ffffff", background="#0284c7", foreground="#0f172a", font=("Segoe UI", 11, "bold"))
        style.configure("TCheckbutton", background="#1e293b", foreground="#f8fafc", font=("Segoe UI", 10))

        # Personalizar menú desplegable de la lista
        self.option_add("*TCombobox*Listbox.background", "#1e293b")
        self.option_add("*TCombobox*Listbox.foreground", "#ffffff")
        self.option_add("*TCombobox*Listbox.selectBackground", "#0284c7")
        self.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        self.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10, "bold"))

    def _construir_ui(self):
        # 1. ENCABEZADO
        header_frame = ttk.Frame(self, padding=(24, 16, 24, 8))
        header_frame.pack(fill="x")

        ttk.Label(header_frame, text="⚡ COLDEV · FOLIADOR OFICIAL PAC / POS", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header_frame,
            text="Inyección limpia de folios correlativos centrados · Logos, encabezados y tablas 100% originales intactos.",
            style="SubHeader.TLabel"
        ).pack(anchor="w", pady=(3, 0))

        # Contenedor principal
        main_frame = ttk.Frame(self, padding=(24, 0, 24, 16))
        main_frame.pack(fill="both", expand=True)

        # 2. SELECCIÓN DE PLANILLAS (CARD CON CHECKBOXES)
        card_sel = ttk.Frame(main_frame, style="Card.TFrame", padding=12)
        card_sel.pack(fill="x", pady=(0, 10))

        # Barra superior con título y botones de acción rápida
        hdr_sel = ttk.Frame(card_sel, style="Card.TFrame")
        hdr_sel.pack(fill="x", pady=(0, 6))

        ttk.Label(hdr_sel, text="1. SELECCIONAR PLANILLAS A EMITIR (Podés marcar una o varias a la vez)", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(side="left")

        btn_box = ttk.Frame(hdr_sel, style="Card.TFrame")
        btn_box.pack(side="right")
        ttk.Button(btn_box, text="✅ Todas", style="ActionText.TButton", command=self._seleccionar_todas).pack(side="left", padx=2)
        ttk.Button(btn_box, text="⬜ Ninguna", style="ActionText.TButton", command=self._deseleccionar_todas).pack(side="left", padx=2)
        ttk.Button(btn_box, text="PAC (8)", style="ActionText.TButton", command=lambda: self._seleccionar_categoria("PAC")).pack(side="left", padx=2)
        ttk.Button(btn_box, text="POS (4)", style="ActionText.TButton", command=lambda: self._seleccionar_categoria("POS")).pack(side="left", padx=2)

        # Contenedor en 2 columnas para PAC y POS
        cols_frame = ttk.Frame(card_sel, style="Card.TFrame")
        cols_frame.pack(fill="x", pady=(2, 0))

        # Columna 1: PAC
        col_pac = ttk.Frame(cols_frame, style="Card.TFrame")
        col_pac.pack(side="left", fill="both", expand=True, padx=(0, 6))
        ttk.Label(col_pac, text="📋 PLANILLAS PAC (Aseguramiento de Calidad)", style="Category.TLabel").pack(anchor="w", pady=(0, 3))

        # Columna 2: POS
        col_pos = ttk.Frame(cols_frame, style="Card.TFrame")
        col_pos.pack(side="left", fill="both", expand=True, padx=(6, 0))
        ttk.Label(col_pos, text="🧹 PLANILLAS POS (Saneamiento y Operación)", style="Category.TLabel").pack(anchor="w", pady=(0, 3))

        for p in self.engine.plantillas:
            cod = p["codigo"]
            var = tk.BooleanVar(value=False)
            self.check_vars[cod] = var
            parent_col = col_pac if p["categoria"] == "PAC" else col_pos

            txt = self._obtener_texto_checkbox(p)
            chk = ttk.Checkbutton(
                parent_col,
                text=txt,
                variable=var,
                command=self._al_cambiar_seleccion_checkboxes
            )
            chk.pack(anchor="w", pady=1)
            self.check_widgets[cod] = chk

        # 3. ESTADO DEL FOLIO Y CANTIDAD (CARD)
        card_folio = ttk.Frame(main_frame, style="Card.TFrame", padding=12)
        card_folio.pack(fill="x", pady=(0, 10))

        ttk.Label(card_folio, text="2. ESTADO DEL PLANILLERO Y CORRELATIVO", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 6))

        grid_frame = ttk.Frame(card_folio, style="Card.TFrame")
        grid_frame.pack(fill="x")

        # Columna 1: Último folio / Selección
        f1 = ttk.Frame(grid_frame, style="Card.TFrame")
        f1.grid(row=0, column=0, sticky="w", padx=(0, 20))

        ttk.Label(f1, text="Planilla / Folio Actual:", style="Muted.TLabel").pack(anchor="w")
        sub_f1 = ttk.Frame(f1, style="Card.TFrame")
        sub_f1.pack(anchor="w", pady=(4, 0))
        ttk.Label(sub_f1, textvariable=self.var_ultimo_folio, style="Warning.TLabel").pack(side="left", padx=(0, 8))
        btn_modificar = ttk.Button(sub_f1, text="✏️ Sincronizar Físico", style="Small.TButton", command=self._sincronizar_folio_fisico)
        btn_modificar.pack(side="left")

        # Columna 2: Cantidad a emitir por planilla
        f2 = ttk.Frame(grid_frame, style="Card.TFrame")
        f2.grid(row=0, column=1, sticky="w", padx=(0, 20))

        ttk.Label(f2, text="Cantidad por Planilla:", style="Muted.TLabel").pack(anchor="w")
        sub_f2 = ttk.Frame(f2, style="Card.TFrame")
        sub_f2.pack(anchor="w", pady=(4, 0))

        spin_cant = ttk.Spinbox(
            sub_f2,
            from_=1,
            to=200,
            textvariable=self.var_cantidad,
            width=5,
            command=self._actualizar_calculo_rango
        )
        spin_cant.pack(side="left", padx=(0, 4))
        spin_cant.bind("<KeyRelease>", lambda e: self._actualizar_calculo_rango())

        ttk.Button(sub_f2, text="+1", style="Secondary.TButton", width=3, command=lambda: self._sumar_cantidad(1)).pack(side="left", padx=1)
        ttk.Button(sub_f2, text="+5", style="Secondary.TButton", width=3, command=lambda: self._sumar_cantidad(5)).pack(side="left", padx=1)
        ttk.Button(sub_f2, text="+10", style="Secondary.TButton", width=3, command=lambda: self._sumar_cantidad(10)).pack(side="left", padx=1)

        # Columna 3: Próximo folio / Resumen total
        f3 = ttk.Frame(grid_frame, style="Card.TFrame")
        f3.grid(row=0, column=2, sticky="w")

        ttk.Label(f3, text="Total a Emitir (Centrados):", style="Muted.TLabel").pack(anchor="w")
        ttk.Label(f3, textvariable=self.var_resumen_rango, style="Highlight.TLabel").pack(anchor="w", pady=(4, 0))

        # 4. CONFIGURACIÓN DE IMPRESIÓN (CARD)
        card_print = ttk.Frame(main_frame, style="Card.TFrame", padding=14)
        card_print.pack(fill="x", pady=(0, 12))

        ttk.Label(card_print, text="3. SALIDA E IMPRESIÓN DIRECTA", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 4))

        self.lbl_impresora = ttk.Label(
            card_print,
            text="Buscando impresora...",
            style="Muted.TLabel"
        )
        self.lbl_impresora.pack(anchor="w", pady=(0, 6))

        self.chk_duplex = ttk.Checkbutton(
            card_print,
            text="📄 Imprimir por ambos lados (Dúplex: Planilla adelante + Instructivo atrás en 1 sola hoja)",
            variable=self.var_doble_cara
        )
        self.chk_duplex.pack(anchor="w", pady=(0, 4))

        self.chk_print = ttk.Checkbutton(
            card_print,
            text="⚡ Mandar a imprimir automáticamente a la impresora al hacer clic en Generar",
            variable=self.var_auto_imprimir
        )
        self.chk_print.pack(anchor="w")

        # 5. BOTONES DE ACCIÓN
        btn_bar = ttk.Frame(main_frame)
        btn_bar.pack(fill="x", pady=(2, 12))

        self.btn_generar = ttk.Button(
            btn_bar,
            text="🚀 GENERAR PLANILLAS FOLIADAS",
            style="Primary.TButton",
            command=self._generar_planillas
        )
        self.btn_generar.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_imprimir_lote = ttk.Button(
            btn_bar,
            text="🖨️ Imprimir Último Lote",
            style="Print.TButton",
            command=self._imprimir_ultimo_lote
        )
        self.btn_imprimir_lote.pack(side="left", padx=(0, 8))

        btn_abrir = ttk.Button(
            btn_bar,
            text="📂 Abrir Carpeta",
            style="Secondary.TButton",
            command=self._abrir_carpeta_salida
        )
        btn_abrir.pack(side="right")

        # 6. PANEL DE REGISTRO / AUDITORÍA
        card_log = ttk.Frame(main_frame, style="Card.TFrame", padding=12)
        card_log.pack(fill="both", expand=True)

        ttk.Label(card_log, text="REGISTRO DE ACTIVIDAD:", font=("Segoe UI", 9, "bold"), style="Card.TLabel").pack(anchor="w")

        self.txt_log = tk.Text(
            card_log,
            bg="#0b1120",
            fg="#94a3b8",
            insertbackground="#ffffff",
            font=("Consolas", 9),
            height=6,
            relief="flat",
            wrap="word"
        )
        self.txt_log.pack(fill="both", expand=True, pady=(4, 0))
        self._log("Listo. Seleccioná una planilla y hacé clic en Generar.")

    def _log(self, mensaje):
        def do_log():
            self.txt_log.insert("end", f"• {mensaje}\n")
            self.txt_log.see("end")
        self.after(0, do_log)

    def _actualizar_estado_impresora(self):
        hay_imp = self.engine.hay_impresora_disponible()
        impresora_actual = self.engine.get_impresora_predeterminada()
        if hay_imp and impresora_actual:
            self.lbl_impresora.config(
                text=f"🖨️ Impresora detectada: {impresora_actual}",
                style="Muted.TLabel"
            )
            self.chk_print.config(state="normal")
            if self.ultimos_archivos_generados:
                self.btn_imprimir_lote.config(state="normal")
            else:
                self.btn_imprimir_lote.config(state="disabled")
        else:
            self.lbl_impresora.config(
                text="⚠️ No se detectó ninguna impresora activa en Windows (Impresión desactivada)",
                style="Danger.TLabel"
            )
            self.var_auto_imprimir.set(False)
            self.chk_print.config(state="disabled")
            self.btn_imprimir_lote.config(state="disabled")

    def _obtener_texto_checkbox(self, plantilla):
        cod = plantilla["codigo"]
        ultimo = self.engine.get_ultimo_folio(cod)
        dig = plantilla.get("digitos", 5)
        nombre = plantilla["nombre"]
        nombre_limpio = nombre.split("·")[-1].strip() if "·" in nombre else nombre
        if len(nombre_limpio) > 28:
            nombre_limpio = nombre_limpio[:26] + ".."
        return f"{cod} · {nombre_limpio}  [Folio: {ultimo:0{dig}d}]"

    def _actualizar_labels_checkboxes(self):
        for p in self.engine.plantillas:
            cod = p["codigo"]
            if cod in self.check_widgets:
                self.check_widgets[cod].config(text=self._obtener_texto_checkbox(p))

    def _seleccionar_todas(self):
        for var in self.check_vars.values():
            var.set(True)
        self._al_cambiar_seleccion_checkboxes()

    def _deseleccionar_todas(self):
        for var in self.check_vars.values():
            var.set(False)
        self._al_cambiar_seleccion_checkboxes()

    def _seleccionar_categoria(self, cat):
        for p in self.engine.plantillas:
            cod = p["codigo"]
            if p["categoria"] == cat:
                self.check_vars[cod].set(True)
            else:
                self.check_vars[cod].set(False)
        self._al_cambiar_seleccion_checkboxes()

    def _al_cambiar_seleccion_checkboxes(self):
        seleccionadas = [cod for cod, var in self.check_vars.items() if var.get()]
        cant = 1
        try:
            cant = int(self.var_cantidad.get())
            if cant < 1:
                cant = 1
        except Exception:
            cant = 1

        total_sel = len(seleccionadas)
        if total_sel == 0:
            self.var_ultimo_folio.set("Ninguna elegida")
            self.var_resumen_rango.set("Marcá al menos 1 planilla")
            self.btn_generar.config(state="disabled")
        elif total_sel == 1:
            self.btn_generar.config(state="normal")
            cod = seleccionadas[0]
            info = self.engine.get_plantilla_info(cod)
            ultimo = self.engine.get_ultimo_folio(cod)
            dig = info.get("digitos", 5) if info else 5
            self.var_ultimo_folio.set(f"{ultimo:0{dig}d} ({cod})")
            ini = ultimo + 1
            fin = ultimo + cant
            if cant == 1:
                self.var_resumen_rango.set(f"Folio {ini:0{dig}d} (1 hoja)")
            else:
                self.var_resumen_rango.set(f"{ini:0{dig}d} al {fin:0{dig}d} ({cant} hojas)")
        else:
            self.btn_generar.config(state="normal")
            self.var_ultimo_folio.set(f"{total_sel} planillas activas")
            total_docs = total_sel * cant
            self.var_resumen_rango.set(f"{total_sel} planillas × {cant} = {total_docs} hojas Word")

    def _sumar_cantidad(self, delta):
        nueva = max(1, self.var_cantidad.get() + delta)
        self.var_cantidad.set(nueva)
        self._al_cambiar_seleccion_checkboxes()

    def _actualizar_calculo_rango(self):
        self._al_cambiar_seleccion_checkboxes()

    def _sincronizar_folio_fisico(self):
        seleccionadas = [cod for cod, var in self.check_vars.items() if var.get()]
        if len(seleccionadas) == 1:
            self._pedir_nuevo_folio(seleccionadas[0])
            return

        # Si hay varias o ninguna marcada, desplegar diálogo para elegir cuál sincronizar
        ventana_sync = tk.Toplevel(self)
        ventana_sync.title("Sincronizar Folio")
        ventana_sync.geometry("460x170")
        ventana_sync.configure(bg="#0f172a")
        ventana_sync.transient(self)
        ventana_sync.grab_set()

        ttk.Label(ventana_sync, text="Elegí la planilla que querés sincronizar:", style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(pady=(16, 6), padx=16, anchor="w")

        opciones = [f"{p['codigo']} · {p['nombre']}" for p in self.engine.plantillas]
        combo = ttk.Combobox(ventana_sync, values=opciones, state="readonly", width=48)
        combo.pack(pady=4, padx=16, fill="x")
        combo.current(0)

        def confirmar():
            idx = combo.current()
            if idx >= 0:
                cod = self.engine.plantillas[idx]["codigo"]
                ventana_sync.destroy()
                self._pedir_nuevo_folio(cod)

        ttk.Button(ventana_sync, text="Continuar", style="Primary.TButton", command=confirmar).pack(pady=12)

    def _pedir_nuevo_folio(self, codigo):
        info = self.engine.get_plantilla_info(codigo)
        if not info:
            return
        actual = self.engine.get_ultimo_folio(codigo)
        nuevo = simpledialog.askinteger(
            "Sincronizar con Planillero Físico",
            f"Planilla: {info['nombre']}\n\n"
            f"Último folio registrado en el sistema: {actual}\n"
            f"¿Cuál es el último folio que ya está usado o impreso físicamente?",
            initialvalue=actual,
            minvalue=0,
            maxvalue=999999,
            parent=self
        )
        if nuevo is not None:
            self.engine.set_ultimo_folio(codigo, nuevo, "Sincronización manual con planillero físico")
            self._log(f"Folio de {codigo} sincronizado a {nuevo}. El próximo correlativo será {nuevo + 1}.")
            self._actualizar_labels_checkboxes()
            self._al_cambiar_seleccion_checkboxes()

    def _generar_planillas(self):
        codigos_seleccionados = [cod for cod, var in self.check_vars.items() if var.get()]
        if not codigos_seleccionados:
            messagebox.showwarning("Atención", "Por favor seleccioná al menos una planilla marcando su casilla.", parent=self)
            return

        try:
            cant = int(self.var_cantidad.get())
            if cant < 1:
                raise ValueError()
        except Exception:
            messagebox.showerror("Error", "La cantidad debe ser un número entero mayor a 0.", parent=self)
            return

        auto_print = self.var_auto_imprimir.get()
        doble_cara = self.var_doble_cara.get()

        if auto_print and not self.engine.hay_impresora_disponible():
            messagebox.showwarning(
                "Sin Impresora",
                "No hay una impresora configurada. Se generarán los archivos Word en disco sin enviarlos a imprimir.",
                parent=self
            )
            auto_print = False
            self.var_auto_imprimir.set(False)

        # Desactivar botones durante el procesamiento para evitar dobles clics
        self.btn_generar.config(state="disabled", text="⏳ Generando planillas...")
        self.btn_imprimir_lote.config(state="disabled")

        def worker():
            try:
                total_planillas = len(codigos_seleccionados)
                self._log(f"Iniciando emisión de {total_planillas} tipo(s) de planilla ({cant} hoja(s) c/u)...")
                resultados = []
                todos_archivos = []
                errores_impresion = []

                for idx, cod in enumerate(codigos_seleccionados, 1):
                    self._log(f"[{idx}/{total_planillas}] Generando {cant} hoja(s) de {cod}...")
                    res = self.engine.generar_planillas(
                        cod,
                        cant,
                        imprimir_al_generar=auto_print,
                        doble_cara=doble_cara
                    )
                    resultados.append(res)
                    todos_archivos.extend([a["archivo"] for a in res["archivos"]])
                    if res.get("error_impresion"):
                        errores_impresion.append(f"{cod}: {res['error_impresion']}")

                self.after(0, lambda: self._on_generacion_multiple_completada(resultados, todos_archivos, auto_print, errores_impresion))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._on_generacion_error(err_msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_generacion_multiple_completada(self, resultados, todos_archivos, auto_print, errores_impresion):
        self.btn_generar.config(state="normal", text="🚀 GENERAR PLANILLAS FOLIADAS")
        self._actualizar_estado_impresora()
        self.ultimos_archivos_generados = todos_archivos
        if self.engine.hay_impresora_disponible() and self.ultimos_archivos_generados:
            self.btn_imprimir_lote.config(state="normal")

        # Actualizar etiquetas de folios en los checkboxes
        self._actualizar_labels_checkboxes()
        self._al_cambiar_seleccion_checkboxes()

        total_hojas = len(todos_archivos)
        num_planillas = len(resultados)

        estado_imp_texto = "Guardado listo en carpeta"
        if auto_print:
            if not errores_impresion:
                estado_imp_texto = "Enviado a impresora con éxito"
            else:
                estado_imp_texto = f"⚠️ Guardado en carpeta (Errores de impresión en {len(errores_impresion)} planilla(s))"

        self._log(f"✅ Completada emisión de {num_planillas} planilla(s) · Total: {total_hojas} archivo(s) Word.")
        if auto_print and not errores_impresion:
            self._log(f"🖨️ {total_hojas} documento(s) enviados al spooler de {self.engine.get_impresora_predeterminada()}.")
        elif errores_impresion:
            for err in errores_impresion:
                self._log(f"⚠️ {err}")

        desglose = "\n".join([
            f"• {r['codigo']}: Folios {r['folio_inicial']} al {r['folio_final']} ({r['cantidad']} hojas)"
            for r in resultados
        ])

        msg_exito = (
            f"✅ ¡Planillas generadas con éxito!\n\n"
            f"• Planillas procesadas: {num_planillas}\n"
            f"• Total emitido: {total_hojas} archivo(s) Word (.docx)\n\n"
            f"• Detalle de rangos:\n{desglose}\n\n"
            f"• Logos y formato: 100% Preservados intactos\n"
            f"• Impresión: {estado_imp_texto}\n\n"
            f"¿Deseás abrir la carpeta de salida ahora?"
        )

        abrir = messagebox.askyesno("Emisión Completada", msg_exito, parent=self)
        if abrir and resultados:
            self._abrir_carpeta(resultados[0]["carpeta_salida"])

    def _on_generacion_error(self, error_str):
        self.btn_generar.config(state="normal", text="🚀 GENERAR PLANILLAS FOLIADAS")
        self._actualizar_estado_impresora()
        self._log(f"❌ Error en generación: {error_str}")
        messagebox.showerror("Error al Generar", f"Ocurrió un error inesperado:\n{error_str}", parent=self)

    def _imprimir_ultimo_lote(self):
        if not self.ultimos_archivos_generados:
            messagebox.showinfo("Imprimir", "Aún no se ha generado ningún lote en esta sesión.", parent=self)
            return

        if not self.engine.hay_impresora_disponible():
            messagebox.showwarning(
                "Sin Impresora",
                "No hay ninguna impresora activa o configurada en Windows.\nVerifique la conexión antes de imprimir.",
                parent=self
            )
            return

        impresora_actual = self.engine.get_impresora_predeterminada()
        confirm = messagebox.askyesno(
            "Confirmar Impresión",
            f"Se enviarán {len(self.ultimos_archivos_generados)} archivo(s) Word a la impresora:\n"
            f"{impresora_actual}\n\n"
            f"¿Deseás continuar?",
            parent=self
        )
        if not confirm:
            return

        doble_cara = self.var_doble_cara.get()
        total = len(self.ultimos_archivos_generados)

        self.btn_imprimir_lote.config(state="disabled", text="⏳ Imprimiendo...")
        self.btn_generar.config(state="disabled")

        def worker():
            try:
                self.after(0, lambda: self._log(f"Iniciando impresión de {total} documento(s) (Dúplex: {'Sí' if doble_cara else 'No'})..."))
                self.engine.imprimir_lote(
                    self.ultimos_archivos_generados,
                    doble_cara=doble_cara,
                    callback_progreso=lambda i, tot, nom: self.after(0, lambda: self._log(f"🖨️ [{i}/{tot}] Enviado a spooler: {nom}"))
                )
                self.after(0, self._on_impresion_completada)
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: self._on_impresion_error(err_msg))

        threading.Thread(target=worker, daemon=True).start()

    def _on_impresion_completada(self):
        self.btn_imprimir_lote.config(state="normal", text="🖨️ Imprimir Último Lote")
        self.btn_generar.config(state="normal")
        total = len(self.ultimos_archivos_generados)
        self._log(f"✅ Se enviaron {total} hojas al spooler de la impresora.")
        messagebox.showinfo("Impresión Enviada", f"Se enviaron {total} hojas a la impresora.", parent=self)

    def _on_impresion_error(self, error_str):
        self.btn_imprimir_lote.config(state="normal", text="🖨️ Imprimir Último Lote")
        self.btn_generar.config(state="normal")
        self._log(f"❌ Error de impresión: {error_str}")
        messagebox.showerror("Error de Impresión", f"No se pudo completar la impresión:\n{error_str}", parent=self)

    def _abrir_carpeta_salida(self):
        self._abrir_carpeta(self.engine.salida_dir)

    def _abrir_carpeta(self, ruta):
        if sys.platform == "win32":
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.run(["open", ruta])
        else:
            subprocess.run(["xdg-open", ruta])


def main():
    app = FoliadorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
