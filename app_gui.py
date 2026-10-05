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
        self.geometry("960x670")
        self.minsize(860, 560)
        self.configure(bg="#0f172a")

        # Configurar estilos de ttk
        self._configurar_estilos()

        # Variables de estado
        self.var_cantidad = tk.IntVar(value=1)
        self.var_ultimo_folio = tk.StringVar()
        self.var_resumen_rango = tk.StringVar()
        self.var_impresora_seleccionada = tk.StringVar()
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

        # Personalizar selector desplegable (Combobox)
        style.configure("TCombobox", fieldbackground="#1e293b", background="#334155", foreground="#ffffff", selectbackground="#0284c7", selectforeground="#ffffff", font=("Segoe UI", 10))
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#1e293b")],
                  foreground=[("readonly", "#ffffff")],
                  selectbackground=[("readonly", "#0284c7")],
                  selectforeground=[("readonly", "#ffffff")])

        # Personalizar pestañas (Notebook)
        style.configure("TNotebook", background="#0f172a", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1e293b", foreground="#94a3b8", font=("Segoe UI", 10, "bold"), padding=[12, 5])
        style.map("TNotebook.Tab",
                  background=[("selected", "#0284c7"), ("active", "#334155")],
                  foreground=[("selected", "#ffffff"), ("active", "#f8fafc")])

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
        # Contenedor principal
        main_frame = ttk.Frame(self, padding=(20, 0, 20, 12))
        main_frame.pack(fill="both", expand=True)

        # 2. SELECCIÓN DE PLANILLAS POR LÍNEAS (CARD CON PESTAÑAS)
        card_sel = ttk.Frame(main_frame, style="Card.TFrame", padding=(12, 8))
        card_sel.pack(fill="x", pady=(0, 8))

        # Barra superior con título y botones de selección rápida
        hdr_sel = ttk.Frame(card_sel, style="Card.TFrame")
        hdr_sel.pack(fill="x", pady=(0, 4))

        ttk.Label(
            hdr_sel,
            text="1. SELECCIÓN DE PLANILLAS POR LÍNEA (Puedes marcar una o varias a la vez)",
            font=("Segoe UI", 10, "bold"),
            style="Card.TLabel"
        ).pack(side="left")

        btn_box = ttk.Frame(hdr_sel, style="Card.TFrame")
        btn_box.pack(side="right")
        ttk.Button(btn_box, text="✅ Todas", style="ActionText.TButton", command=self._seleccionar_todas).pack(side="left", padx=2)
        ttk.Button(btn_box, text="⬜ Ninguna", style="ActionText.TButton", command=self._deseleccionar_todas).pack(side="left", padx=2)
        ttk.Button(btn_box, text="🦑 Cefalópodos", style="ActionText.TButton", command=lambda: self._seleccionar_linea("cefalopodos")).pack(side="left", padx=2)
        ttk.Button(btn_box, text="🐟 Pescados", style="ActionText.TButton", command=lambda: self._seleccionar_linea("pescados")).pack(side="left", padx=2)
        ttk.Button(btn_box, text="🧹 POS", style="ActionText.TButton", command=lambda: self._seleccionar_linea("pos")).pack(side="left", padx=2)

        # Notebook de pestañas por línea
        self.notebook_lineas = ttk.Notebook(card_sel)
        self.notebook_lineas.pack(fill="x", pady=(2, 0))

        tab_cefalopodos = ttk.Frame(self.notebook_lineas, style="Card.TFrame", padding=(8, 6))
        tab_pescados = ttk.Frame(self.notebook_lineas, style="Card.TFrame", padding=(8, 6))
        tab_pos = ttk.Frame(self.notebook_lineas, style="Card.TFrame", padding=(8, 6))
        tab_todas = ttk.Frame(self.notebook_lineas, style="Card.TFrame", padding=(8, 6))

        self.notebook_lineas.add(tab_cefalopodos, text="🦑 Línea Cefalópodos (8)")
        self.notebook_lineas.add(tab_pescados, text="🐟 Línea Pescados (4)")
        self.notebook_lineas.add(tab_pos, text="🧹 Saneamiento POS (4)")
        self.notebook_lineas.add(tab_todas, text="📋 Todas las Planillas (12)")

        # Inicializar variables y mapa de widgets
        for p in self.engine.plantillas:
            cod = p["codigo"]
            if cod not in self.check_vars:
                self.check_vars[cod] = tk.BooleanVar(value=False)
            self.check_widgets[cod] = []

        def poblar_pestana(tab_parent, lista_filtrada, mensaje_info=None):
            f_grid = ttk.Frame(tab_parent, style="Card.TFrame")
            f_grid.pack(fill="x")
            f_grid.columnconfigure(0, weight=1)
            f_grid.columnconfigure(1, weight=1)

            mitad = (len(lista_filtrada) + 1) // 2
            for idx, p in enumerate(lista_filtrada):
                col = 0 if idx < mitad else 1
                row = idx if idx < mitad else (idx - mitad)
                cod = p["codigo"]
                var = self.check_vars[cod]
                txt = self._obtener_texto_checkbox(p)
                chk = ttk.Checkbutton(
                    f_grid,
                    text=txt,
                    variable=var,
                    command=self._al_cambiar_seleccion_checkboxes
                )
                chk.grid(row=row, column=col, sticky="w", padx=6, pady=1)
                self.check_widgets[cod].append(chk)

            if mensaje_info:
                ttk.Label(
                    tab_parent,
                    text=mensaje_info,
                    style="Muted.TLabel",
                    font=("Segoe UI", 8, "italic")
                ).pack(anchor="w", pady=(4, 0), padx=6)

        # Filtrar plantillas según línea configurada
        plantillas_cefalopodos = [p for p in self.engine.plantillas if p.get("linea", "").lower() in ("cefalopodos", "ambas")]
        plantillas_pescados = [p for p in self.engine.plantillas if p.get("linea", "").lower() in ("pescados", "ambas")]
        plantillas_pos = [p for p in self.engine.plantillas if p.get("linea", "").lower() == "pos" or p.get("categoria") == "POS"]

        poblar_pestana(tab_cefalopodos, plantillas_cefalopodos)
        poblar_pestana(
            tab_pescados,
            plantillas_pescados,
            "ℹ️ Línea Pescados (Merluza / Congrio / Reineta). Próximamente se integrarán las plantillas de recepción específicas."
        )
        poblar_pestana(tab_pos, plantillas_pos)
        poblar_pestana(tab_todas, self.engine.plantillas)

        # 3. ESTADO DEL FOLIO Y CANTIDAD (CARD)
        card_folio = ttk.Frame(main_frame, style="Card.TFrame", padding=(12, 6))
        card_folio.pack(fill="x", pady=(0, 8))

        grid_frame = ttk.Frame(card_folio, style="Card.TFrame")
        grid_frame.pack(fill="x")

        # Columna 1: Folio actual
        f1 = ttk.Frame(grid_frame, style="Card.TFrame")
        f1.grid(row=0, column=0, sticky="w", padx=(0, 24))
        ttk.Label(f1, text="Planilla / Folio Actual:", style="Muted.TLabel").pack(anchor="w")
        sub_f1 = ttk.Frame(f1, style="Card.TFrame")
        sub_f1.pack(anchor="w", pady=(2, 0))
        ttk.Label(sub_f1, textvariable=self.var_ultimo_folio, style="Warning.TLabel").pack(side="left", padx=(0, 8))
        ttk.Button(sub_f1, text="✏️ Sincronizar Físico", style="Small.TButton", command=self._sincronizar_folio_fisico).pack(side="left")

        # Columna 2: Cantidad por planilla
        f2 = ttk.Frame(grid_frame, style="Card.TFrame")
        f2.grid(row=0, column=1, sticky="w", padx=(0, 24))
        ttk.Label(f2, text="Cantidad por Planilla:", style="Muted.TLabel").pack(anchor="w")
        sub_f2 = ttk.Frame(f2, style="Card.TFrame")
        sub_f2.pack(anchor="w", pady=(2, 0))
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

        # Columna 3: Resumen
        f3 = ttk.Frame(grid_frame, style="Card.TFrame")
        f3.grid(row=0, column=2, sticky="w")
        ttk.Label(f3, text="Total a Emitir (Centrados):", style="Muted.TLabel").pack(anchor="w")
        ttk.Label(f3, textvariable=self.var_resumen_rango, style="Highlight.TLabel").pack(anchor="w", pady=(2, 0))

        # 4. CONFIGURACIÓN DE IMPRESIÓN (CARD)
        card_print = ttk.Frame(main_frame, style="Card.TFrame", padding=(12, 8))
        card_print.pack(fill="x", pady=(0, 8))

        f_printer_top = ttk.Frame(card_print, style="Card.TFrame")
        f_printer_top.pack(fill="x", pady=(0, 4))

        ttk.Label(f_printer_top, text="🖨️ Impresora destino:", font=("Segoe UI", 10, "bold"), style="Card.TLabel").pack(side="left", padx=(0, 8))

        self.combo_impresoras = ttk.Combobox(
            f_printer_top,
            textvariable=self.var_impresora_seleccionada,
            state="readonly",
            font=("Segoe UI", 10),
            width=36
        )
        self.combo_impresoras.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.combo_impresoras.bind("<<ComboboxSelected>>", self._on_cambio_impresora)

        self.btn_recargar_imp = ttk.Button(
            f_printer_top,
            text="🔄 Actualizar",
            style="Secondary.TButton",
            command=self._actualizar_estado_impresora
        )
        self.btn_recargar_imp.pack(side="left", padx=(0, 4))

        self.btn_fijar_predeterminada = ttk.Button(
            f_printer_top,
            text="📌 Fijar por defecto",
            style="Secondary.TButton",
            command=self._fijar_impresora_como_predeterminada
        )
        self.btn_fijar_predeterminada.pack(side="left")

        # Fila 2 de impresión: Estado de impresora y opciones dúplex / auto-print
        f_printer_sub = ttk.Frame(card_print, style="Card.TFrame")
        f_printer_sub.pack(fill="x", pady=(2, 0))

        self.lbl_impresora = ttk.Label(
            f_printer_sub,
            text="Buscando impresoras del sistema...",
            style="Muted.TLabel"
        )
        self.lbl_impresora.pack(side="left", padx=(0, 16))

        self.chk_duplex = ttk.Checkbutton(
            f_printer_sub,
            text="📄 Imprimir por ambos lados (Dúplex)",
            variable=self.var_doble_cara
        )
        self.chk_duplex.pack(side="left", padx=(0, 16))

        self.chk_print = ttk.Checkbutton(
            f_printer_sub,
            text="⚡ Mandar a imprimir automáticamente al hacer clic en Generar",
            variable=self.var_auto_imprimir
        )
        self.chk_print.pack(side="left")

        # 5. BOTONES DE ACCIÓN (SIEMPRE VISIBLES Y DESTACADOS)
        btn_bar = ttk.Frame(main_frame)
        btn_bar.pack(fill="x", pady=(2, 8))

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
        card_log = ttk.Frame(main_frame, style="Card.TFrame", padding=(12, 6))
        card_log.pack(fill="both", expand=True)

        ttk.Label(card_log, text="REGISTRO DE ACTIVIDAD:", font=("Segoe UI", 9, "bold"), style="Card.TLabel").pack(anchor="w")

        self.txt_log = tk.Text(
            card_log,
            bg="#0b1120",
            fg="#94a3b8",
            insertbackground="#ffffff",
            font=("Consolas", 9),
            height=3,
            relief="flat",
            wrap="word"
        )
        self.txt_log.pack(fill="both", expand=True, pady=(2, 0))
        self._log("Listo. Selecciona una planilla y haz clic en Generar.")

    def _log(self, mensaje):
        def do_log():
            self.txt_log.insert("end", f"• {mensaje}\n")
            self.txt_log.see("end")
        self.after(0, do_log)

    def _actualizar_estado_impresora(self):
        impresoras = self.engine.get_impresoras_disponibles()
        guardada = self.engine.get_impresora_guardada()
        predeterminada = self.engine.get_impresora_predeterminada()

        if impresoras:
            self.combo_impresoras["values"] = impresoras
            seleccion_actual = self.var_impresora_seleccionada.get()

            # Prioridad de selección:
            # 1. Si la selección actual en UI es válida en el sistema, conservarla
            # 2. Si la guardada previamente por el usuario existe, restaurarla
            # 3. Si la predeterminada de Windows existe, seleccionarla
            # 4. Primera impresora de la lista
            if seleccion_actual in impresoras:
                seleccionada = seleccion_actual
            elif guardada in impresoras:
                seleccionada = guardada
            elif predeterminada in impresoras:
                seleccionada = predeterminada
            else:
                seleccionada = impresoras[0]

            self.var_impresora_seleccionada.set(seleccionada)
            self.engine.guardar_impresora_guardada(seleccionada)

            tag_defecto = " [Predeterminada del sistema]" if (predeterminada and seleccionada == predeterminada) else ""
            self.lbl_impresora.config(
                text=f"🖨️ Conectada: {seleccionada}{tag_defecto}",
                style="Muted.TLabel"
            )
            self.chk_print.config(state="normal")
            if self.ultimos_archivos_generados:
                self.btn_imprimir_lote.config(state="normal")
            else:
                self.btn_imprimir_lote.config(state="disabled")
        else:
            self.combo_impresoras["values"] = []
            self.var_impresora_seleccionada.set("")
            self.lbl_impresora.config(
                text="⚠️ No se detectó ninguna impresora activa en Windows (Impresión desactivada)",
                style="Danger.TLabel"
            )
            self.var_auto_imprimir.set(False)
            self.chk_print.config(state="disabled")
            self.btn_imprimir_lote.config(state="disabled")

    def _on_cambio_impresora(self, event=None):
        seleccionada = self.var_impresora_seleccionada.get()
        if seleccionada:
            self.engine.guardar_impresora_guardada(seleccionada)
            predeterminada = self.engine.get_impresora_predeterminada()
            tag_defecto = " [Predeterminada del sistema]" if (predeterminada and seleccionada == predeterminada) else ""
            self.lbl_impresora.config(
                text=f"🖨️ Conectada: {seleccionada}{tag_defecto}",
                style="Muted.TLabel"
            )
            self._log(f"🖨️ Impresora destino seleccionada: {seleccionada}")

    def _fijar_impresora_como_predeterminada(self):
        seleccionada = self.var_impresora_seleccionada.get()
        if not seleccionada:
            messagebox.showwarning("Atención", "No hay ninguna impresora seleccionada.", parent=self)
            return
        try:
            exito = self.engine.set_impresora_predeterminada_sistema(seleccionada)
            if exito:
                self._log(f"📌 '{seleccionada}' fijada como impresora predeterminada de Windows.")
                self._actualizar_estado_impresora()
                messagebox.showinfo(
                    "Impresora Predeterminada",
                    f"'{seleccionada}' se configuró con éxito como la impresora predeterminada del sistema Windows.",
                    parent=self
                )
            else:
                messagebox.showerror("Error", "No se pudo cambiar la impresora predeterminada de Windows.", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al establecer impresora predeterminada: {e}", parent=self)

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
                txt = self._obtener_texto_checkbox(p)
                widgets = self.check_widgets[cod]
                if isinstance(widgets, list):
                    for w in widgets:
                        w.config(text=txt)
                else:
                    widgets.config(text=txt)

    def _seleccionar_todas(self):
        for var in self.check_vars.values():
            var.set(True)
        self._al_cambiar_seleccion_checkboxes()

    def _deseleccionar_todas(self):
        for var in self.check_vars.values():
            var.set(False)
        self._al_cambiar_seleccion_checkboxes()

    def _seleccionar_linea(self, linea_deseada):
        for p in self.engine.plantillas:
            cod = p["codigo"]
            linea = p.get("linea", p.get("categoria", "")).lower()
            if linea_deseada == "cefalopodos" and linea in ("cefalopodos", "ambas"):
                self.check_vars[cod].set(True)
            elif linea_deseada == "pescados" and linea in ("pescados", "ambas"):
                self.check_vars[cod].set(True)
            elif linea_deseada == "pos" and (linea == "pos" or p.get("categoria") == "POS"):
                self.check_vars[cod].set(True)
            else:
                self.check_vars[cod].set(False)
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
            self.var_resumen_rango.set("Marca al menos 1 planilla")
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

        ttk.Label(ventana_sync, text="Elige la planilla que deseas sincronizar:", style="Card.TLabel", font=("Segoe UI", 10, "bold")).pack(pady=(16, 6), padx=16, anchor="w")

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
            messagebox.showwarning("Atención", "Por favor selecciona al menos una planilla marcando su casilla.", parent=self)
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
                    impresora_destino = self.var_impresora_seleccionada.get()
                    res = self.engine.generar_planillas(
                        cod,
                        cant,
                        imprimir_al_generar=auto_print,
                        doble_cara=doble_cara,
                        nombre_impresora=impresora_destino
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
        imp_nombre = self.var_impresora_seleccionada.get() or self.engine.get_impresora_predeterminada()
        if auto_print and not errores_impresion:
            self._log(f"🖨️ {total_hojas} documento(s) enviados al spooler de {imp_nombre}.")
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
            f"¿Deseas abrir la carpeta de salida ahora?"
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

        impresora_actual = self.var_impresora_seleccionada.get() or self.engine.get_impresora_predeterminada()
        confirm = messagebox.askyesno(
            "Confirmar Impresión",
            f"Se enviarán {len(self.ultimos_archivos_generados)} archivo(s) Word a la impresora:\n"
            f"{impresora_actual}\n\n"
            f"¿Deseas continuar?",
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
                self.after(0, lambda: self._log(f"Iniciando impresión de {total} documento(s) en '{impresora_actual}' (Dúplex: {'Sí' if doble_cara else 'No'})..."))
                self.engine.imprimir_lote(
                    self.ultimos_archivos_generados,
                    doble_cara=doble_cara,
                    nombre_impresora=impresora_actual,
                    callback_progreso=lambda i, tot, nom: self.after(0, lambda: self._log(f"🖨️ [{i}/{tot}] Enviado a spooler ({impresora_actual}): {nom}"))
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
        imp_act = self.var_impresora_seleccionada.get() or "la impresora"
        self._log(f"✅ Se enviaron {total} hojas al spooler de {imp_act}.")
        messagebox.showinfo("Impresión Enviada", f"Se enviaron {total} hojas a {imp_act}.", parent=self)

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
