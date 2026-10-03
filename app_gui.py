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
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.folio_engine import FolioEngine


class FoliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.engine = FolioEngine()

        self.title("ColDev · Foliador Oficial PAC / POS (Preservación Exacta Word)")
        self.geometry("820x720")
        self.minsize(760, 660)
        self.configure(bg="#0f172a")

        # Configurar estilos de ttk
        self._configurar_estilos()

        # Variables de estado
        self.var_codigo = tk.StringVar()
        self.var_cantidad = tk.IntVar(value=1)
        self.var_ultimo_folio = tk.StringVar()
        self.var_proximo_folio = tk.StringVar()
        self.var_resumen_rango = tk.StringVar()
        self.var_auto_imprimir = tk.BooleanVar(value=False)
        self.var_doble_cara = tk.BooleanVar(value=True)
        self.ultimos_archivos_generados = []

        # Mapa de opciones para el selector
        self.opciones_dict = {}

        # Construir UI
        self._construir_ui()

        # Cargar primera plantilla por defecto
        if self.engine.plantillas:
            self.combo_plantillas.current(0)
            self._al_cambiar_plantilla()

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

        # ESTILO DEL SELECTOR (Alto Contraste y Letra Negrita Ultra Visible)
        style.configure("HighContrast.TCombobox",
            fieldbackground="#ffffff",
            background="#0284c7",
            foreground="#0f172a",
            arrowcolor="#0f172a",
            bordercolor="#38bdf8",
            padding=6,
            font=("Segoe UI", 11, "bold")
        )
        style.map("HighContrast.TCombobox",
            fieldbackground=[("readonly", "#ffffff")],
            foreground=[("readonly", "#0f172a")],
            selectbackground=[("readonly", "#38bdf8")],
            selectforeground=[("readonly", "#0f172a")]
        )

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

        # 2. SELECCIÓN DE PLANILLA (CARD)
        card_sel = ttk.Frame(main_frame, style="Card.TFrame", padding=14)
        card_sel.pack(fill="x", pady=(0, 12))

        ttk.Label(card_sel, text="1. SELECCIONAR PLANILLA", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(anchor="w")
        ttk.Label(card_sel, text="Elegí la planilla oficial que necesitás emitir:", style="Muted.TLabel").pack(anchor="w", pady=(2, 6))

        # Opciones limpias y bien categorizadas
        opciones_display = []
        for p in self.engine.plantillas:
            display_name = f"{p['categoria']} · {p['nombre']}"
            opciones_display.append(display_name)
            self.opciones_dict[display_name] = p["codigo"]

        self.combo_plantillas = ttk.Combobox(
            card_sel,
            values=opciones_display,
            state="readonly",
            style="HighContrast.TCombobox",
            height=13
        )
        self.combo_plantillas.pack(fill="x", pady=(0, 2))
        self.combo_plantillas.bind("<<ComboboxSelected>>", self._al_cambiar_plantilla)

        # 3. ESTADO DEL FOLIO Y CANTIDAD (CARD)
        card_folio = ttk.Frame(main_frame, style="Card.TFrame", padding=14)
        card_folio.pack(fill="x", pady=(0, 12))

        ttk.Label(card_folio, text="2. ESTADO DEL PLANILLERO Y CORRELATIVO", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 8))

        grid_frame = ttk.Frame(card_folio, style="Card.TFrame")
        grid_frame.pack(fill="x")

        # Columna 1: Último folio
        f1 = ttk.Frame(grid_frame, style="Card.TFrame")
        f1.grid(row=0, column=0, sticky="w", padx=(0, 20))

        ttk.Label(f1, text="Último Folio Registrado:", style="Muted.TLabel").pack(anchor="w")
        sub_f1 = ttk.Frame(f1, style="Card.TFrame")
        sub_f1.pack(anchor="w", pady=(4, 0))
        ttk.Label(sub_f1, textvariable=self.var_ultimo_folio, style="Warning.TLabel").pack(side="left", padx=(0, 8))
        btn_modificar = ttk.Button(sub_f1, text="✏️ Sincronizar Físico", style="Small.TButton", command=self._sincronizar_folio_fisico)
        btn_modificar.pack(side="left")

        # Columna 2: Cantidad a emitir
        f2 = ttk.Frame(grid_frame, style="Card.TFrame")
        f2.grid(row=0, column=1, sticky="w", padx=(0, 20))

        ttk.Label(f2, text="Cantidad de Hojas:", style="Muted.TLabel").pack(anchor="w")
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

        # Columna 3: Próximo folio
        f3 = ttk.Frame(grid_frame, style="Card.TFrame")
        f3.grid(row=0, column=2, sticky="w")

        ttk.Label(f3, text="Folios a Emitir (Centrados):", style="Muted.TLabel").pack(anchor="w")
        ttk.Label(f3, textvariable=self.var_resumen_rango, style="Highlight.TLabel").pack(anchor="w", pady=(4, 0))

        # 4. CONFIGURACIÓN DE IMPRESIÓN (CARD)
        card_print = ttk.Frame(main_frame, style="Card.TFrame", padding=14)
        card_print.pack(fill="x", pady=(0, 12))

        ttk.Label(card_print, text="3. SALIDA E IMPRESIÓN DIRECTA", font=("Segoe UI", 11, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 4))

        impresora_actual = self.engine.get_impresora_predeterminada()
        ttk.Label(
            card_print,
            text=f"🖨️ Impresora detectada: {impresora_actual}",
            style="Muted.TLabel"
        ).pack(anchor="w", pady=(0, 6))

        chk_duplex = ttk.Checkbutton(
            card_print,
            text="📄 Imprimir por ambos lados (Dúplex: Planilla adelante + Instructivo atrás en 1 sola hoja)",
            variable=self.var_doble_cara
        )
        chk_duplex.pack(anchor="w", pady=(0, 4))

        chk_print = ttk.Checkbutton(
            card_print,
            text="⚡ Mandar a imprimir automáticamente a la impresora al hacer clic en Generar",
            variable=self.var_auto_imprimir
        )
        chk_print.pack(anchor="w")

        # 5. BOTONES DE ACCIÓN
        btn_bar = ttk.Frame(main_frame)
        btn_bar.pack(fill="x", pady=(2, 12))

        btn_generar = ttk.Button(
            btn_bar,
            text="🚀 GENERAR PLANILLAS FOLIADAS",
            style="Primary.TButton",
            command=self._generar_planillas
        )
        btn_generar.pack(side="left", fill="x", expand=True, padx=(0, 8))

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
        self.txt_log.insert("end", f"• {mensaje}\n")
        self.txt_log.see("end")

    def _obtener_codigo_actual(self):
        val = self.combo_plantillas.get()
        return self.opciones_dict.get(val, None)

    def _al_cambiar_plantilla(self, event=None):
        codigo = self._obtener_codigo_actual()
        if not codigo:
            return

        info = self.engine.get_plantilla_info(codigo)
        ultimo = self.engine.get_ultimo_folio(codigo)
        dig = info.get("digitos", 5)

        self.var_ultimo_folio.set(f"{ultimo:0{dig}d}")
        self._actualizar_calculo_rango()

    def _sumar_cantidad(self, delta):
        nueva = max(1, self.var_cantidad.get() + delta)
        self.var_cantidad.set(nueva)
        self._actualizar_calculo_rango()

    def _actualizar_calculo_rango(self):
        codigo = self._obtener_codigo_actual()
        if not codigo:
            return
        info = self.engine.get_plantilla_info(codigo)
        ultimo = self.engine.get_ultimo_folio(codigo)
        dig = info.get("digitos", 5)

        try:
            cant = int(self.var_cantidad.get())
            if cant < 1:
                cant = 1
        except Exception:
            cant = 1

        ini = ultimo + 1
        fin = ultimo + cant

        if cant == 1:
            self.var_resumen_rango.set(f"{ini:0{dig}d} (1 hoja centrada)")
        else:
            self.var_resumen_rango.set(f"{ini:0{dig}d} al {fin:0{dig}d} ({cant} hojas)")

    def _sincronizar_folio_fisico(self):
        codigo = self._obtener_codigo_actual()
        if not codigo:
            return
        info = self.engine.get_plantilla_info(codigo)
        actual = self.engine.get_ultimo_folio(codigo)

        nuevo = simpledialog.askinteger(
            "Sincronizar con Planillero Físico",
            f"Planilla: {info['nombre']}\n\n"
            f"El último folio en el sistema es: {actual}\n"
            f"¿Cuál es el último folio que ya está usado o impreso físicamente?",
            initialvalue=actual,
            minvalue=0,
            maxvalue=999999,
            parent=self
        )

        if nuevo is not None:
            self.engine.set_ultimo_folio(codigo, nuevo, "Sincronización manual con planillero físico")
            self._log(f"Folio de {codigo} sincronizado a {nuevo}. El próximo correlativo será {nuevo + 1}.")
            self._al_cambiar_plantilla()

    def _generar_planillas(self):
        codigo = self._obtener_codigo_actual()
        if not codigo:
            messagebox.showerror("Error", "Por favor seleccioná una planilla.")
            return

        try:
            cant = int(self.var_cantidad.get())
            if cant < 1:
                raise ValueError()
        except Exception:
            messagebox.showerror("Error", "La cantidad debe ser un número entero mayor a 0.")
            return

        info = self.engine.get_plantilla_info(codigo)
        auto_print = self.var_auto_imprimir.get()
        doble_cara = self.var_doble_cara.get()

        try:
            res = self.engine.generar_planillas(
                codigo,
                cant,
                imprimir_al_generar=auto_print,
                doble_cara=doble_cara
            )
            self.ultimos_archivos_generados = [a["archivo"] for a in res["archivos"]]
            self._al_cambiar_plantilla()

            msg_exito = (
                f"✅ ¡Planillas generadas con éxito!\n\n"
                f"• Planilla: {res['nombre']}\n"
                f"• Rango de Folios: {res['folio_inicial']} al {res['folio_final']} (centrados)\n"
                f"• Total emitido: {res['cantidad']} archivo(s) Word (.docx)\n"
                f"• Logos y formato: 100% Preservados intactos\n"
                f"• Impresión: {'Enviado a impresora' if auto_print else 'Guardado listo en carpeta'}\n\n"
                f"¿Deseás abrir la carpeta de salida ahora?"
            )

            self._log(f"Generadas {res['cantidad']} hojas de {codigo} [Folios {res['folio_inicial']} al {res['folio_final']}].")
            if auto_print:
                self._log(f"🖨️ {res['cantidad']} documento(s) enviados a imprimir a {self.engine.get_impresora_predeterminada()}.")

            abrir = messagebox.askyesno("Emisión Completada", msg_exito, parent=self)
            if abrir:
                self._abrir_carpeta(res["carpeta_salida"])

        except Exception as e:
            messagebox.showerror("Error al Generar", f"Ocurrió un error inesperado:\n{e}", parent=self)

    def _imprimir_ultimo_lote(self):
        if not self.ultimos_archivos_generados:
            messagebox.showinfo("Imprimir", "Aún no se ha generado ningún lote en esta sesión.")
            return

        confirm = messagebox.askyesno(
            "Confirmar Impresión",
            f"Se enviarán {len(self.ultimos_archivos_generados)} archivo(s) Word a la impresora:\n"
            f"{self.engine.get_impresora_predeterminada()}\n\n"
            f"¿Deseás continuar?",
            parent=self
        )
        if not confirm:
            return

        try:
            doble_cara = self.var_doble_cara.get()
            for ruta in self.ultimos_archivos_generados:
                self.engine.imprimir_documento(ruta, doble_cara=doble_cara)
                self._log(f"🖨️ Impreso (dúplex: {'Sí' if doble_cara else 'No'}): {os.path.basename(ruta)}")
            messagebox.showinfo("Impresión Enviada", f"Se enviaron {len(self.ultimos_archivos_generados)} hojas a la impresora.")
        except Exception as e:
            messagebox.showerror("Error de Impresión", f"No se pudo completar la impresión:\n{e}")

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
