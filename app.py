import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd
from informes_iva import GeneradorInformeIVA
import tempfile
from PIL import Image, ImageTk

from config import (
    APP_NAME,
    APP_VERSION,
    DEFAULT_SHEET_NAMES,
    LOG_FILE,
    OPTIONAL_SHEETS,
    SHEET_LABELS,
    SHEET_ORDER,
    COLUMNAS_EXCLUIDAS_AUD_COMP,
    CATEGORY_COLORS,
)
from errors import ErrorSistema, ErrorUsuario, logger
from excel_engine import ConciliadorAuditoria
from preview_widget import VistaPreviaExcel
from informes_iva import GeneradorInformeIVA
from token_engine import FormateadorToken



PALETTE = {
    "bg": "#E9ECF2",
    "surface": "#FFFFFF",
    "surface_alt": "#F3F5F9",
    "border": "#D3D8E0",
    "border_strong": "#B9C0CC",
    "text": "#111827",
    "text_muted": "#5B6472",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_active": "#1E40AF",
    "primary_fg": "#FFFFFF",
    "success": "#16A34A",
    "danger": "#DC2626",
    "console_bg": "#12141A",
    "console_fg": "#E5E7EB",
}


class ScrollableChecklist(ttk.Frame):
    """Contenedor con scroll vertical para listas largas de checkboxes."""

    def __init__(self, parent, height=170, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self._xls_actual=None
        self.canvas = tk.Canvas(
            self, bg=PALETTE["surface"], highlightthickness=0, bd=0, height=height
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="Card.TFrame")

        self._scroll_job = None
        self.inner.bind("<Configure>", self._on_inner_configure)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(self._window, width=e.width)
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.canvas.bind_all("<Button-4>", self._on_wheel, add="+")
        self.canvas.bind_all("<Button-5>", self._on_wheel, add="+")

    def _on_inner_configure(self, _event):
        if self._scroll_job is not None:
            self.after_cancel(self._scroll_job)
        self._scroll_job = self.after(30, self._actualizar_scrollregion)

    def _actualizar_scrollregion(self):
        self._scroll_job = None
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_wheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        w = widget
        while w is not None and w is not self:
            w = w.master
        if w is not self:
            return
        if getattr(event, "num", None) == 4:
            self.canvas.yview_scroll(-3, "units")
        elif getattr(event, "num", None) == 5:
            self.canvas.yview_scroll(3, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120) * 3), "units")


class ConciliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.style_engine = ttk.Style(self)
        self.style_engine.theme_use("clam")
        self.configure(bg=PALETTE["bg"])
        self.title(f"{APP_NAME} - v{APP_VERSION}")
        self.geometry("880x700")
        self.minsize(800, 520)
        self.set_app_icon()
        self.file_path = tk.StringVar()
        self.sheet_vars = {clave: tk.StringVar(value=nombre) for clave, nombre in DEFAULT_SHEET_NAMES.items()}
        self.hojas_disponibles = []
        self.column_vars = {}
        self._cola_eventos = queue.Queue()
        self.hoja_iva_var = tk.StringVar()
        self._procesando_iva = False
        self._procesando_token=False
        self._procesando = False
        self._generador_iva=None
        self._procesando_pdf_iva = False
        self._procesando_iva=False

        self._setup_styles()
        self._build_ui()
        self.after(150, self._procesar_cola)


    def set_app_icon(self):
        icon_path = Path(__file__).resolve().parent
        icon_path = icon_path / "assets" / "icon.ico"

        if icon_path.exists():
            try:
                self.iconbitmap(icon_path)
            except Exception as e:
                logger.error(f"Error al establecer el icono de la aplicación: {e}")
        else:
                logger.info("Icono de la aplicación establecido correctamente.")

    # ------------------------------------------------------------------ STYLES & UI SETUP

    def _on_liberar_archivo(self):
        if self._procesando or self._procesando_iva or self._procesando_token:
            messagebox.showwarning("Proceso en curso", "Espera a que finalice el proceso actual antes de liberar el archivo.")
            return

        self.file_path.set("")
        self.hojas_disponibles = []

        for clave, combo in self.combos_hojas.items():
            combo['values'] = []
            self.sheet_vars[clave].set("" if clave in OPTIONAL_SHEETS else DEFAULT_SHEET_NAMES[clave])

        for clave, combo in self.combos_token.items():
            combo['values'] = []

        if hasattr(self, "lbl_estado_token"):
            self.lbl_estado_token.configure(text="")

        self._poblar_checklist_columnas([])
        

        if hasattr(self, "combo_hoja_iva"):
            self.combo_hoja_iva['values'] = []
        if hasattr(self, "hoja_iva_var"):
            self.hoja_iva_var.set("")
        if hasattr(self, "lbl_estado_iva"):
            self.lbl_estado_iva.configure(text="")

        for widget in self.tab_preview.winfo_children():
            widget.destroy()
        self._build_tab_preview()

        import gc
        gc.collect()
        # gui.py — _on_liberar_archivo (agregar dentro del método)
        self._generador_iva = None
        if hasattr(self, "btn_exportar_pdf"):
            self.btn_exportar_pdf.configure(state="disabled")
        if hasattr(self, "btn_vista_previa_pdf"):
            self.btn_vista_previa_pdf.configure(state="disabled")

        self._set_estado("Sesión liberada. El archivo quedó disponible para otros programas.")
            
# gui.py — reemplazar _on_generar_informe_iva y _generar_informe_iva_en_hilo

    def _on_generar_informe_iva(self):
        if self._procesando_iva:
            return
        if not self.file_path.get():
            messagebox.showwarning("Archivo requerido", "Primero debes seleccionar un archivo de Excel.")
            return
        hoja = self.hoja_iva_var.get().strip()
        if not hoja:
            messagebox.showwarning("Hoja requerida", "Debes seleccionar la hoja de token principal.")
            return

        self._procesando_iva = True
        self._set_btn_enabled(self.btn_generar_iva, False)
        self.btn_exportar_pdf.configure(state="disabled")
        self.btn_vista_previa_pdf.configure(state="disabled")
        self.lbl_estado_iva.configure(text="Generando informe...")

        self._generador_iva = GeneradorInformeIVA(
            self.file_path.get(), hoja, progress_callback=self._on_progreso_iva
        )

        hilo = threading.Thread(
            target=self._generar_informe_iva_en_hilo,
            args=(self._generador_iva,),
            daemon=True
        )
        hilo.start()

    def _generar_informe_iva_en_hilo(self, generador):
        try:
            resumen = generador.generar()
            self._cola_eventos.put(("iva_exito", resumen))
        except ErrorUsuario as e:
            logger.warning(f"Error de usuario al generar informe de IVA: {e}")
            self._cola_eventos.put(("iva_error_usuario", str(e)))
        except ErrorSistema as e:
            logger.error(f"Error de sistema al generar informe de IVA: {e}")
            self._cola_eventos.put(("iva_error_sistema", str(e)))
        except Exception as e:
            logger.exception("Error inesperado al generar informe de IVA")
            self._cola_eventos.put(("iva_error_sistema", f"Error inesperado: {e}"))

    def _on_progreso_iva(self, mensaje):
        self._cola_eventos.put(("iva_log", mensaje))

    def _finalizar_iva(self):
        self._procesando_iva = False
        self._set_btn_enabled(self.btn_generar_iva, True)

    # gui.py — nuevos métodos (agregar cerca de _finalizar_iva)
    def _on_exportar_pdf_iva(self):
            if self._procesando_pdf_iva:
                return
            if not self._generador_iva or self._generador_iva.df_filtrado is None:
                messagebox.showwarning("Informe requerido", "Primero debes generar el informe de IVA.")
                return

            # Generar la ruta sugerida aplicando la lógica de Nombre Emisor + Informe de IVA + Periodo
            cliente_nombre = ""
            df_f = self._generador_iva.df_filtrado
            if "Grupo" in df_f.columns and "Nombre Emisor" in df_f.columns:
                df_emitidos = df_f[df_f["Grupo"].astype(str).str.strip() == "Emitido"]
                if not df_emitidos.empty:
                    val_emisor = df_emitidos["Nombre Emisor"].iloc[0]
                    if pd.notna(val_emisor) and str(val_emisor).strip():
                        cliente_nombre = str(val_emisor).strip() + " - "

            periodo_limpio = self._generador_iva.periodo_texto.replace("PERÍODO:", "").strip() if self._generador_iva.periodo_texto else ""
            sufijo_periodo = f" - {periodo_limpio}" if periodo_limpio else ""

            nombre_sugerido = f"{cliente_nombre}INFORME DE IVA{sufijo_periodo}.pdf"
            for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                nombre_sugerido = nombre_sugerido.replace(char, '')

            ruta_sugerida = Path(self._generador_iva.ruta_archivo).with_name(nombre_sugerido)

            ruta_pdf = filedialog.asksaveasfilename(
                title="Guardar informe de IVA en PDF",
                defaultextension=".pdf",
                initialfile=ruta_sugerida.name,
                filetypes=[("Archivo PDF", "*.pdf")],
            )
            if not ruta_pdf:
                return

            self._procesando_pdf_iva = True
            self.btn_exportar_pdf.configure(state="disabled")
            self.lbl_estado_iva.configure(text="Generando PDF...")

            hilo = threading.Thread(
                target=self._exportar_pdf_iva_en_hilo,
                args=(ruta_pdf,),
                daemon=True
            )
            hilo.start()

    def _exportar_pdf_iva_en_hilo(self, ruta_pdf):
        try:
            ruta_final = self._generador_iva.generar_pdf(ruta_pdf)
            self._cola_eventos.put(("pdf_iva_exito", ruta_final))
        except ErrorUsuario as e:
            self._cola_eventos.put(("pdf_iva_error_usuario", str(e)))
        except ErrorSistema as e:
            self._cola_eventos.put(("pdf_iva_error_sistema", str(e)))
        except Exception as e:
            logger.exception("Error inesperado al exportar PDF de IVA")
            self._cola_eventos.put(("pdf_iva_error_sistema", f"Error inesperado: {e}"))

    def _finalizar_pdf_iva(self):
        self._procesando_pdf_iva = False
        self.btn_exportar_pdf.configure(state="normal")

    def _on_vista_previa_pdf(self):
        if not self._generador_iva or self._generador_iva.df_filtrado is None:
            messagebox.showwarning("Informe requerido", "Primero debes generar el informe de IVA.")
            return

        try:
            import pymupdf as fitzz  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Dependencia faltante",
                "No se encontró PyMuPDF. Instálalo con: pip install pymupdf"
            )
            return

        ruta_temporal = Path(tempfile.gettempdir()) / "vista_previa_informe_iva.pdf"
        try:
            self._generador_iva.generar_pdf(str(ruta_temporal))
        except (ErrorUsuario, ErrorSistema) as e:
            messagebox.showerror("Error al generar vista previa", str(e))
            return

        VentanaVistaPreviaPDF(self, str(ruta_temporal))

    def _setup_styles(self):
        self.style = self.style_engine

        font_family = "Segoe UI" if "win32" in self.tk.call("tk", "windowingsystem") else "Helvetica"
        self._font = font_family

        self.style.configure(".", font=(font_family, 9), background=PALETTE["bg"])
        self.style.configure("TFrame", background=PALETTE["bg"])
        self.style.configure("Card.TFrame", background=PALETTE["surface"])
        self.style.configure("CardAlt.TFrame", background=PALETTE["surface_alt"])

        self.style.configure(
            "Header.TLabel",
            font=(font_family, 12, "bold"),
            foreground=PALETTE["text"],
            background=PALETTE["bg"],
        )
        self.style.configure(
            "CardHeader.TLabel",
            font=(font_family, 12, "bold"),
            foreground=PALETTE["text"],
            background=PALETTE["surface"],
        )
        self.style.configure(
            "Subheader.TLabel",
            font=(font_family, 9),
            foreground=PALETTE["text_muted"],
            background=PALETTE["bg"],
        )
        self.style.configure(
            "SubheaderCard.TLabel",
            font=(font_family, 9),
            foreground=PALETTE["text_muted"],
            background=PALETTE["surface"],
        )
        self.style.configure(
            "FieldLabel.TLabel",
            font=(font_family, 9, "bold"),
            foreground=PALETTE["text"],
            background=PALETTE["surface"],
        )
        self.style.configure(
            "FieldLabelOptional.TLabel",
            font=(font_family, 9),
            foreground=PALETTE["text_muted"],
            background=PALETTE["surface"],
        )

        self.style.configure(
            "TNotebook",
            background=PALETTE["bg"],
            borderwidth=0,
            tabmargins=(4, 6, 4, 0),
        )
        self.style.configure(
            "TNotebook.Tab",
            font=(font_family, 9, "bold"),
            padding=(16, 8),
        )

        self.style.configure(
            "Secondary.TButton",
            font=(font_family, 9),
            padding=(10, 6),
        )
        self.style.configure(
            "Toggle.TButton",
            font=(font_family, 8, "bold"),
            padding=(8, 4),
        )

        self.style.configure(
            "TCombobox",
            padding=(6, 5),
            fieldbackground=PALETTE["surface_alt"],
            background=PALETTE["surface_alt"],
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", PALETTE["surface_alt"])],
        )
        self.style.configure(
            "TEntry",
            padding=(8, 6),
            fieldbackground=PALETTE["surface_alt"],
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("readonly", PALETTE["surface_alt"])],
        )

        self.style.configure(
            "Status.TLabel",
            font=(font_family, 8),
            padding=(14, 7),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text_muted"],
        )
        self.style.configure(
            "StatusDot.TLabel",
            font=(font_family, 10),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text_muted"],
        )

        self.style.configure("TSeparator", background=PALETTE["border_strong"])
        self.style.configure(
            "Card.TLabelframe",
            background=PALETTE["surface"],
            borderwidth=1,
            relief="solid",
            bordercolor=PALETTE["border_strong"],
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            font=(font_family, 9, "bold"),
            foreground=PALETTE["text"],
            background=PALETTE["surface"],
        )
        self.style.configure("Card.TCheckbutton", background=PALETTE["surface"], font=(font_family, 9))

        self.style.configure(
            "TProgressbar",
            thickness=6,
            background=PALETTE["primary"],
            troughcolor=PALETTE["border"],
            borderwidth=0,
        )
        self.style.configure(
            "TNotebook",
            background=PALETTE["bg"],
            borderwidth=0,
            tabmargins=(4, 6, 4, 0),
        )
        self.style.configure(
            "TNotebook.Tab",
            font=(font_family, 9, "bold"),
            padding=(16, 8),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text_muted"],
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", PALETTE["surface"])],
            foreground=[("selected", PALETTE["text"])],
        )

        self.style.configure(
            "Secondary.TButton",
            font=(font_family, 9),
            padding=(10, 6),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text"],
            bordercolor=PALETTE["border_strong"],
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", PALETTE["border"])],
        )
        self.style.configure(
            "Toggle.TButton",
            font=(font_family, 8, "bold"),
            padding=(8, 4),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text"],
            bordercolor=PALETTE["border_strong"],
        )
        self.style.map(
            "Toggle.TButton",
            background=[("active", PALETTE["border"])],
        )

        self.style.configure(
            "TCombobox",
            padding=(6, 5),
            fieldbackground=PALETTE["surface_alt"],
            background=PALETTE["surface_alt"],
            arrowsize=14,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", PALETTE["surface_alt"])],
            background=[("readonly", PALETTE["surface_alt"])],
        )
        self.style.configure(
            "TEntry",
            padding=(8, 6),
            fieldbackground=PALETTE["surface_alt"],
            bordercolor=PALETTE["border_strong"],
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("readonly", PALETTE["surface_alt"])],
        )

        self.style.configure(
            "Status.TLabel",
            font=(font_family, 8),
            padding=(14, 7),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text_muted"],
        )
        self.style.configure(
            "StatusDot.TLabel",
            font=(font_family, 10),
            background=PALETTE["surface_alt"],
            foreground=PALETTE["text_muted"],
        )

        self.style.configure("TSeparator", background=PALETTE["border_strong"])
        self.style.configure(
            "Card.TLabelframe",
            background=PALETTE["surface"],
            borderwidth=1,
            relief="solid",
            bordercolor=PALETTE["border_strong"],
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            font=(font_family, 9, "bold"),
            foreground=PALETTE["text"],
            background=PALETTE["surface"],
        )
        self.style.configure("Card.TCheckbutton", background=PALETTE["surface"], font=(font_family, 9))
        self.style.map(
            "Card.TCheckbutton",
            background=[("active", PALETTE["surface"])],
        )

        self.style.configure(
            "TProgressbar",
            thickness=6,
            background=PALETTE["primary"],
            troughcolor=PALETTE["border"],
            borderwidth=0,
        )

    def _build_ui(self):
        self._build_barra_superior()

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=(14, 12))

        self.tab_config = ttk.Frame(self.notebook, style="TFrame")
        self.tab_preview = ttk.Frame(self.notebook, style="TFrame")
        self.tab_ejecucion = ttk.Frame(self.notebook, style="TFrame")
        self.tab_reportes_iva=ttk.Frame(self.notebook, style="TFrame")
        self.tab_token=ttk.Frame(self.notebook, style="TFrame")

        self.notebook.add(self.tab_token, text=" 🧾  Formatear Token  ")
        self.notebook.add(self.tab_config, text="  ⚙  Configuración para auditoria  ")
        self.notebook.add(self.tab_reportes_iva, text="  📊  Reportes IVA  ")
        self.notebook.add(self.tab_preview, text="  👁  Vista Previa  ")
        self.notebook.add(self.tab_ejecucion, text="  ▶  Ejecución y Logs  ")
        
        self._build_tab_config()
        self._build_tab_preview()
        self._build_tab_ejecucion()
        self._build_tab_reportes_iva()
        self._build_tab_token()

        self._build_barra_estado()

    def _build_barra_superior(self):
        frame_top = ttk.Frame(self, padding=(18, 16, 18, 14), style="TFrame")
        frame_top.pack(fill=tk.X)

        titulo = ttk.Frame(frame_top, style="TFrame")
        titulo.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(titulo, text=APP_NAME, style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Label(titulo, text=f"  ·  v{APP_VERSION}", style="Subheader.TLabel").pack(side=tk.LEFT)

        fila_archivo = ttk.Frame(frame_top, style="TFrame")
        fila_archivo.pack(fill=tk.X)

        lbl_instruccion = ttk.Label(
            fila_archivo,
            text="Archivo de trabajo",
            style="Subheader.TLabel"
        )
        lbl_instruccion.pack(side=tk.LEFT, padx=(0, 12))

        entry_path = ttk.Entry(
            fila_archivo,
            textvariable=self.file_path,
            state="readonly"
        )
        entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

        btn_browse = self._btn_primary(fila_archivo, "📂  Examinar...", self._on_abrir_archivo)
        btn_browse.pack(side=tk.LEFT)

        btn_liberar=ttk.Button(            
            fila_archivo,
            text="🔓 Liberar archivo",
            command=self._on_liberar_archivo,
            cursor="hand2"
        )
        btn_liberar.pack(side=tk.LEFT, padx=(12, 0))

    def _card(self, parent, **pack_kwargs):
        outer = tk.Frame(parent, bg=PALETTE["border_strong"])
        inner = ttk.Frame(outer, style="Card.TFrame", padding=20)
        inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        outer.pack(**pack_kwargs)
        return inner

    def _btn_primary(self, parent, text, command):
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=(self._font, 10, "bold"),
            bg=PALETTE["primary"],
            fg=PALETTE["primary_fg"],
            activebackground=PALETTE["primary_hover"],
            activeforeground=PALETTE["primary_fg"],
            disabledforeground="#E5E7EB",
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=PALETTE["primary_hover"]) if btn["state"] != "disabled" else None)
        btn.bind("<Leave>", lambda e: btn.configure(bg=PALETTE["primary"]) if btn["state"] != "disabled" else None)
        return btn

    def _set_btn_enabled(self, btn, enabled):
        btn.configure(
            state="normal" if enabled else "disabled",
            bg=PALETTE["primary"] if enabled else "#9CA3AF",
        )

    # gui.py — corregido


       

    def _build_tab_config(self):
        wrapper = ttk.Frame(self.tab_config, padding=(4, 16, 4, 4), style="TFrame")
        wrapper.pack(fill=tk.BOTH, expand=True)

        card_hojas = self._card(wrapper, fill=tk.X, pady=(0, 16))

        ttk.Label(
            card_hojas,
            text="Asignación de Estructura de Hojas",
            style="CardHeader.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            card_hojas,
            text="Relaciona cada sección requerida con el nombre exacto de la pestaña cargada en tu archivo.",
            style="SubheaderCard.TLabel"
        ).pack(anchor="w", pady=(4, 16))

        grid_frame = ttk.Frame(card_hojas, style="Card.TFrame")
        grid_frame.pack(fill=tk.X, anchor="n")

        self.combos_hojas = {}
        for i, clave in enumerate(SHEET_ORDER):
            etiqueta = SHEET_LABELS[clave]
            es_opcional = clave in OPTIONAL_SHEETS

            lbl_texto = f"{etiqueta}" if es_opcional else f"{etiqueta}  *"
            estilo_lbl = "FieldLabelOptional.TLabel" if es_opcional else "FieldLabel.TLabel"

            lbl = ttk.Label(grid_frame, text=lbl_texto, style=estilo_lbl)
            lbl.grid(row=i, column=0, sticky="w", pady=9, padx=(0, 24))

            combo = ttk.Combobox(
                grid_frame,
                textvariable=self.sheet_vars[clave],
                width=45,
                state="readonly",
                cursor="hand2"
            )
            combo.grid(row=i, column=1, sticky="ew", pady=9)
            self.combos_hojas[clave] = combo

            if clave == "aud_comp":
                combo.bind("<<ComboboxSelected>>", self._refrescar_columnas_aud_comp)

            # --- NUEVO: Botón para limpiar campos opcionales ---
            if es_opcional:
                btn_limpiar = tk.Button(
                    grid_frame,
                    text="✖",
                    command=lambda c=clave: self.sheet_vars[c].set(""), # Borra el contenido de la variable
                    bg=PALETTE["surface"],
                    fg=PALETTE["text_muted"],
                    font=(self._font, 10, "bold"),
                    relief="flat",
                    bd=0,
                    cursor="hand2"
                )
                # Efectos hover para que se ponga rojo al pasar el mouse
                btn_limpiar.bind("<Enter>", lambda e, b=btn_limpiar: b.configure(fg=PALETTE["danger"]))
                btn_limpiar.bind("<Leave>", lambda e, b=btn_limpiar: b.configure(fg=PALETTE["text_muted"]))
                
                btn_limpiar.grid(row=i, column=2, padx=(8, 0))

        grid_frame.columnconfigure(1, weight=1)

        ttk.Label(
            card_hojas,
            text="* Campo obligatorio",
            style="SubheaderCard.TLabel"
        ).pack(anchor="w", pady=(10, 0))

        card_seriales = self._card(wrapper, fill=tk.BOTH, expand=True)

        header_seriales = ttk.Frame(card_seriales, style="Card.TFrame")
        header_seriales.pack(fill=tk.X)

        texto_seriales = ttk.Frame(header_seriales, style="Card.TFrame")
        texto_seriales.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(
            texto_seriales,
            text="Configuración de Seriales",
            style="CardHeader.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            texto_seriales,
            text="Marca a qué categoría(s) pertenece cada columna de la hoja AUD-COMP.",
            style="SubheaderCard.TLabel"
        ).pack(anchor="w", pady=(4, 0))

        barra_toggle = ttk.Frame(card_seriales, style="Card.TFrame")
        barra_toggle.pack(fill=tk.X, pady=(14, 10))

        self._build_leyenda_categorias(barra_toggle)

        acciones_masivas = ttk.Frame(card_seriales, style="Card.TFrame")
        acciones_masivas.pack(fill=tk.X, pady=(0, 10))

        for etiqueta, categoria, color in (
            ("IVA", "iva", CATEGORY_COLORS["iva"]),
            ("BASE", "base", CATEGORY_COLORS["base"]),
            ("BASE 2", "base2", CATEGORY_COLORS["base2"]),
            ("Autorretenedor", "autorretenedor", CATEGORY_COLORS["autorretenedor"])
        ):
            grupo = tk.Frame(acciones_masivas, bg=PALETTE["surface"])
            grupo.pack(side=tk.LEFT, padx=(0, 20))
            tk.Frame(grupo, bg=color, width=10, height=10).pack(side=tk.LEFT, padx=(0, 6), pady=2)
            ttk.Button(
                grupo, text=f"Todo {etiqueta}", style="Toggle.TButton", cursor="hand2",
                command=lambda c=categoria: self._marcar_todos(c, True)
            ).pack(side=tk.LEFT)
            ttk.Button(
                grupo, text="Ninguno", style="Toggle.TButton", cursor="hand2",
                command=lambda c=categoria: self._marcar_todos(c, False)
            ).pack(side=tk.LEFT, padx=(4, 0))

        self.lista_columnas_aud = ScrollableChecklist(card_seriales, height=240)
        self.lista_columnas_aud.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self._poblar_checklist_columnas([])

    def _build_leyenda_categorias(self, parent):
        for etiqueta, color in (
            ("IVA", CATEGORY_COLORS["iva"]),
            ("BASE", CATEGORY_COLORS["base"]),
            ("BASE 2", CATEGORY_COLORS["base2"]),
            ("Autorretenedor", CATEGORY_COLORS["autorretenedor"])
        ):
            chip = tk.Frame(parent, bg=color)
            chip.pack(side=tk.LEFT, padx=(0, 16))
            tk.Label(
                chip, text=f"  {etiqueta}  ", bg=color, fg="#FFFFFF",
                font=(self._font, 8, "bold")
            ).pack(ipady=2)

    def _refrescar_columnas_aud_comp(self, *_):
        ruta = self.file_path.get()
        hoja = self.sheet_vars["aud_comp"].get().strip()
        if not ruta or not hoja or hoja not in self.hojas_disponibles:
            self._poblar_checklist_columnas([])
            return

        try:
            df = ConciliadorAuditoria._cargar_hoja_con_encabezado_variable(ruta, hoja)
        except Exception as e:
            logger.warning(f"No se pudieron leer las columnas de AUD-COMP: {e}")
            self._poblar_checklist_columnas([])
            return

        columnas = []
        for col in df.columns:
            nombre = str(col).strip()
            if not nombre or nombre.startswith("Unnamed"):
                continue
            if nombre in COLUMNAS_EXCLUIDAS_AUD_COMP:
                continue
            if df[col].dropna().empty:
                continue
            columnas.append(nombre)

        self._poblar_checklist_columnas(columnas)

    def _poblar_checklist_columnas(self, columnas):
        self.lista_columnas_aud.canvas.unbind("<Configure>")
        for widget in self.lista_columnas_aud.inner.winfo_children():
            widget.destroy()

        self.column_vars = {}

        if not columnas:
            ttk.Label(
                self.lista_columnas_aud.inner,
                text="Selecciona un archivo y la hoja AUD-COMP para ver las columnas disponibles.",
                style="SubheaderCard.TLabel"
            ).pack(anchor="w", padx=4, pady=8)
        else:
            for idx, col in enumerate(columnas):
                fondo_fila = PALETTE["surface"] if idx % 2 == 0 else PALETTE["surface_alt"]
                fila = tk.Frame(self.lista_columnas_aud.inner, bg=fondo_fila)
                fila.pack(fill=tk.X, pady=1)

                contenido = tk.Frame(fila, bg=fondo_fila, padx=10, pady=7)
                contenido.pack(fill=tk.X)

                # SOLUCIÓN VISUAL: Quitamos expand=True y fijamos un width para 
                # mantener alineadas las casillas cerca del texto.
                tk.Label(
                    contenido, text=col, bg=fondo_fila, fg=PALETTE["text"],
                    font=(self._font, 9), anchor="w", width=50
                ).pack(side=tk.LEFT, padx=(0, 20))

                vars_col = {}
                for clave_cat, color in CATEGORY_COLORS.items():
                    var = tk.BooleanVar(value=False)
                    chip = self._crear_chip_toggle(contenido, color, var)
                    chip.pack(side=tk.LEFT, padx=(6, 0))
                    vars_col[clave_cat] = var

                # SOLUCIÓN LÓGICA: Exclusión mutua entre IVA y BASE
                var_iva = vars_col.get("iva")
                var_base = vars_col.get("base")
                var_retenedor = vars_col.get("autorretenedor")

                if var_iva and var_base and var_retenedor:
                    # Usamos una función constructora para evitar problemas de alcance (scope) en el ciclo for
                    def hacer_exclusivo(v_activa, v_otra, v_retenedor=None):
                        def _trace(*args):
                            if v_activa.get():  # Si esta variable se enciende
                                v_otra.set(False) # Apagamos la otra
                                v_retenedor.set(False) if v_retenedor else None # Apagamos la otra si existe
                        return _trace

                    # Enlazamos los eventos a las variables
                    var_iva.trace_add("write", hacer_exclusivo(var_iva, var_base, var_retenedor))
                    var_base.trace_add("write", hacer_exclusivo(var_base, var_iva, var_retenedor))
                    var_retenedor.trace_add("write", hacer_exclusivo(var_retenedor, var_iva, var_base))

                self.column_vars[col] = vars_col

        self.lista_columnas_aud.canvas.bind(
            "<Configure>",
            lambda e: self.lista_columnas_aud.canvas.itemconfigure(self.lista_columnas_aud._window, width=e.width)
        )
        self.lista_columnas_aud._actualizar_scrollregion()

    def _crear_chip_toggle(self, parent, color, var):
        chip = tk.Label(
            parent, text=" ", width=3, bg=PALETTE["border"], relief="flat",
            cursor="hand2", font=(self._font, 8, "bold")
        )

        def _refrescar(*_):
            chip.configure(bg=color if var.get() else PALETTE["border"])

        def _toggle(_event=None):
            var.set(not var.get())
            _refrescar()

        chip.bind("<Button-1>", _toggle)
        var.trace_add("write", _refrescar)
        _refrescar()
        return chip

    def _marcar_todos(self, tipo, valor):
        for vars_col in self.column_vars.values():
            vars_col[tipo].set(valor)

    def _build_tab_preview(self):
        self.vista_previa = VistaPreviaExcel(self.tab_preview)
        self.vista_previa.pack(fill=tk.BOTH, expand=True, padx=4, pady=(16, 4))

    def _build_tab_token(self):
        wrapper = ttk.Frame(self.tab_token, padding=(4, 16, 4, 4), style="TFrame")
        wrapper.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            wrapper,
            text="Formateo de Token",
            style="CardHeader.TLabel"
        ).pack(anchor="w", pady=(0, 16))

        card = self._card(wrapper, fill=tk.X, pady=(0, 16))

        ttk.Label(card, text="Hojas de origen", style="CardHeader.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Selecciona la hoja de Token y, si aplica, Contabilidad y Terceros para asignar Concepto y Tercero.",
            style="SubheaderCard.TLabel"
        ).pack(anchor="w", pady=(4, 16))

        grid_frame = ttk.Frame(card, style="Card.TFrame")
        grid_frame.pack(fill=tk.X, anchor="n")

        self.combos_token = {}
        for i, clave in enumerate(("principal", "contabilidad", "terceros")):
            etiqueta = SHEET_LABELS[clave]
            es_opcional = clave in OPTIONAL_SHEETS
            lbl_texto = f"{etiqueta}" if es_opcional else f"{etiqueta}  *"
            estilo_lbl = "FieldLabelOptional.TLabel" if es_opcional else "FieldLabel.TLabel"

            ttk.Label(grid_frame, text=lbl_texto, style=estilo_lbl).grid(row=i, column=0, sticky="w", pady=9, padx=(0, 24))

            combo = ttk.Combobox(
                grid_frame,
                textvariable=self.sheet_vars[clave],
                width=45,
                state="readonly",
                cursor="hand2"
            )
            combo.grid(row=i, column=1, sticky="ew", pady=9)
            self.combos_token[clave] = combo

        grid_frame.columnconfigure(1, weight=1)

        fila_accion = ttk.Frame(card, style="Card.TFrame")
        fila_accion.pack(fill=tk.X, pady=(20, 0))

        self.btn_formatear_token = self._btn_primary(fila_accion, "🧾  Formatear Token", self._on_formatear_token)
        self.btn_formatear_token.pack(side=tk.LEFT)

        self.lbl_estado_token = ttk.Label(fila_accion, text="", style="SubheaderCard.TLabel")
        self.lbl_estado_token.pack(side=tk.LEFT, padx=(16, 0))

    def _on_formatear_token(self):
        if self._procesando_token:
            return
        if not self.file_path.get():
            messagebox.showwarning("Archivo requerido", "Primero debes seleccionar un archivo de Excel.")
            return

        nombre_principal = self.sheet_vars["principal"].get().strip()
        if not nombre_principal:
            messagebox.showwarning("Hoja requerida", "Debes indicar la hoja de Token.")
            return
        if nombre_principal not in self.hojas_disponibles:
            messagebox.showwarning("Hoja inválida", f"La hoja '{nombre_principal}' no existe en el archivo.")
            return

        respuesta = messagebox.askyesno(
            "Confirmar formateo",
            "Este proceso modificará la hoja de Token en el archivo seleccionado. ¿Deseas continuar?"
        )
        if not respuesta:
            return

        self._procesando_token = True
        self._set_btn_enabled(self.btn_formatear_token, False)
        self.lbl_estado_token.configure(text="Formateando...")

        sheet_names = {
            "principal": nombre_principal,
            "contabilidad": self.sheet_vars["contabilidad"].get().strip(),
            "terceros": self.sheet_vars["terceros"].get().strip(),
        }

        hilo = threading.Thread(
            target=self._formatear_token_en_hilo,
            args=(self.file_path.get(), sheet_names),
            daemon=True
        )
        hilo.start()

    def _formatear_token_en_hilo(self, ruta, sheet_names):
        formateador = FormateadorToken(ruta, sheet_names, progress_callback=self._on_progreso_token)
        try:
            resumen = formateador.ejecutar()
            self._cola_eventos.put(("token_exito", resumen))
        except ErrorUsuario as e:
            logger.warning(f"Error de usuario al formatear token: {e}")
            self._cola_eventos.put(("token_error_usuario", str(e)))
        except ErrorSistema as e:
            logger.error(f"Error de sistema al formatear token: {e}")
            self._cola_eventos.put(("token_error_sistema", str(e)))
        except Exception as e:
            logger.exception("Error inesperado al formatear token")
            self._cola_eventos.put(("token_error_sistema", f"Error inesperado: {e}"))

    def _on_progreso_token(self, mensaje):
        self._cola_eventos.put(("token_log", mensaje))

    def _finalizar_token(self):
        self._procesando_token = False
        self._set_btn_enabled(self.btn_formatear_token, True)


    def _build_tab_reportes_iva(self):
        wrapper = ttk.Frame(self.tab_reportes_iva, padding=(4, 16, 4, 4), style="TFrame")
        wrapper.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            wrapper,
            text="Generación de Reportes de IVA",
            style="CardHeader.TLabel"
        ).pack(anchor="w", pady=(0, 16))

        card = self._card(wrapper, fill=tk.X, pady=(0, 16))

        ttk.Label(card, text="Configuración del Informe", style="CardHeader.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Selecciona la hoja de token principal sobre la cual se generará el informe de IVA.",
            style="SubheaderCard.TLabel"
        ).pack(anchor="w", pady=(4, 16))

        fila_hoja = ttk.Frame(card, style="Card.TFrame")
        fila_hoja.pack(fill=tk.X)

        ttk.Label(fila_hoja, text="Hoja de token", style="FieldLabel.TLabel").pack(side=tk.LEFT, padx=(0, 24))

        self.combo_hoja_iva = ttk.Combobox(
            fila_hoja,
            textvariable=self.hoja_iva_var,
            width=45,
            state="readonly",
            cursor="hand2"
        )
        self.combo_hoja_iva.pack(side=tk.LEFT, fill=tk.X, expand=True)

        fila_accion = ttk.Frame(card, style="Card.TFrame")
        fila_accion.pack(fill=tk.X, pady=(20, 0))

        self.btn_generar_iva = self._btn_primary(fila_accion, "📊  Generar Informe", self._on_generar_informe_iva)
        self.btn_generar_iva.pack(side=tk.LEFT)

        self.lbl_estado_iva = ttk.Label(fila_accion, text="", style="SubheaderCard.TLabel")
        self.lbl_estado_iva.pack(side=tk.LEFT, padx=(16, 0))

        fila_pdf = ttk.Frame(card, style="Card.TFrame")
        fila_pdf.pack(fill=tk.X, pady=(12, 0))

        self.btn_exportar_pdf = ttk.Button(
            fila_pdf, text="📄  Exportar a PDF", style="Secondary.TButton",
            command=self._on_exportar_pdf_iva, cursor="hand2", state="disabled"
        )
        self.btn_exportar_pdf.pack(side=tk.LEFT)

        self.btn_vista_previa_pdf = ttk.Button(
            fila_pdf, text="👁  Vista previa PDF", style="Secondary.TButton",
            command=self._on_vista_previa_pdf, cursor="hand2", state="disabled"
        )
        self.btn_vista_previa_pdf.pack(side=tk.LEFT, padx=(10, 0))

    def _build_tab_ejecucion(self):
        contenedor = ttk.Frame(self.tab_ejecucion, padding=(4, 16, 4, 4), style="TFrame")
        contenedor.pack(fill=tk.BOTH, expand=True)

        panel_acciones = ttk.Frame(contenedor, style="TFrame")
        panel_acciones.pack(fill=tk.X, pady=(0, 16))

        self.btn_ejecutar = self._btn_primary(panel_acciones, "▶  Iniciar Conciliación", self._on_ejecutar)
        self.btn_ejecutar.pack(side=tk.LEFT)

        btn_logs = ttk.Button(
            panel_acciones,
            text="📁  Abrir Carpeta de Logs",
            style="Secondary.TButton",
            command=self._on_abrir_logs,
            cursor="hand2"
        )
        btn_logs.pack(side=tk.LEFT, padx=(12, 0))

        self.progress = ttk.Progressbar(contenedor, mode="indeterminate", style="TProgressbar")
        self.progress.pack(fill=tk.X, pady=(0, 16))

        lbl_console = ttk.Label(contenedor, text="Registro de Actividad", style="Header.TLabel")
        lbl_console.pack(anchor="w", pady=(0, 8))

        frame_consola = tk.Frame(contenedor, bg=PALETTE["console_bg"], bd=0, highlightthickness=1,
                                  highlightbackground=PALETTE["border"])
        frame_consola.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(
            frame_consola,
            state="disabled",
            height=18,
            wrap="word",
            bg=PALETTE["console_bg"],
            fg=PALETTE["console_fg"],
            insertbackground="#FFFFFF",
            selectbackground="#2A3A5C",
            font=("Consolas", 10),
            relief="flat",
            borderwidth=0,
            padx=14,
            pady=14
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

    def _build_barra_estado(self):
        contenedor = tk.Frame(self, bg=PALETTE["surface_alt"])
        contenedor.pack(fill=tk.X, side=tk.BOTTOM)

        self.lbl_estado_dot = ttk.Label(contenedor, text="●", style="StatusDot.TLabel", foreground=PALETTE["success"])
        self.lbl_estado_dot.pack(side=tk.LEFT, padx=(14, 4), pady=6)

        self.lbl_estado = ttk.Label(
            contenedor,
            text="Listo",
            anchor="w",
            style="Status.TLabel"
        )
        self.lbl_estado.pack(fill=tk.X, side=tk.LEFT)

    # ------------------------------------------------------------ ACCIONES
    def _on_abrir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo de Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xlsm")]
        )
        if not ruta:
            return

        try:
            xls = pd.ExcelFile(ruta)
        except Exception as e:
            logger.error(f"No se pudo abrir el archivo seleccionado: {e}")
            messagebox.showerror("Error al abrir archivo", f"No se pudo abrir el archivo:\n{e}")
            return

        self.file_path.set(ruta)
        self.hojas_disponibles = xls.sheet_names
        xls.close()

        for clave, combo in self.combos_hojas.items():
            if clave in OPTIONAL_SHEETS:
                combo['values'] = [""] + self.hojas_disponibles
            else:
                combo['values'] = self.hojas_disponibles

            actual = self.sheet_vars[clave].get()

            if actual not in self.hojas_disponibles and actual != "":
                if clave in OPTIONAL_SHEETS:
                    self.sheet_vars[clave].set("")
                elif self.hojas_disponibles:
                    self.sheet_vars[clave].set(self.hojas_disponibles[0])

        for clave, combo in self.combos_token.items():
            if clave in OPTIONAL_SHEETS:
                combo['values'] = [""] + self.hojas_disponibles
            else:
                combo['values'] = self.hojas_disponibles

        self.vista_previa.cargar_archivo(ruta, self.hojas_disponibles)

        self.combo_hoja_iva['values'] = self.hojas_disponibles
        if self.hoja_iva_var.get() not in self.hojas_disponibles:
            self.hoja_iva_var.set(self.hojas_disponibles[0] if self.hojas_disponibles else "")

        self._set_estado(f"Archivo cargado: {Path(ruta).name} ({len(self.hojas_disponibles)} hojas)")

    def _on_abrir_logs(self):
        import os
        import subprocess
        import sys

        carpeta = LOG_FILE.parent
        carpeta.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(carpeta)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(carpeta)])
            else:
                subprocess.Popen(["xdg-open", str(carpeta)])
        except Exception as e:
            logger.warning(f"No se pudo abrir la carpeta de logs: {e}")
            messagebox.showinfo("Carpeta de logs", f"Los logs se encuentran en:\n{carpeta}")

    def _validar_entradas(self):
        if not self.file_path.get():
            raise ErrorUsuario("Primero debes seleccionar un archivo de Excel.")

        for clave in SHEET_ORDER:
            if clave in OPTIONAL_SHEETS:
                continue
            valor = self.sheet_vars[clave].get().strip()
            if not valor:
                raise ErrorUsuario(f"Debes indicar el nombre de la hoja: {SHEET_LABELS[clave]}")
            if valor not in self.hojas_disponibles:
                raise ErrorUsuario(
                    f"La hoja '{valor}' indicada para '{SHEET_LABELS[clave]}' no existe en el archivo."
                )

    def _on_ejecutar(self):
        if self._procesando:
            return
        try:
            self._validar_entradas()
        except ErrorUsuario as e:
            logger.warning(f"Error de usuario al validar entradas: {e}")
            messagebox.showwarning("Datos incompletos", str(e))
            return

        seriales_iva_seleccionados = [col for col, v in self.column_vars.items() if v["iva"].get()]
        seriales_base_seleccionados = [col for col, v in self.column_vars.items() if v["base"].get()]
        seriales_base2_seleccionados = [col for col, v in self.column_vars.items() if v["base2"].get()]

        if not seriales_iva_seleccionados or not seriales_base_seleccionados:
            respuesta_vacia = messagebox.askyesno(
                "Seriales incompletos",
                "Has dejado una de las categorías de columnas vacía (sin marcar). ¿Deseas continuar?"
            )
            if not respuesta_vacia:
                return

        respuesta = messagebox.askyesno(
            "Confirmar ejecución",
            "Este proceso modificará el archivo de Excel seleccionado (se reescribirán "
            "hojas de resultados). ¿Deseas continuar?"
        )
        if not respuesta:
            return

        self._procesando = True
        self._set_btn_enabled(self.btn_ejecutar, False)
        self.lbl_estado_dot.configure(foreground=PALETTE["primary"])
        self.progress.start(12)
        self._limpiar_log()
        self.notebook.select(self.tab_ejecucion)

        sheet_names = {clave: var.get().strip() for clave, var in self.sheet_vars.items()}

        hilo = threading.Thread(
            target=self._ejecutar_en_hilo,
            args=(
                self.file_path.get(), sheet_names,
                seriales_iva_seleccionados, seriales_base_seleccionados, seriales_base2_seleccionados
            ),
            daemon=True
        )
        hilo.start()

    def _ejecutar_en_hilo(self, ruta, sheet_names, seriales_iva, seriales_base, seriales_base2):
        conciliador = ConciliadorAuditoria(
            ruta,
            sheet_names,
            seriales_iva=seriales_iva,
            seriales_base=seriales_base,
            seriales_base2=seriales_base2,
            progress_callback=self._on_progreso
        )
        try:
            resumen = conciliador.ejecutar()
            self._cola_eventos.put(("exito", resumen))
        except ErrorUsuario as e:
            logger.warning(f"Error de usuario durante la ejecución: {e}")
            self._cola_eventos.put(("error_usuario", str(e)))
        except ErrorSistema as e:
            logger.error(f"Error de sistema durante la ejecución: {e}")
            self._cola_eventos.put(("error_sistema", str(e)))
        except Exception as e:
            logger.exception("Error inesperado durante la ejecución")
            self._cola_eventos.put(("error_sistema", f"Error inesperado: {e}"))
    def _on_progreso(self, mensaje):
        self._cola_eventos.put(("log", mensaje))

    # ------------------------------------------------------------- COLA UI
    def _procesar_cola(self):
        try:
            while True:
                # Agrega este bloque try-except interno
                try:
                    tipo, payload = self._cola_eventos.get_nowait()
                except queue.Empty:
                    break  # Salimos del bucle while porque la cola ya está vacía

                # A partir de aquí mantienes tu lógica original intacta
                if tipo == "log":
                    self._agregar_log(payload)
                elif tipo == "exito":
                    self._finalizar_ejecucion(exito=True)
                    self._agregar_log(
                        f"Filas procesadas: {payload['filas_procesadas']} | "
                        f"Personales excluidas: {payload['filas_personales']} | "
                        f"Sin pareja: {payload['filas_sin_pareja']}"
                    )
                    messagebox.showinfo("Proceso completado", "La conciliación finalizó con éxito.")
                elif tipo == "error_usuario":
                    self._finalizar_ejecucion(exito=False)
                    self._agregar_log(f"ERROR DE USUARIO: {payload}")
                    messagebox.showwarning("Error en los datos", payload)
                elif tipo == "error_sistema":
                    self._finalizar_ejecucion(exito=False)
                    self._agregar_log(f"ERROR DEL SISTEMA: {payload}")
                    messagebox.showerror(
                        "Error del sistema",
                        f"{payload}\n\nRevisa el log para más detalles:\n{LOG_FILE}"
                    )
                elif tipo == "token_log":
                    self.lbl_estado_token.configure(text=payload)
                elif tipo == "token_exito":
                    self._finalizar_token()
                    self.lbl_estado_token.configure(
                        text=f"Filas: {payload['filas_procesadas']} | Personales: {payload['filas_personales']}"
                    )
                    if self.file_path.get():
                        self.vista_previa.cargar_archivo(self.file_path.get(), self.hojas_disponibles)
                    messagebox.showinfo("Token formateado", "La hoja de Token se formateó correctamente.")
                elif tipo == "token_error_usuario":
                    self._finalizar_token()
                    self.lbl_estado_token.configure(text=f"Error: {payload}")
                    messagebox.showwarning("Datos incompletos", payload)
                elif tipo == "token_error_sistema":
                    self._finalizar_token()
                    self.lbl_estado_token.configure(text=f"Error: {payload}")
                    messagebox.showerror("Error del sistema", payload)
                elif tipo == "iva_log":
                    self.lbl_estado_iva.configure(text=payload)
                elif tipo == "iva_exito":
                    self._finalizar_iva()
                    t = payload["totales"]
                    self.lbl_estado_iva.configure(text=f"IVA a pagar: {t['iva_a_pagar']:,.2f}")
                    self.btn_exportar_pdf.configure(state="normal")
                    self.btn_vista_previa_pdf.configure(state="normal")
                    messagebox.showinfo(
                        "Informe generado",
                        "El informe de IVA se generó correctamente en la hoja 'Informe IVA'."
                    )

                elif tipo == "iva_error_usuario":
                    self._finalizar_iva()
                    self.lbl_estado_iva.configure(text=f"Error: {payload}")
                    messagebox.showwarning("Datos incompletos", payload)
                elif tipo == "iva_error_sistema":
                    self._finalizar_iva()
                    self.lbl_estado_iva.configure(text=f"Error: {payload}")
                    messagebox.showerror("Error del sistema", payload)
                elif tipo == "pdf_iva_exito":
                    self._finalizar_pdf_iva()
                    self.lbl_estado_iva.configure(text=f"PDF exportado: {Path(payload).name}")
                    messagebox.showinfo("PDF generado", f"El informe se exportó correctamente en:\n{payload}")
                elif tipo == "pdf_iva_error_usuario":
                    self._finalizar_pdf_iva()
                    self.lbl_estado_iva.configure(text=f"Error: {payload}")
                    messagebox.showwarning("Datos incompletos", payload)
                elif tipo == "pdf_iva_error_sistema":
                    self._finalizar_pdf_iva()
                    self.lbl_estado_iva.configure(text=f"Error: {payload}")
                    messagebox.showerror("Error del sistema", payload)
            pass
        finally:
            self.after(300, self._procesar_cola)
    def _finalizar_ejecucion(self, exito=True):
        self._procesando = False
        self._set_btn_enabled(self.btn_ejecutar, True)
        self.progress.stop()
        color = PALETTE["success"] if exito else PALETTE["danger"]
        self.lbl_estado_dot.configure(foreground=color)
        self._set_estado("Listo")

    def _limpiar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

    def _agregar_log(self, mensaje):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, f"{mensaje}\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")
        self._set_estado(mensaje)

    def _set_estado(self, mensaje):
        self.lbl_estado.configure(text=f" {mensaje}")

# gui.py — nueva clase (agregar antes de class ConciliadorApp)

class VentanaVistaPreviaPDF(tk.Toplevel):
    def __init__(self, parent, ruta_pdf):
        super().__init__(parent)
        self.title("Vista previa - Informe de IVA")
        self.geometry("980x760")
        self.configure(bg=PALETTE["bg"])
        self.ruta_pdf = ruta_pdf
        self.pagina_actual = 0
        self._imagen_tk = None
        self.set_app_icon()

        import pymupdf as fitz
        self._fitz = fitz
        self.documento = fitz.open(ruta_pdf)

        barra = ttk.Frame(self, style="TFrame", padding=10)
        barra.pack(fill=tk.X)

        self.btn_anterior = ttk.Button(
            barra, text="◀ Anterior", style="Secondary.TButton",
            command=self._pagina_anterior, cursor="hand2"
        )
        self.btn_anterior.pack(side=tk.LEFT)

        self.lbl_pagina = ttk.Label(barra, text="", style="Subheader.TLabel")
        self.lbl_pagina.pack(side=tk.LEFT, padx=12)

        self.btn_siguiente = ttk.Button(
            barra, text="Siguiente ▶", style="Secondary.TButton",
            command=self._pagina_siguiente, cursor="hand2"
        )
        self.btn_siguiente.pack(side=tk.LEFT)

        btn_abrir = ttk.Button(
            barra, text="🗖  Abrir con visor predeterminado", style="Secondary.TButton",
            command=self._abrir_externo, cursor="hand2"
        )
        btn_abrir.pack(side=tk.RIGHT)

        contenedor = tk.Frame(self, bg=PALETTE["surface_alt"])
        contenedor.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(contenedor, bg=PALETTE["surface_alt"], highlightthickness=0)
        scrollbar_y = ttk.Scrollbar(contenedor, orient="vertical", command=self.canvas.yview)
        scrollbar_x = ttk.Scrollbar(contenedor, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.bind_all("<MouseWheel>", self._on_wheel, add="+")

        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self._mostrar_pagina()


    def set_app_icon(self):
        icon_path = Path(__file__).resolve().parent
        icon_path = icon_path / "assets" / "icon.ico"

        if icon_path.exists():
            try:
                self.iconbitmap(icon_path)
            except Exception as e:
                logger.error(f"Error al establecer el icono de la aplicación: {e}")
        else:
                logger.info("Icono de la aplicación establecido correctamente.")

    def _mostrar_pagina(self):
        pagina = self.documento[self.pagina_actual]
        matriz = self._fitz.Matrix(1.6, 1.6)
        pixmap = pagina.get_pixmap(matrix=matriz)
        modo = "RGBA" if pixmap.alpha else "RGB"
        imagen = Image.frombytes(modo, [pixmap.width, pixmap.height], pixmap.samples)
        self._imagen_tk = ImageTk.PhotoImage(imagen)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self._imagen_tk)
        self.canvas.configure(scrollregion=(0, 0, pixmap.width, pixmap.height))

        self.lbl_pagina.configure(text=f"Página {self.pagina_actual + 1} de {len(self.documento)}")
        self.btn_anterior.configure(state="normal" if self.pagina_actual > 0 else "disabled")
        self.btn_siguiente.configure(
            state="normal" if self.pagina_actual < len(self.documento) - 1 else "disabled"
        )

    def _pagina_anterior(self):
        if self.pagina_actual > 0:
            self.pagina_actual -= 1
            self._mostrar_pagina()

    def _pagina_siguiente(self):
        if self.pagina_actual < len(self.documento) - 1:
            self.pagina_actual += 1
            self._mostrar_pagina()

    def _on_wheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        w = widget
        while w is not None and w is not self:
            w = w.master
        if w is not self:
            return
        self.canvas.yview_scroll(int(-1 * (event.delta / 120) * 3), "units")

    def _abrir_externo(self):
        try:
            if os.name == "nt":
                os.startfile(self.ruta_pdf)
            else:
                import subprocess
                import sys
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", self.ruta_pdf])
        except Exception:
            messagebox.showinfo("Abrir PDF", f"El archivo se encuentra en:\n{self.ruta_pdf}")

    def _cerrar(self):
        self.documento.close()
        self.canvas.unbind_all("<MouseWheel>")
        self.destroy()