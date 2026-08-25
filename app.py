import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd


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
from excel_engine import ConciliadorExcel
from preview_widget import VistaPreviaExcel


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
        self._procesando = False

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

        self.notebook.add(self.tab_config, text="  ⚙  Configuración  ")
        self.notebook.add(self.tab_preview, text="  👁  Vista Previa  ")
        self.notebook.add(self.tab_ejecucion, text="  ▶  Ejecución y Logs  ")

        self._build_tab_config()
        self._build_tab_preview()
        self._build_tab_ejecucion()

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
            df = ConciliadorExcel._cargar_hoja_con_encabezado_variable(ruta, hoja)
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

                if var_iva and var_base:
                    # Usamos una función constructora para evitar problemas de alcance (scope) en el ciclo for
                    def hacer_exclusivo(v_activa, v_otra):
                        def _trace(*args):
                            if v_activa.get():  # Si esta variable se enciende
                                v_otra.set(False) # Apagamos la otra
                        return _trace

                    # Enlazamos los eventos a las variables
                    var_iva.trace_add("write", hacer_exclusivo(var_iva, var_base))
                    var_base.trace_add("write", hacer_exclusivo(var_base, var_iva))

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

        for clave, combo in self.combos_hojas.items():
            # Si es opcional, le agregamos una opción vacía ("") al principio de la lista
            if clave in OPTIONAL_SHEETS:
                combo['values'] = [""] + self.hojas_disponibles
            else:
                combo['values'] = self.hojas_disponibles
                
            actual = self.sheet_vars[clave].get()
            
            # Si el valor actual no está en el archivo y no es un campo vacío intencional
            if actual not in self.hojas_disponibles and actual != "":
                if clave in OPTIONAL_SHEETS:
                    self.sheet_vars[clave].set("")
                elif self.hojas_disponibles:
                    self.sheet_vars[clave].set(self.hojas_disponibles[0])

        self.vista_previa.cargar_archivo(ruta, self.hojas_disponibles)
        self._refrescar_columnas_aud_comp()
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
        conciliador = ConciliadorExcel(
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
                tipo, payload = self._cola_eventos.get_nowait()
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
        except queue.Empty:
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