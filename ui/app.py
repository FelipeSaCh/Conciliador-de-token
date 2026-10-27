import gc
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.config import APP_NAME, APP_VERSION, DEFAULT_SHEET_NAMES
from core.errors import logger
from shared.excel_io import hojas_de
from ui.pages.auditoria_page import AuditoriaPage
from ui.pages.autorretenedores_page import AutorretenedoresPage
from ui.pages.ejecucion_page import EjecucionPage
from ui.pages.iva_page import IvaPage
from ui.pages.movimientos_page import MovimientosPage
from ui.pages.preview_page import PreviewPage
from ui.pages.token_page import TokenPage
from ui.theme import PALETTE, aplicar_icono, configurar_estilos, fuente, fuente_titulo
from ui.widgets import boton_secundario
from ui.worker import EjecutorSegundoPlano

# (sección, [(id, texto del menú, clase)]) en el orden del menú lateral
NAVEGACION = [
    ("Proceso", [
        ("token", "1   Formatear token", TokenPage),
        ("pivote", "2   Movimientos", MovimientosPage),
        ("config", "3   Configurar auditoría", AuditoriaPage),
        ("exec", "4   Conciliar", EjecucionPage),
    ]),
    ("Reportes", [
        ("iva", "Informe de IVA", IvaPage),
    ]),
    ("Consulta", [
        ("preview", "Vista previa", PreviewPage),
        ("autorretenedores", "Autorretenedores", AutorretenedoresPage),
    ]),
]
PAGINA_INICIAL = "config"


class ConciliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        configurar_estilos(self)
        self.configure(bg=PALETTE["bg"])
        self.title(f"{APP_NAME} - v{APP_VERSION}")
        self.state("zoomed")
        self.minsize(980, 600)
        aplicar_icono(self)

        self.ejecutor = EjecutorSegundoPlano(self)
        self.file_path = tk.StringVar()
        self.sheet_vars = {clave: tk.StringVar(value=nombre) for clave, nombre in DEFAULT_SHEET_NAMES.items()}
        self.hojas_disponibles = []
        self.paginas = {}
        self.nav_buttons = {}
        self.current_page = None

        self._construir()

    # ---------------------------------------------------------------- layout
    def _construir(self):
        main = tk.Frame(self, bg=PALETTE["bg"])
        main.pack(fill=tk.BOTH, expand=True)
        self._construir_sidebar(main)

        self.content_frame = tk.Frame(main, bg=PALETTE["bg"])
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._construir_barra_archivo()

        contenedor = tk.Frame(self.content_frame, bg=PALETTE["bg"])
        contenedor.pack(fill=tk.BOTH, expand=True, padx=28, pady=(24, 0))
        for _, items in NAVEGACION:
            for page_id, _, clase in items:
                self.paginas[page_id] = clase(contenedor, self)

        self._construir_barra_estado()
        self.show_page(PAGINA_INICIAL)

    def _construir_sidebar(self, parent):
        self.sidebar = tk.Frame(parent, bg=PALETTE["sidebar_bg"], width=230)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        tk.Frame(parent, bg=PALETTE["border"], width=1).pack(side=tk.LEFT, fill=tk.Y)

        cabecera = tk.Frame(self.sidebar, bg=PALETTE["sidebar_bg"])
        cabecera.pack(fill=tk.X, pady=(28, 12))
        tk.Label(cabecera, text=APP_NAME, bg=PALETTE["sidebar_bg"], fg=PALETTE["text"], font=fuente_titulo(13),
                 anchor="w", justify="left", wraplength=190).pack(fill=tk.X, padx=24)
        tk.Label(cabecera, text=f"v{APP_VERSION}", bg=PALETTE["sidebar_bg"], fg=PALETTE["text_muted"],
                 font=(fuente(), 9), anchor="w").pack(fill=tk.X, padx=24, pady=(2, 0))

        for seccion, items in NAVEGACION:
            tk.Label(self.sidebar, text=seccion.upper(), bg=PALETTE["sidebar_bg"], fg=PALETTE["sidebar_section"],
                     font=(fuente(), 8, "bold"), anchor="w").pack(fill=tk.X, padx=28, pady=(18, 6))
            for page_id, texto, _ in items:
                self._boton_nav(page_id, texto)

    def _boton_nav(self, page_id, texto):
        fila = tk.Frame(self.sidebar, bg=PALETTE["sidebar_bg"])
        fila.pack(fill=tk.X, padx=12, pady=1)
        indicador = tk.Frame(fila, bg=PALETTE["sidebar_bg"], width=3)
        indicador.pack(side=tk.LEFT, fill=tk.Y)
        btn = tk.Button(fila, text=texto, bg=PALETTE["sidebar_bg"], fg=PALETTE["sidebar_fg"],
                        activebackground=PALETTE["sidebar_hover"], activeforeground=PALETTE["text"],
                        font=(fuente(), 10), bd=0, relief="flat", highlightthickness=0, anchor="w",
                        padx=14, pady=9, cursor="hand2", command=lambda: self.show_page(page_id))
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn.bind("<Enter>", lambda e: self._estilo_nav(page_id, "hover") if self.current_page != page_id else None)
        btn.bind("<Leave>", lambda e: self._estilo_nav(page_id, "normal") if self.current_page != page_id else None)
        self.nav_buttons[page_id] = (btn, indicador)

    def _estilo_nav(self, page_id, estado):
        btn, indicador = self.nav_buttons[page_id]
        activo = estado == "activo"
        bg = {"activo": PALETTE["sidebar_active"], "hover": PALETTE["sidebar_hover"]}.get(estado, PALETTE["sidebar_bg"])
        btn.configure(bg=bg, fg=PALETTE["sidebar_accent"] if activo else PALETTE["sidebar_fg"],
                      font=fuente_titulo(10) if activo else (fuente(), 10))
        indicador.configure(bg=PALETTE["sidebar_accent"] if activo else PALETTE["sidebar_bg"])

    def show_page(self, page_id):
        if self.current_page:
            self.paginas[self.current_page].pack_forget()
            self._estilo_nav(self.current_page, "normal")
        self.paginas[page_id].pack(fill=tk.BOTH, expand=True)
        self._estilo_nav(page_id, "activo")
        self.current_page = page_id

    def _construir_barra_archivo(self):
        tarjeta = tk.Frame(self.content_frame, bg=PALETTE["surface"], highlightthickness=1,
                           highlightbackground=PALETTE["border"], highlightcolor=PALETTE["border"])
        tarjeta.pack(fill=tk.X, padx=28, pady=(24, 0))
        fila = ttk.Frame(tarjeta, style="Card.TFrame", padding=(20, 12))
        fila.pack(fill=tk.BOTH)
        ttk.Label(fila, text="Archivo de trabajo", style="FieldLabelOptional.TLabel").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Entry(fila, textvariable=self.file_path, state="readonly", font=(fuente(), 10)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
        boton_secundario(fila, "Seleccionar…", self._on_abrir_archivo).pack(side=tk.LEFT)
        boton_secundario(fila, "Liberar", self._on_liberar_archivo).pack(side=tk.LEFT, padx=(8, 0))

    def _construir_barra_estado(self):
        barra = tk.Frame(self.content_frame, bg=PALETTE["surface"])
        barra.pack(fill=tk.X, side=tk.BOTTOM, pady=(20, 0))
        tk.Frame(barra, bg=PALETTE["border"], height=1).pack(fill=tk.X, side=tk.TOP)
        self.lbl_indicador = ttk.Label(barra, text="●", style="StatusDot.TLabel", foreground=PALETTE["success"])
        self.lbl_indicador.pack(side=tk.LEFT, padx=(28, 2), pady=6)
        self.lbl_estado = ttk.Label(barra, text="Listo", anchor="w", style="Status.TLabel")
        self.lbl_estado.pack(fill=tk.X, side=tk.LEFT)

    # --------------------------------------------------------- estado común
    def set_estado(self, mensaje):
        self.lbl_estado.configure(text=f" {mensaje}")

    def set_indicador(self, color):
        self.lbl_indicador.configure(foreground=color)

    def requiere_archivo(self):
        if self.file_path.get():
            return True
        messagebox.showwarning("Archivo requerido", "Primero debes seleccionar un archivo de Excel.")
        return False

    def recargar_vista_previa(self):
        if self.file_path.get():
            self.paginas["preview"].on_archivo_cargado(self.hojas_disponibles)

    # ------------------------------------------------------------- archivo
    def _on_abrir_archivo(self):
        ruta = filedialog.askopenfilename(title="Selecciona el archivo de Excel",
                                          filetypes=[("Archivos de Excel", "*.xlsx *.xlsm")])
        if not ruta:
            return
        try:
            hojas = hojas_de(ruta)
        except Exception as e:
            logger.error(f"No se pudo abrir el archivo seleccionado: {e}")
            messagebox.showerror("Error al abrir archivo", f"No se pudo abrir el archivo:\n{e}")
            return

        self.file_path.set(ruta)
        self.hojas_disponibles = hojas
        for pagina in self.paginas.values():
            pagina.on_archivo_cargado(hojas)
        self.set_estado(f"Archivo cargado: {Path(ruta).name} ({len(hojas)} hojas)")

    def _on_liberar_archivo(self):
        if any(p.ocupada for p in self.paginas.values()):
            messagebox.showwarning("Proceso en curso", "Espera a que finalice el proceso actual antes de liberar el archivo.")
            return
        self.file_path.set("")
        self.hojas_disponibles = []
        for pagina in self.paginas.values():
            pagina.on_archivo_liberado()
        gc.collect()
        self.set_estado("Sesión liberada. El archivo quedó disponible para otros programas.")