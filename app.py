import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import sv_ttk
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
)
from errors import ErrorSistema, ErrorUsuario, logger
from excel_engine import ConciliadorExcel
from preview_widget import VistaPreviaExcel


class ConciliadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        sv_ttk.set_theme("light")
        self.title(f"{APP_NAME} - v{APP_VERSION}")
        self.geometry("800x680")
        self.minsize(760, 460)
        self.set_app_icon()
        self.file_path = tk.StringVar()
        self.sheet_vars = {clave: tk.StringVar(value=nombre) for clave, nombre in DEFAULT_SHEET_NAMES.items()}
        self.hojas_disponibles = []
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
        """Configura el tema ttk y define estilos globales de UX/UI."""
        self.style = ttk.Style(self)
        
        # Tipografía base
        font_family = "Segoe UI" if "win32" in self.tk.call("tk", "windowingsystem") else "Helvetica"
        
        # Estilos personalizados LIMPIOS (Sin forzar 'vista' para que no choque con sv_ttk)
        self.style.configure(".", font=(font_family, 9))
        self.style.configure("Header.TLabel", font=(font_family, 11, "bold"), foreground="#111827")
        self.style.configure("Subheader.TLabel", font=(font_family, 9), foreground="#6B7280")
        self.style.configure("Card.TFrame", background="#FFFFFF")
        
        self.style.configure("Action.TButton", font=(font_family, 9, "bold"), padding=(10, 4))
        self.style.configure("TButton", padding=(8, 4))
        self.style.configure("Status.TLabel", font=(font_family, 8), padding=(12, 6), background="#F3F4F6", foreground="#374151")

    def _build_ui(self):
        self._build_barra_superior()

        # Línea divisoria sutil para separar el header del contenido
        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=12)

        # Aumentamos el padding del notebook para que respire
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=(12, 16))

        self.tab_config = ttk.Frame(self.notebook)
        self.tab_preview = ttk.Frame(self.notebook)
        self.tab_ejecucion = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_config, text="  ⚙ Configuración  ")
        self.notebook.add(self.tab_preview, text="  👁 Vista Previa  ")
        self.notebook.add(self.tab_ejecucion, text="  ▶ Ejecución y Logs  ")

        self._build_tab_config()
        self._build_tab_preview()
        self._build_tab_ejecucion()

        self._build_barra_estado()

    def _build_barra_superior(self):
        frame_top = ttk.Frame(self, padding=(16, 16, 16, 12))
        frame_top.pack(fill=tk.X)

        lbl_instruccion = ttk.Label(
            frame_top, 
            text="Archivo de trabajo:", 
            style="Header.TLabel"
        )
        lbl_instruccion.pack(side=tk.LEFT, padx=(0, 12))

        entry_path = ttk.Entry(
            frame_top, 
            textvariable=self.file_path, 
            state="readonly"
        )
        entry_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

        btn_browse = ttk.Button(
            frame_top, 
            text="Examinar...", 
            style="Action.TButton",
            command=self._on_abrir_archivo,
            cursor="hand2"
        )
        btn_browse.pack(side=tk.LEFT)

    def _build_tab_config(self):
        wrapper = ttk.Frame(self.tab_config, padding=24)
        wrapper.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(wrapper)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        ttk.Label(
            header_frame,
            text="Asignación de Estructura de Hojas",
            style="Header.TLabel"
        ).pack(anchor="w")
        
        ttk.Label(
            header_frame,
            text="Relaciona cada sección requerida con el nombre exacto de la pestaña cargada en tu archivo.",
            style="Subheader.TLabel"
        ).pack(anchor="w", pady=(4, 0))

        grid_frame = ttk.Frame(wrapper)
        grid_frame.pack(fill=tk.X, anchor="n")

        self.combos_hojas = {}
        for i, clave in enumerate(SHEET_ORDER):
            etiqueta = SHEET_LABELS[clave]
            es_opcional = clave in OPTIONAL_SHEETS
            
            lbl_texto = f"{etiqueta}" if es_opcional else f"{etiqueta} *"
            
            lbl = ttk.Label(
                grid_frame, 
                text=lbl_texto, 
                font=("Segoe UI", 9, "normal" if es_opcional else "bold")
            )
            lbl.grid(row=i, column=0, sticky="w", pady=8, padx=(0, 20))

            combo = ttk.Combobox(
                grid_frame, 
                textvariable=self.sheet_vars[clave], 
                width=45, 
                state="readonly",
                cursor="hand2"
            )
            combo.grid(row=i, column=1, sticky="ew", pady=8)
            self.combos_hojas[clave] = combo

        grid_frame.columnconfigure(1, weight=1)

    def _build_tab_preview(self):
        # ESTE MÉTODO FALTABA EN TU CÓDIGO PEGADO
        self.vista_previa = VistaPreviaExcel(self.tab_preview)
        self.vista_previa.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    def _build_tab_ejecucion(self):
        contenedor = ttk.Frame(self.tab_ejecucion, padding=20)
        contenedor.pack(fill=tk.BOTH, expand=True)

        panel_acciones = ttk.Frame(contenedor)
        panel_acciones.pack(fill=tk.X, pady=(0, 16))

        self.btn_ejecutar = ttk.Button(
            panel_acciones, 
            text="▶ Iniciar Conciliación", 
            style="Action.TButton",
            command=self._on_ejecutar,
            cursor="hand2"
        )
        self.btn_ejecutar.pack(side=tk.LEFT)

        btn_logs = ttk.Button(
            panel_acciones, 
            text="📁 Abrir Carpeta de Logs", 
            command=self._on_abrir_logs,
            cursor="hand2"
        )
        btn_logs.pack(side=tk.LEFT, padx=(12, 0))

        self.progress = ttk.Progressbar(contenedor, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(0, 16))

        lbl_console = ttk.Label(contenedor, text="Registro de Actividad (Console output):", style="Header.TLabel")
        lbl_console.pack(anchor="w", pady=(0, 8))

        frame_consola = tk.Frame(contenedor, bg="#1E1E1E", bd=1, relief="sunken")
        frame_consola.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(
            frame_consola, 
            state="disabled", 
            height=18, 
            wrap="word", 
            bg="#1E1E1E", 
            fg="#E5E5E5",
            insertbackground="#FFFFFF",
            selectbackground="#264F78",
            font=("Consolas", 10),
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=12
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

    def _build_barra_estado(self):
        self.lbl_estado = ttk.Label(
            self, 
            text="Estado: Listo", 
            anchor="w",
            style="Status.TLabel"
        )
        self.lbl_estado.pack(fill=tk.X, side=tk.BOTTOM)

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
            combo['values'] = self.hojas_disponibles
            actual = self.sheet_vars[clave].get()
            if actual not in self.hojas_disponibles:
                if clave in OPTIONAL_SHEETS:
                    self.sheet_vars[clave].set("")
                elif self.hojas_disponibles:
                    self.sheet_vars[clave].set(self.hojas_disponibles[0])

        self.vista_previa.cargar_archivo(ruta, self.hojas_disponibles)
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

        respuesta = messagebox.askyesno(
            "Confirmar ejecución",
            "Este proceso modificará el archivo de Excel seleccionado (se reescribirán "
            "hojas de resultados). ¿Deseas continuar?"
        )
        if not respuesta:
            return

        self._procesando = True
        self.btn_ejecutar.configure(state="disabled")
        self.progress.start(12)
        self._limpiar_log()
        self.notebook.select(self.tab_ejecucion)

        sheet_names = {clave: var.get().strip() for clave, var in self.sheet_vars.items()}
        hilo = threading.Thread(
            target=self._ejecutar_en_hilo, args=(self.file_path.get(), sheet_names), daemon=True
        )
        hilo.start()

    def _ejecutar_en_hilo(self, ruta, sheet_names):
        conciliador = ConciliadorExcel(ruta, sheet_names, progress_callback=self._on_progreso)
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
                    self._finalizar_ejecucion()
                    self._agregar_log(
                        f"Filas procesadas: {payload['filas_procesadas']} | "
                        f"Personales excluidas: {payload['filas_personales']} | "
                        f"Sin pareja: {payload['filas_sin_pareja']}"
                    )
                    messagebox.showinfo("Proceso completado", "La conciliación finalizó con éxito.")
                elif tipo == "error_usuario":
                    self._finalizar_ejecucion()
                    self._agregar_log(f"ERROR DE USUARIO: {payload}")
                    messagebox.showwarning("Error en los datos", payload)
                elif tipo == "error_sistema":
                    self._finalizar_ejecucion()
                    self._agregar_log(f"ERROR DEL SISTEMA: {payload}")
                    messagebox.showerror(
                        "Error del sistema",
                        f"{payload}\n\nRevisa el log para más detalles:\n{LOG_FILE}"
                    )
        except queue.Empty:
            pass
        finally:
            self.after(300, self._procesar_cola)

    def _finalizar_ejecucion(self):
        self._procesando = False
        self.btn_ejecutar.configure(state="normal")
        self.progress.stop()
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
        self.lbl_estado.configure(text=f" Estado: {mensaje}")