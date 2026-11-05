import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from core.config import LOG_FILE, OPTIONAL_SHEETS, PERFIL_TRIBUTARIO_ACTIVO, SHEET_LABELS, SHEET_ORDER
from core.errors import ErrorUsuario, logger
from data import perfiles_repo
from services.conciliacion import ConciliadorAuditoria
from services.reportes.acta_errores import GeneradorActaErrores, cargar_auditoria_guardada
from ui.dialogs import Dialogo, abrir_con_sistema, mostrar_guia_alertas
from ui.pages.base import Pagina
from ui.theme import PALETTE
from ui.widgets import boton_primario, boton_secundario, habilitar_boton


class EjecucionPage(Pagina):
    titulo = "Conciliar"
    descripcion = "Valida la configuración y cruza la contabilidad contra el token."

    def construir(self):
        self.encabezado()
        self.df_acta = None

        acciones = ttk.Frame(self, style="TFrame")
        acciones.pack(fill=tk.X, pady=(0, 16))
        self.btn_ejecutar = boton_primario(acciones, "Iniciar conciliación", self._on_ejecutar)
        self.btn_ejecutar.pack(side=tk.LEFT)
        self.btn_acta = boton_secundario(acciones, "Generar acta", self._on_generar_acta)
        self.btn_acta.pack(side=tk.LEFT, padx=(12, 0))
        boton_secundario(acciones, "Abrir logs", self._on_abrir_logs).pack(side=tk.RIGHT)
        boton_secundario(acciones, "?  Guía de alertas",
                         lambda: mostrar_guia_alertas(self.winfo_toplevel())).pack(side=tk.RIGHT, padx=(0, 12))

        self.progress = ttk.Progressbar(self, mode="indeterminate", style="TProgressbar")
        self.progress.pack(fill=tk.X, pady=(0, 16))

        consola = tk.Frame(self, bg=PALETTE["console_bg"], bd=0, highlightthickness=0)
        consola.pack(fill=tk.BOTH, expand=True)
        self.txt_log = tk.Text(
            consola, state="disabled", height=18, wrap="word", bg=PALETTE["console_bg"], fg=PALETTE["console_fg"],
            insertbackground="#FFFFFF", selectbackground=PALETTE["primary"], font=("Consolas", 10),
            relief="flat", borderwidth=0, padx=20, pady=20,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

    def on_archivo_liberado(self):
        self.df_acta = None

    def agregar_log(self, mensaje):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, f"{mensaje}\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")
        self.app.set_estado(mensaje)

    def _limpiar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

    def _validar(self, filtro, perfil):
        if not self.app.file_path.get():
            raise ErrorUsuario("Primero debes seleccionar un archivo de Excel.")
        if not filtro:
            raise ErrorUsuario("Debes seleccionar qué grupo de documentos procesar (Solo Recibidos, Solo Emitidos o Ambos) en Configurar auditoría.")
        if PERFIL_TRIBUTARIO_ACTIVO and not (perfil.get("tipo") and perfil.get("periodicidad")):
            raise ErrorUsuario("Debes indicar el tipo de contribuyente y la periodicidad de IVA de la empresa en Configurar auditoría.")
        for clave in SHEET_ORDER:
            if clave in OPTIONAL_SHEETS:
                continue
            valor = self.app.sheet_vars[clave].get().strip()
            if not valor:
                raise ErrorUsuario(f"Debes indicar el nombre de la hoja: {SHEET_LABELS[clave]}")
            if valor not in self.app.hojas_disponibles:
                raise ErrorUsuario(f"La hoja '{valor}' indicada para '{SHEET_LABELS[clave]}' no existe.")

    def _on_ejecutar(self):
        if self.ocupada:
            return
        config = self.app.paginas["config"]
        filtro = config.filtro_grupo.get()
        perfil = config.obtener_perfil()
        try:
            self._validar(filtro, perfil)
        except ErrorUsuario as e:
            logger.warning(f"Error de usuario al validar entradas: {e}")
            messagebox.showwarning("Datos incompletos", str(e))
            return

        s = config.obtener_seriales()
        if not s["iva"] or not s["base"]:
            if not messagebox.askyesno("Seriales incompletos", "Has dejado una de las categorías vacía. ¿Deseas continuar?"):
                return
        if not messagebox.askyesno("Confirmar", "Este proceso modificará el archivo. ¿Deseas continuar?"):
            return

        if PERFIL_TRIBUTARIO_ACTIVO:
            perfiles_repo.guardar(perfil["nit"], perfil)
        self.ocupada = True
        habilitar_boton(self.btn_ejecutar, False)
        self.app.set_indicador(PALETTE["primary"])
        self.progress.start(12)
        self._limpiar_log()
        self.app.show_page("exec")

        conciliador = ConciliadorAuditoria(
            self.app.file_path.get(), {k: v.get().strip() for k, v in self.app.sheet_vars.items()},
            seriales_iva=s["iva"], seriales_base=s["base"], seriales_base2=s["base2"], seriales_ret=s["ret"],
            seriales_reteiva=s["reteiva"], perfil=perfil,
            seriales_iva_dc=s["iva_dc"], seriales_base_dc=s["base_dc"], seriales_base2_dc=s["base2_dc"],
            filtro_grupo=filtro, progress_callback=self.app.ejecutor.en_ui(self.agregar_log),
        )
        self.app.ejecutor.lanzar(conciliador.ejecutar, self._on_exito, self._on_error, "conciliación")

    def _finalizar(self, exito):
        self.ocupada = False
        habilitar_boton(self.btn_ejecutar, True)
        self.progress.stop()
        self.app.set_indicador(PALETTE["success"] if exito else PALETTE["danger"])
        self.app.set_estado("Listo")

    def _on_exito(self, resumen):
        self._finalizar(True)
        self.agregar_log(
            f"Filas procesadas: {resumen['filas_procesadas']} | "
            f"Personales excluidas: {resumen['filas_personales']} | "
            f"Sin pareja: {resumen['filas_sin_pareja']}"
        )
        self.df_acta = resumen.get("df_completo")
        messagebox.showinfo("Completado", "La conciliación finalizó con éxito.")

    def _on_error(self, mensaje, es_sistema):
        self._finalizar(False)
        if es_sistema:
            self.agregar_log(f"ERROR DEL SISTEMA: {mensaje}")
            messagebox.showerror("Error", f"{mensaje}\n\nRevisa el log:\n{LOG_FILE}")
        else:
            self.agregar_log(f"ERROR DE USUARIO: {mensaje}")
            messagebox.showwarning("Error", mensaje)

    def _elegir_origen_acta(self):
        """'actual', 'archivo' o None (cancelado)."""
        dlg = Dialogo(self.winfo_toplevel(), "Generar acta", min_ancho=460, min_alto=180)
        ttk.Label(dlg.cuerpo, text="¿De qué auditoría quieres generar el acta?", font=("", 11, "bold")).pack(anchor="w")
        ttk.Label(dlg.cuerpo, wraplength=420, justify="left", text=(
            "• Auditoría actual: la conciliación ejecutada en esta sesión.\n"
            "• Cargar archivo auditado: un libro ya conciliado. Se usa su hoja DIAN VS CONT; si no la tiene, "
            "la hoja de auditoría (AUD-COMP / MOVS_AUD) con la columna ALERTA, y el acta trae solo las alertas tributarias."
        )).pack(anchor="w", pady=(8, 0))
        dlg.boton("Cancelar", dlg.cerrar)
        dlg.boton("Cargar archivo auditado…", lambda: dlg.cerrar("archivo"))
        actual = dlg.boton("Auditoría actual", lambda: dlg.cerrar("actual"), principal=self.df_acta is not None)
        if self.df_acta is None:
            actual.state(["disabled"])
        return dlg.mostrar()

    def _on_generar_acta(self):
        origen = self._elegir_origen_acta()
        if origen is None:
            return
        if origen == "actual":
            df, ruta_base = self.df_acta, self.app.file_path.get()
        else:
            ruta_base = filedialog.askopenfilename(
                title="Seleccionar archivo auditado", initialdir=str(Path(self.app.file_path.get() or ".").parent),
                filetypes=[("Archivos de Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
            )
            if not ruta_base:
                return
            try:
                df, descripcion = cargar_auditoria_guardada(ruta_base)
            except ErrorUsuario as e:
                messagebox.showwarning("Archivo sin auditoría", str(e))
                return
            except Exception as e:
                logger.exception("Error leyendo el archivo auditado para el acta")
                messagebox.showerror("Error", f"No se pudo leer el archivo auditado:\n{e}")
                return
            self.agregar_log(f"Acta desde archivo auditado: {Path(ruta_base).name} ({descripcion})")
        generador = GeneradorActaErrores(df, ruta_base)
        ruta = filedialog.asksaveasfilename(
            title="Guardar Acta de Errores (PDF)", initialdir=str(Path(ruta_base).parent),
            initialfile=generador.nombre_sugerido(), defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")],
        )
        if not ruta:
            self.agregar_log("El usuario canceló el guardado del Acta PDF.")
            return

        self.btn_acta.configure(state="disabled")
        self.app.set_estado("Generando Acta de Errores...")
        try:
            generador.generar(ruta)
            self.agregar_log(f"Acta PDF guardada en: {ruta}")
            messagebox.showinfo("Acta Generada", f"El acta se guardó correctamente en:\n{ruta}")
        except Exception as e:
            logger.exception("Error al generar el acta de errores")
            self.agregar_log(f"Error al generar Acta: {e}")
            messagebox.showerror("Error", f"No se pudo generar el acta:\n{e}")
        finally:
            self.btn_acta.configure(state="normal")
            self.app.set_estado("Listo")

    def _on_abrir_logs(self):
        carpeta = LOG_FILE.parent
        carpeta.mkdir(parents=True, exist_ok=True)
        try:
            abrir_con_sistema(carpeta)
        except Exception as e:
            logger.warning(f"No se pudo abrir la carpeta de logs: {e}")
            messagebox.showinfo("Carpeta de logs", f"Los logs se encuentran en:\n{carpeta}")