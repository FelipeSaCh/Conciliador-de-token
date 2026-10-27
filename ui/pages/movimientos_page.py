import tkinter as tk
from tkinter import messagebox, ttk

from services.movimientos_service import TransformadorMovimientos
from ui.pages.base import Pagina, asignar_hojas
from ui.theme import PALETTE
from ui.widgets import ScrollableChecklist, boton_primario, card, encabezado_card, habilitar_boton

COLUMNAS_GRID_TIPOS = 6


class MovimientosPage(Pagina):
    titulo = "Movimientos contables"
    descripcion = "Reestructura los movimientos en la hoja MOVS_AUD según los tipos de documento elegidos."

    def construir(self):
        self.encabezado()
        self.tipo_vars = {}

        c = card(self, fill=tk.X, pady=(0, 20))
        encabezado_card(c, "Hoja de origen", "Selecciona la hoja de movimientos para cargar sus tipos documentales.")
        fila = ttk.Frame(c, style="Card.TFrame")
        fila.pack(fill=tk.X)
        ttk.Label(fila, text="Hoja de movimientos", style="FieldLabel.TLabel").pack(side=tk.LEFT, padx=(0, 30))
        self.hoja_var = tk.StringVar()
        self.combo = ttk.Combobox(fila, textvariable=self.hoja_var, width=50, state="readonly", cursor="hand2")
        self.combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.combo.bind("<<ComboboxSelected>>", lambda _e: self._refrescar_tipos())

        tipos = card(self, fill=tk.BOTH, expand=True, pady=(0, 20))
        encabezado_card(tipos, "Tipos de documento a procesar",
                        "Selecciona los prefijos que se incluirán y pivotarán en la hoja MOVS_AUD.", pady_sub=(4, 12))
        botones = ttk.Frame(tipos, style="Card.TFrame")
        botones.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(botones, text="Marcar todos", command=lambda: self._marcar_todos(True), style="Toggle.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(botones, text="Desmarcar todos", command=lambda: self._marcar_todos(False), style="Toggle.TButton").pack(side=tk.LEFT)
        self.lista = ScrollableChecklist(tipos, height=180)
        self.lista.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self._poblar_tipos([])

        acciones = ttk.Frame(self, style="TFrame")
        acciones.pack(fill=tk.X, pady=(0, 20))
        self.btn = boton_primario(acciones, "Transformar y exportar", self._on_pivotar)
        self.btn.pack(side=tk.LEFT)
        self.lbl_estado = ttk.Label(acciones, text="", style="Subheader.TLabel")
        self.lbl_estado.pack(side=tk.LEFT, padx=(20, 0))

        progreso = ttk.Frame(self, style="TFrame")
        progreso.pack(fill=tk.X, pady=(0, 20))
        self.progress = ttk.Progressbar(progreso, mode="determinate", maximum=100, style="TProgressbar")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_pct = ttk.Label(progreso, text="", style="Subheader.TLabel", width=6, anchor="e")
        self.lbl_pct.pack(side=tk.LEFT, padx=(12, 0))

    def on_archivo_cargado(self, hojas):
        asignar_hojas(self.combo, self.hoja_var, hojas)
        if self.hoja_var.get() not in hojas:
            self.hoja_var.set(hojas[0] if hojas else "")

    def on_archivo_liberado(self):
        self.combo['values'] = []
        self.hoja_var.set("")
        self.lbl_estado.configure(text="")
        self._reiniciar_progreso()
        self._poblar_tipos([])

    def _marcar_todos(self, valor):
        for var in self.tipo_vars.values():
            var.set(valor)

    def _refrescar_tipos(self):
        ruta, hoja = self.app.file_path.get(), self.hoja_var.get().strip()
        self._poblar_tipos([])
        if not ruta or not hoja or hoja not in self.app.hojas_disponibles:
            return
        self.lista.limpiar()
        self.lista.mensaje("Extrayendo tipos únicos...")

        def exito(resultado):
            tipos, formato = resultado
            self._poblar_tipos(tipos)
            self.lbl_estado.configure(text=f"Formato detectado: {formato}")

        def error(mensaje, es_sistema):
            self._poblar_tipos([])
            if not es_sistema:
                self.lbl_estado.configure(text=mensaje)

        self.app.ejecutor.lanzar(lambda: TransformadorMovimientos.obtener_tipos(ruta, hoja), exito, error,
                                 f"extracción de tipos de '{hoja}'")

    def _poblar_tipos(self, tipos):
        self.lista.limpiar()
        self.tipo_vars = {}
        if not tipos:
            self.lista.mensaje("Selecciona una hoja válida de movimientos (formato estándar o auxiliar por documento).")
        else:
            grid = tk.Frame(self.lista.inner, bg=PALETTE["surface"])
            grid.pack(fill=tk.X, padx=8, pady=8)
            for i, tipo in enumerate(tipos):
                var = tk.BooleanVar(value=tipo.startswith(TransformadorMovimientos.PREFIJOS_PERMITIDOS))
                self.tipo_vars[tipo] = var
                ttk.Checkbutton(grid, text=tipo, variable=var, style="Card.TCheckbutton").grid(
                    row=i // COLUMNAS_GRID_TIPOS, column=i % COLUMNAS_GRID_TIPOS, sticky="w", padx=(0, 20), pady=6)
        self.lista.actualizar_scrollregion()

    def _on_pivotar(self):
        if self.ocupada or not self.app.requiere_archivo():
            return
        hoja = self.hoja_var.get().strip()
        if not hoja:
            messagebox.showwarning("Hoja requerida", "Debes seleccionar la hoja de movimientos.")
            return
        seleccion = [t for t, v in self.tipo_vars.items() if v.get()]
        if not seleccion:
            messagebox.showwarning("Tipos requeridos", "Debes seleccionar al menos un tipo de documento para procesar.")
            return
        if not messagebox.askyesno("Confirmar pivoteo", "Se agregará una nueva hoja 'MOVS_AUD' a este archivo. ¿Deseas continuar?"):
            return

        self.ocupada = True
        habilitar_boton(self.btn, False)
        self.lbl_estado.configure(text="Transformando movimientos...")
        self._set_avance(0)
        transformador = TransformadorMovimientos(
            self.app.file_path.get(), hoja, prefijos_permitidos=seleccion,
            progress_callback=self.app.ejecutor.en_ui(lambda m: self.lbl_estado.configure(text=m)),
            avance_callback=self.app.ejecutor.en_ui(self._set_avance),
        )
        self.app.ejecutor.lanzar(transformador.ejecutar, self._on_exito, self._on_error, "pivoteo de movimientos")

    def _set_avance(self, valor):
        if valor is None:
            if str(self.progress["mode"]) != "indeterminate":
                self.progress.configure(mode="indeterminate")
                self.progress.start(12)
            return
        if str(self.progress["mode"]) != "determinate":
            self.progress.stop()
            self.progress.configure(mode="determinate")
        self.progress["value"] = valor
        self.lbl_pct.configure(text=f"{int(valor)}%")

    def _reiniciar_progreso(self):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.lbl_pct.configure(text="")

    def _finalizar(self, exito):
        self.ocupada = False
        self._set_avance(100) if exito else self._reiniciar_progreso()
        habilitar_boton(self.btn, True)

    def _on_exito(self, ruta):
        self._finalizar(True)
        self.lbl_estado.configure(text="Transformación exitosa.")
        messagebox.showinfo("Éxito", f"El archivo ha sido guardado en:\n{ruta}")

    def _on_error(self, mensaje, es_sistema):
        self._finalizar(False)
        if es_sistema:
            self.lbl_estado.configure(text="Error de procesamiento")
            messagebox.showerror("Error", mensaje)
        else:
            self.lbl_estado.configure(text=f"Error: {mensaje}")
            messagebox.showwarning("Faltan datos", mensaje)