import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.errors import logger
from data.autorretenedores_repo import GestorAutorretenedores
from ui.pages.base import Pagina
from ui.widgets import boton_primario, boton_secundario, card

# (clave, encabezado, ancho de columna)
CAMPOS = (
    ("nit", "NIT", 120),
    ("nombre", "Nombre o razón social", 380),
    ("resolucion", "N.º resolución", 120),
    ("fecha", "Fecha resolución", 130),
    ("comentario", "Comentario", 180),
)
# clave: (fila, columna de la etiqueta, columnas que ocupa el campo, ancho)
DISPOSICION = {
    "nit": (0, 0, 1, 16), "nombre": (0, 2, 3, 40),
    "resolucion": (1, 0, 1, 16), "fecha": (1, 2, 1, 16), "comentario": (1, 4, 1, 24),
}
OBLIGATORIOS = ("nit", "nombre")


class AutorretenedoresPage(Pagina):
    titulo = "Autorretenedores"
    descripcion = "Listado de autorretenedores del impuesto sobre la renta que usa la conciliación."

    def construir(self):
        self.encabezado()
        self._registros = []
        self._items = {}
        self._nit_edicion = None
        self.vars = {clave: tk.StringVar() for clave, *_ in CAMPOS}
        self._construir_formulario()
        self._construir_listado()
        self._limpiar_formulario()
        self.recargar()

    # ---------------------------------------------------------- formulario
    def _construir_formulario(self):
        c = card(self, fill=tk.X, pady=(0, 20))
        self.lbl_form = ttk.Label(c, text="", style="CardHeader.TLabel")
        self.lbl_form.pack(anchor="w")
        ttk.Label(c, text="NIT sin dígito de verificación. Fecha en formato AAAA-MM-DD o DD/MM/AAAA.",
                  style="SubheaderCard.TLabel").pack(anchor="w", pady=(4, 12))

        grid = ttk.Frame(c, style="Card.TFrame")
        grid.pack(fill=tk.X)
        self.entries = {}
        for clave, etiqueta, _ in CAMPOS:
            fila, col, span, ancho = DISPOSICION[clave]
            obligatorio = clave in OBLIGATORIOS
            ttk.Label(grid, text=etiqueta + ("  *" if obligatorio else ""),
                      style="FieldLabel.TLabel" if obligatorio else "FieldLabelOptional.TLabel").grid(
                row=fila, column=col, sticky="w", padx=(0 if col == 0 else 20, 10), pady=6)
            entry = ttk.Entry(grid, textvariable=self.vars[clave], width=ancho)
            entry.grid(row=fila, column=col + 1, columnspan=span, sticky="ew", pady=6)
            entry.bind("<Return>", lambda _e: self._on_guardar())
            entry.bind("<Escape>", lambda _e: self._limpiar_formulario())
            self.entries[clave] = entry
        for col in (1, 3, 5):
            grid.columnconfigure(col, weight=1)

        acciones = ttk.Frame(c, style="Card.TFrame")
        acciones.pack(fill=tk.X, pady=(16, 0))
        self.btn_guardar = boton_primario(acciones, "Agregar", self._on_guardar)
        self.btn_guardar.pack(side=tk.LEFT)
        self.btn_cancelar = boton_secundario(acciones, "Limpiar", self._limpiar_formulario)
        self.btn_cancelar.pack(side=tk.LEFT, padx=(8, 0))

    def _limpiar_formulario(self):
        self._nit_edicion = None
        for var in self.vars.values():
            var.set("")
        self.vars["comentario"].set(GestorAutorretenedores.COMENTARIO_DEFECTO)
        self.lbl_form.configure(text="Nuevo autorretenedor")
        self.btn_guardar.configure(text="Agregar")
        self.btn_cancelar.configure(text="Limpiar")

    def _cargar_en_formulario(self, registro):
        self._nit_edicion = registro["nit"]
        for clave, var in self.vars.items():
            var.set(registro.get(clave, ""))
        self.lbl_form.configure(text=f"Editar autorretenedor · NIT {registro['nit']}")
        self.btn_guardar.configure(text="Guardar cambios")
        self.btn_cancelar.configure(text="Cancelar edición")
        self.entries["nombre"].focus_set()

    # ------------------------------------------------------------- listado
    def _construir_listado(self):
        lista = card(self, fill=tk.BOTH, expand=True)
        cabecera = ttk.Frame(lista, style="Card.TFrame")
        cabecera.pack(fill=tk.X)
        ttk.Label(cabecera, text="Registrados", style="CardHeader.TLabel").pack(side=tk.LEFT)
        self.lbl_total = ttk.Label(cabecera, text="", style="SubheaderCard.TLabel")
        self.lbl_total.pack(side=tk.RIGHT)

        busqueda = ttk.Frame(lista, style="Card.TFrame")
        busqueda.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(busqueda, text="Buscar", style="FieldLabelOptional.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        self.busqueda_var = tk.StringVar()
        self.busqueda_var.trace_add("write", lambda *_: self._mostrar())
        ttk.Entry(busqueda, textvariable=self.busqueda_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        marco = ttk.Frame(lista, style="Card.TFrame")
        marco.pack(fill=tk.BOTH, expand=True, pady=12)
        self.tree = ttk.Treeview(marco, show="headings", columns=[c for c, *_ in CAMPOS], selectmode="extended")
        for clave, encabezado, ancho in CAMPOS:
            self.tree.heading(clave, text=encabezado, anchor="w")
            self.tree.column(clave, width=ancho, anchor="w", stretch=clave == "nombre")
        sy = ttk.Scrollbar(marco, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sy.set)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", lambda _e: self._on_editar())
        self.tree.bind("<Delete>", lambda _e: self._on_eliminar())

        botones = ttk.Frame(lista, style="Card.TFrame")
        botones.pack(fill=tk.X)
        boton_secundario(botones, "Editar", self._on_editar).pack(side=tk.LEFT)
        boton_secundario(botones, "Eliminar", self._on_eliminar).pack(side=tk.LEFT, padx=(8, 0))
        boton_secundario(botones, "Exportar a Excel", self._on_exportar).pack(side=tk.RIGHT)

    def recargar(self):
        """Lee el JSON una sola vez; el filtrado posterior trabaja en memoria."""
        try:
            self._registros = sorted(GestorAutorretenedores.obtener_registros(), key=lambda r: r["nombre"])
        except (FileNotFoundError, ConnectionError, PermissionError) as e:
            logger.error(f"Listado de autorretenedores no disponible: {e}")
            messagebox.showerror("Archivo no disponible", f"{e}\n\nRuta esperada:\n{GestorAutorretenedores.ARCHIVO_DATOS}")
            self._registros = []
        except ValueError as e:
            logger.error(f"JSON corrupto: {e}")
            messagebox.showerror("Archivo inválido", str(e))
            self._registros = []
        self._mostrar()

    def _mostrar(self):
        self.tree.delete(*self.tree.get_children())
        self._items = {}
        filtro = self.busqueda_var.get().strip().lower()
        for r in self._registros:
            valores = tuple(r.get(clave, "") for clave, *_ in CAMPOS)
            if not filtro or any(filtro in v.lower() for v in valores):
                self._items[self.tree.insert("", tk.END, values=valores)] = r
        total, visibles = len(self._registros), len(self._items)
        self.lbl_total.configure(text=f"{visibles} de {total} registros" if filtro else f"{total} registros")

    # ------------------------------------------------------------ acciones
    def _operar(self, accion):
        try:
            return accion()
        except ValueError as e:
            messagebox.showwarning("Revisa los datos", str(e))
        except (FileNotFoundError, ConnectionError, PermissionError) as e:
            logger.error(f"Listado de autorretenedores no disponible: {e}")
            messagebox.showerror("Archivo no disponible", str(e))
        except Exception as e:
            logger.exception("Error al modificar autorretenedores")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")
        return None

    def _on_guardar(self):
        registro = {clave: var.get() for clave, var in self.vars.items()}
        if self._nit_edicion:
            nit = self._nit_edicion
            r = self._operar(lambda: GestorAutorretenedores.actualizar(nit, registro))
            verbo = "actualizado"
        else:
            r = self._operar(lambda: GestorAutorretenedores.agregar(registro))
            verbo = "agregado"
        if r:
            self.app.set_estado(f"Autorretenedor {verbo}: {r['nombre']} (NIT {r['nit']})")
            self._limpiar_formulario()
            self.recargar()

    def _seleccionados(self):
        return [self._items[i] for i in self.tree.selection() if i in self._items]

    def _on_editar(self):
        seleccion = self._seleccionados()
        if len(seleccion) != 1:
            messagebox.showwarning("Editar", "Selecciona un único autorretenedor para editar.")
            return
        self._cargar_en_formulario(seleccion[0])

    def _on_eliminar(self):
        seleccion = self._seleccionados()
        if not seleccion:
            messagebox.showwarning("Eliminar", "Selecciona al menos un autorretenedor de la lista.")
            return
        pregunta = (f"¿Eliminar a {seleccion[0]['nombre']} (NIT {seleccion[0]['nit']})?" if len(seleccion) == 1
                    else f"¿Eliminar {len(seleccion)} autorretenedores seleccionados?")
        if not messagebox.askyesno("Confirmar eliminación", pregunta):
            return
        nits = [r["nit"] for r in seleccion]
        eliminados = self._operar(lambda: GestorAutorretenedores.eliminar(nits))
        if eliminados is None:
            return
        if self._nit_edicion in nits:
            self._limpiar_formulario()
        self.app.set_estado(f"Autorretenedores eliminados: {eliminados}")
        self.recargar()

    def _on_exportar(self):
        try:
            if not GestorAutorretenedores.obtener_todos():
                messagebox.showwarning("Sin datos", "No hay autorretenedores registrados para exportar.")
                return
            ruta = filedialog.asksaveasfilename(title="Exportar autorretenedores", defaultextension=".xlsx",
                                                initialfile="autorretenedores.xlsx",
                                                filetypes=[("Archivos de Excel", "*.xlsx")])
            if not ruta:
                return
            GestorAutorretenedores.exportar_excel(ruta)
            self.app.set_estado(f"Autorretenedores exportados en: {ruta}")
            messagebox.showinfo("Exportación exitosa", f"Archivo exportado correctamente en:\n{ruta}")
        except Exception as e:
            logger.exception("Error al exportar autorretenedores")
            messagebox.showerror("Error de exportación", f"No se pudo exportar el archivo:\n{e}")