import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from ui.theme import PALETTE

WIDGETS_CON_SCROLL_PROPIO = (ttk.Treeview, tk.Listbox, tk.Text, tk.Canvas)


def abrir_con_sistema(ruta):
    if sys.platform.startswith("win"):
        os.startfile(ruta)
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(ruta)])


def centrar(ventana, parent):
    ventana.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - ventana.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - ventana.winfo_height()) // 2
    ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")


# ============================================================== base común de diálogos
class Dialogo:
    """Ventana emergente modal con cuerpo desplazable (vertical/horizontal) y pie fijo para los botones.

    ancho/alto: píxeles; None = tamaño natural del contenido. En ambos casos se limita al 90 % de la pantalla y,
    si el contenido no cabe, aparecen las barras de desplazamiento. La rueda del mouse desplaza el cuerpo salvo sobre tablas y
    listas, que conservan su propio desplazamiento.
    """

    def __init__(self, parent, titulo, ancho=None, alto=None, min_ancho=420, min_alto=280):
        self.parent = parent
        self.resultado = None
        self.ventana = v = tk.Toplevel(parent)
        v.title(titulo)
        v.transient(parent)
        v.configure(bg=PALETTE["bg"])

        self._tamano = (ancho, alto)
        self._minimo = (min_ancho, min_alto)

        self.pie = ttk.Frame(v, padding=(16, 10))
        self.pie.pack(side="bottom", fill="x")
        ttk.Separator(v).pack(side="bottom", fill="x")

        zona = ttk.Frame(v)
        zona.pack(fill="both", expand=True)
        zona.rowconfigure(0, weight=1)
        zona.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(zona, bg=PALETTE["bg"], highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self._sy = ttk.Scrollbar(zona, orient="vertical", command=self.canvas.yview)
        self._sx = ttk.Scrollbar(zona, orient="horizontal", command=self.canvas.xview)
        self._sy.grid(row=0, column=1, sticky="ns")
        self._sx.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=self._sy.set, xscrollcommand=self._sx.set)

        self.cuerpo = ttk.Frame(self.canvas, padding=16)
        self._ventana_cuerpo = self.canvas.create_window((0, 0), window=self.cuerpo, anchor="nw")
        self.cuerpo.bind("<Configure>", self._ajustar)
        self.canvas.bind("<Configure>", self._ajustar)

        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            v.bind(seq, self._rueda, add="+")
        v.bind("<Shift-MouseWheel>", lambda e: self._rueda(e, horizontal=True), add="+")
        v.bind("<Escape>", lambda _e: self.cerrar())
        v.protocol("WM_DELETE_WINDOW", self.cerrar)

    def _dimensionar(self):
        v = self.ventana
        v.update_idletasks()
        natural_w = self.cuerpo.winfo_reqwidth() + self._sy.winfo_reqwidth()
        natural_h = self.cuerpo.winfo_reqheight() + self.pie.winfo_reqheight() + 4
        ancho, alto = (valor or natural for valor, natural in zip(self._tamano, (natural_w, natural_h)))
        ancho = min(ancho, int(v.winfo_screenwidth() * 0.9))
        alto = min(alto, int(v.winfo_screenheight() * 0.85))
        v.geometry(f"{ancho}x{alto}")
        v.minsize(min(self._minimo[0], ancho), min(self._minimo[1], alto))

    def _ajustar(self, _e=None):
        req_w, req_h = self.cuerpo.winfo_reqwidth(), self.cuerpo.winfo_reqheight()
        vis_w, vis_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        ancho, alto = max(req_w, vis_w), max(req_h, vis_h)
        self.canvas.itemconfigure(self._ventana_cuerpo, width=ancho, height=alto)
        self.canvas.configure(scrollregion=(0, 0, ancho, alto))
        for barra, sobra in ((self._sy, req_h > vis_h), (self._sx, req_w > vis_w)):
            barra.grid() if sobra else barra.grid_remove()

    def _rueda(self, event, horizontal=False):
        try:
            widget = self.ventana.winfo_containing(event.x_root, event.y_root)
        except (KeyError, tk.TclError):
            return
        if (widget is None or widget.winfo_toplevel() is not self.ventana
                or isinstance(widget, WIDGETS_CON_SCROLL_PROPIO) and widget is not self.canvas):
            return
        barra = self._sx if horizontal else self._sy
        if not barra.winfo_ismapped():
            return
        num = getattr(event, "num", None)
        pasos = -3 if num == 4 else 3 if num == 5 else int(-1 * (event.delta / 120) * 3)
        (self.canvas.xview_scroll if horizontal else self.canvas.yview_scroll)(pasos, "units")

    def boton(self, texto, comando, lado="right", principal=False):
        btn = ttk.Button(self.pie, text=texto, command=comando)
        btn.pack(side=lado, padx=(6, 0) if lado == "right" else (0, 6))
        if principal:
            self.ventana.bind("<Return>", lambda _e: comando())
        return btn

    def cerrar(self, resultado=None):
        if resultado is not None:
            self.resultado = resultado
        self.ventana.destroy()

    def mostrar(self):
        previo = self.ventana.grab_current()
        self._dimensionar()
        centrar(self.ventana, self.parent)
        self.ventana.grab_set()
        self.ventana.focus_set()
        self.parent.wait_window(self.ventana)
        if previo is not None:
            try:
                previo.grab_set()
            except tk.TclError:
                pass
        return self.resultado


def ajustar_columnas(arbol, pesos, minimos=None):
    """Reparte el ancho de la tabla entre sus columnas según `pesos` cada vez que cambia de tamaño."""
    minimos = minimos or {}
    total_pesos = sum(pesos.values())

    def repartir(event):
        disponible = max(event.width - 4, 0)
        for col, peso in pesos.items():
            arbol.column(col, width=max(int(disponible * peso / total_pesos), minimos.get(col, 60)), stretch=False)

    for col in pesos:
        arbol.column(col, minwidth=minimos.get(col, 60))
    arbol.bind("<Configure>", repartir, add="+")


def tabla_con_scroll(parent, columnas, pesos, minimos=None, ancho_min=420, alto_min=260):
    """Treeview con barras vertical y horizontal y columnas ajustadas al ancho. columnas: [(id, título)]."""
    # Tamaño pedido fijo: la tabla ocupa el espacio asignado sin empujar el ancho del diálogo (sin bucle de resize)
    marco = ttk.Frame(parent, width=ancho_min, height=alto_min)
    marco.grid_propagate(False)
    marco.rowconfigure(0, weight=1)
    marco.columnconfigure(0, weight=1)
    arbol = ttk.Treeview(marco, columns=[c for c, _ in columnas], show="headings")
    for col, titulo in columnas:
        arbol.heading(col, text=titulo)
        arbol.column(col, anchor="w")
    sy = ttk.Scrollbar(marco, orient="vertical", command=arbol.yview)
    sx = ttk.Scrollbar(marco, orient="horizontal", command=arbol.xview)
    arbol.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
    arbol.grid(row=0, column=0, sticky="nsew")
    sy.grid(row=0, column=1, sticky="ns")
    sx.grid(row=1, column=0, sticky="ew")
    ajustar_columnas(arbol, pesos, minimos)
    return marco, arbol


def lista_con_scroll(parent, alto=8):
    """Listbox con barra vertical que se estira al ancho disponible."""
    marco = ttk.Frame(parent)
    marco.columnconfigure(0, weight=1)
    marco.rowconfigure(0, weight=1)
    lb = tk.Listbox(marco, height=alto, exportselection=False, relief="solid", bd=1, highlightthickness=0,
                    selectbackground=PALETTE["primary_soft"], selectforeground=PALETTE["text"])
    sy = ttk.Scrollbar(marco, orient="vertical", command=lb.yview)
    lb.configure(yscrollcommand=sy.set)
    lb.grid(row=0, column=0, sticky="nsew")
    sy.grid(row=0, column=1, sticky="ns")
    return marco, lb


def _nivel_ok(codigo, nivel):
    return (nivel == "Todas" or (nivel.startswith("Cuentas") and len(codigo) == 4)
            or (nivel.startswith("Subcuentas") and len(codigo) == 6) or (nivel.startswith("Auxiliares") and len(codigo) >= 9))


def _numero(texto):
    """'52.374' / '52374' / '52374,5' -> float."""
    s = str(texto).strip().replace("$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")
    return float(s)


def _formato(valor):
    return str(int(valor)) if float(valor).is_integer() else str(valor)


# ============================================================== tipos de documento (informe de IVA)
def seleccionar_tipos_documento(parent, tipos, preseleccionados=None):
    preseleccionados = set(preseleccionados or []) & set(tipos) or set(tipos)
    dlg = Dialogo(parent, "Tipos de documento", min_ancho=360, min_alto=260)
    marco = dlg.cuerpo
    ttk.Label(marco, text="Seleccione los tipos de documento a incluir en el informe de IVA:").pack(anchor="w", pady=(0, 8))

    variables = {t: tk.BooleanVar(value=t in preseleccionados) for t in tipos}
    var_todos = tk.BooleanVar(value=all(v.get() for v in variables.values()))

    def alternar_todos():
        for v in variables.values():
            v.set(var_todos.get())

    def actualizar_todos():
        var_todos.set(all(v.get() for v in variables.values()))

    ttk.Checkbutton(marco, text="Seleccionar todos", variable=var_todos, command=alternar_todos).pack(anchor="w")
    ttk.Separator(marco).pack(fill="x", pady=6)
    for tipo, var in variables.items():
        ttk.Checkbutton(marco, text=tipo, variable=var, command=actualizar_todos).pack(anchor="w", padx=(12, 0))

    def aceptar():
        seleccion = [t for t, v in variables.items() if v.get()]
        if not seleccion:
            messagebox.showwarning("Tipos de documento", "Debe seleccionar al menos un tipo de documento.",
                                   parent=dlg.ventana)
            return
        dlg.cerrar(seleccion)

    dlg.boton("Cancelar", dlg.cerrar)
    dlg.boton("Aceptar", aceptar, principal=True)
    return dlg.mostrar()


# ============================================================== UVT y bases mínimas
def configurar_parametros_tributarios(parent):
    """Ventana para editar UVT por año y bases mínimas de retención por vigencia. Devuelve True si se guardó."""
    from core.config import CONCEPTOS_BASE_MINIMA, INFO_PARAMETROS_UVT
    from data import parametros_repo
    from ui.widgets import chip_info

    dlg = Dialogo(parent, "Parámetros tributarios: UVT y bases mínimas", min_ancho=600)
    marco = dlg.cuerpo

    cabecera = ttk.Frame(marco)
    cabecera.pack(fill="x")
    ttk.Label(cabecera, text="UVT por año", font=("", 11, "bold")).pack(side="left")
    chip_info(cabecera, "Parámetros tributarios", INFO_PARAMETROS_UVT, bg=PALETTE["bg"]).pack(side="left", padx=6)

    filas_uvt, filas_bases = [], []
    marco_uvt = ttk.Frame(marco)
    marco_uvt.pack(fill="x", pady=(6, 4))
    marco_bases = ttk.Frame(marco)

    def redibujar_uvt():
        for w in marco_uvt.winfo_children():
            w.destroy()
        ttk.Label(marco_uvt, text="Año").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(marco_uvt, text="Valor UVT ($)").grid(row=0, column=1, sticky="w", padx=4)
        for i, (v_anio, v_valor) in enumerate(filas_uvt, start=1):
            ttk.Entry(marco_uvt, textvariable=v_anio, width=8).grid(row=i, column=0, padx=4, pady=2, sticky="ew")
            ttk.Entry(marco_uvt, textvariable=v_valor, width=14).grid(row=i, column=1, padx=4, pady=2, sticky="ew")
            ttk.Button(marco_uvt, text="✕", width=3,
                       command=lambda k=i - 1: (filas_uvt.pop(k), redibujar_uvt())).grid(row=i, column=2, padx=4)

    def redibujar_bases():
        for w in marco_bases.winfo_children():
            w.destroy()
        ttk.Label(marco_bases, text="Vigente desde").grid(row=0, column=0, sticky="w", padx=4)
        for j, concepto in enumerate(CONCEPTOS_BASE_MINIMA, start=1):
            ttk.Label(marco_bases, text=concepto.capitalize()).grid(row=0, column=j, sticky="w", padx=4)
            marco_bases.columnconfigure(j, weight=1)
        for i, (v_desde, v_conceptos) in enumerate(filas_bases, start=1):
            ttk.Entry(marco_bases, textvariable=v_desde, width=11).grid(row=i, column=0, padx=4, pady=2, sticky="w")
            for j, concepto in enumerate(CONCEPTOS_BASE_MINIMA, start=1):
                ttk.Entry(marco_bases, textvariable=v_conceptos[concepto], width=6).grid(
                    row=i, column=j, padx=4, pady=2, sticky="ew")
            ttk.Button(marco_bases, text="✕", width=3,
                       command=lambda k=i - 1: (filas_bases.pop(k), redibujar_bases())).grid(
                row=i, column=len(CONCEPTOS_BASE_MINIMA) + 1, padx=4)

    def agregar_uvt(anio="", valor=""):
        filas_uvt.append((tk.StringVar(value=str(anio)), tk.StringVar(value=_formato(valor) if valor != "" else "")))

    def agregar_base(desde="", valores=None):
        valores = valores or {}
        filas_bases.append((tk.StringVar(value=desde),
                            {c: tk.StringVar(value=_formato(valores.get(c, 0))) for c in CONCEPTOS_BASE_MINIMA}))

    def cargar_en_pantalla(d):
        filas_uvt.clear()
        filas_bases.clear()
        for anio, valor in sorted(d["uvt"].items()):
            agregar_uvt(anio, valor)
        for desde, valores in d["bases"]:
            agregar_base(desde, valores)
        redibujar_uvt()
        redibujar_bases()

    def nueva_vigencia():
        ultima = {c: v.get() for c, v in filas_bases[-1][1].items()} if filas_bases else {}
        agregar_base("", {c: _numero(v or 0) for c, v in ultima.items()})
        redibujar_bases()

    ttk.Button(marco, text="+ Agregar año", command=lambda: (agregar_uvt(), redibujar_uvt())).pack(anchor="w")
    ttk.Separator(marco).pack(fill="x", pady=10)
    ttk.Label(marco, text="Bases mínimas de retención en la fuente (en UVT) por vigencia",
              font=("", 11, "bold")).pack(anchor="w")
    fila_ayuda = ttk.Frame(marco)
    fila_ayuda.pack(fill="x")
    ttk.Label(fila_ayuda, text="Fecha en AAAA-MM-DD o DD/MM/AAAA. 0 = sin base mínima.").pack(side="left")
    ttk.Button(fila_ayuda, text="Ver / editar cuentas de cada concepto",
               command=lambda: configurar_conceptos_retencion(dlg.ventana)).pack(side="right")
    marco_bases.pack(fill="x", pady=(6, 4))
    ttk.Button(marco, text="+ Agregar vigencia", command=nueva_vigencia).pack(anchor="w")

    def guardar():
        try:
            uvt = {int(a.get().strip()): _numero(v.get()) for a, v in filas_uvt if a.get().strip()}
            bases = [(d.get(), {c: _numero(var.get() or 0) for c, var in conceptos.items()})
                     for d, conceptos in filas_bases if d.get().strip()]
            ruta = parametros_repo.guardar(parametros_repo.validar(uvt, bases))
        except (ValueError, OSError) as e:
            messagebox.showwarning("Parámetros tributarios", str(e), parent=dlg.ventana)
            return
        messagebox.showinfo("Parámetros tributarios", f"Parámetros guardados en:\n{ruta}", parent=dlg.ventana)
        dlg.cerrar(True)

    def restaurar():
        if messagebox.askyesno("Parámetros tributarios", "¿Restaurar los valores de fábrica? (no se guarda hasta "
                               "presionar Guardar)", parent=dlg.ventana):
            cargar_en_pantalla(parametros_repo.por_defecto())

    dlg.boton("Restaurar valores de fábrica", restaurar, lado="left")
    dlg.boton("Cancelar", dlg.cerrar)
    dlg.boton("Guardar", guardar)
    cargar_en_pantalla(parametros_repo.cargar())
    return bool(dlg.mostrar())


# ============================================================== estructura de los diálogos de consulta de cuentas
def _panel_consulta(dlg, titulo_selector, opciones, valor_inicial, info_titulo, info_texto, niveles, columnas, pesos):
    """Arma selector superior + panel izquierdo (reglas) + panel derecho (consulta). Devuelve los widgets clave."""
    from ui.widgets import chip_info

    marco = dlg.cuerpo
    marco.columnconfigure(0, weight=2, minsize=320)
    marco.columnconfigure(1, weight=3, minsize=420)
    marco.rowconfigure(1, weight=1)

    superior = ttk.Frame(marco)
    superior.grid(row=0, column=0, columnspan=2, sticky="ew")
    ttk.Label(superior, text=titulo_selector, font=("", 11, "bold")).pack(side="left")
    var_selector = tk.StringVar(value=valor_inicial)
    combo = ttk.Combobox(superior, textvariable=var_selector, values=opciones, state="readonly", width=28)
    combo.pack(side="left", padx=8)
    chip_info(superior, info_titulo, info_texto, bg=PALETTE["bg"]).pack(side="left")

    izq = ttk.Frame(marco)
    izq.grid(row=1, column=0, sticky="nsew", padx=(0, 16), pady=(12, 0))
    izq.columnconfigure(0, weight=1)

    der = ttk.Frame(marco)
    der.grid(row=1, column=1, sticky="nsew", pady=(12, 0))
    der.columnconfigure(0, weight=1)
    der.rowconfigure(1, weight=1)
    filtros = ttk.Frame(der)
    filtros.grid(row=0, column=0, sticky="ew")
    ttk.Label(filtros, text="Buscar").pack(side="left")
    var_buscar = tk.StringVar()
    ttk.Entry(filtros, textvariable=var_buscar, width=14).pack(side="left", padx=6, fill="x", expand=True)
    var_nivel = tk.StringVar(value=niveles[1] if len(niveles) > 1 else niveles[0])
    ttk.Combobox(filtros, textvariable=var_nivel, state="readonly", width=20, values=niveles).pack(side="left", padx=6)
    marco_tabla, arbol = tabla_con_scroll(der, columnas, pesos)
    marco_tabla.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
    lbl_total = ttk.Label(der, text="", foreground=PALETTE["text_muted"])
    lbl_total.grid(row=2, column=0, sticky="w", pady=(6, 0))
    return {"combo": combo, "selector": var_selector, "izq": izq, "filtros": filtros, "buscar": var_buscar,
            "nivel": var_nivel, "arbol": arbol, "total": lbl_total}


def _listas_y_prefijo(izq, titulos):
    """Listas de prefijos + campo Prefijo con nombre PUC. Devuelve (listas, var_prefijo, marco_botones)."""
    from shared.cuentas import nombre_cuenta

    listas = {}
    fila = 0
    for clave, titulo in titulos:
        ttk.Label(izq, text=titulo, font=("", 10, "bold")).grid(row=fila, column=0, sticky="w", pady=(0, 4))
        marco, lb = lista_con_scroll(izq, alto=6)
        marco.grid(row=fila + 1, column=0, sticky="nsew", pady=(0, 10))
        izq.rowconfigure(fila + 1, weight=1)
        listas[clave] = lb
        fila += 2

    var_prefijo = tk.StringVar()
    linea = ttk.Frame(izq)
    linea.grid(row=fila, column=0, sticky="ew")
    ttk.Label(linea, text="Prefijo").pack(side="left")
    entrada = ttk.Entry(linea, textvariable=var_prefijo, width=14)
    entrada.pack(side="left", padx=6)
    lbl_nombre = ttk.Label(izq, text="", foreground=PALETTE["text_muted"], wraplength=360)
    lbl_nombre.grid(row=fila + 1, column=0, sticky="w", pady=(4, 6))
    var_prefijo.trace_add("write", lambda *_: lbl_nombre.configure(text=nombre_cuenta(var_prefijo.get().strip())))
    botones = ttk.Frame(izq)
    botones.grid(row=fila + 2, column=0, sticky="ew")
    ttk.Label(izq, text="Tip: doble clic en una cuenta de la consulta la copia al campo Prefijo.",
              foreground=PALETTE["text_muted"], wraplength=360).grid(row=fila + 3, column=0, sticky="w", pady=(10, 0))
    return listas, var_prefijo, entrada, botones


# ============================================================== categorías de cuentas (seriales)
def configurar_categorias_cuentas(parent):
    """Ventana para consultar y editar qué cuentas se marcan en cada categoría de seriales. True si se guardó."""
    from core.config import CATEGORIAS_EXCLUYENTES, CATEGORY_COLORS, ETIQUETAS_CATEGORIAS, INFO_CATEGORIAS
    from data import parametros_repo
    from shared.cuentas import clasificar, cuentas_puc, nombre_cuenta, recargar_reglas

    reglas = {c: {k: list(v) for k, v in r.items()} for c, r in parametros_repo.cargar_reglas().items()}
    puc = cuentas_puc()
    etiquetas = {ETIQUETAS_CATEGORIAS.get(c, c): c for c in CATEGORY_COLORS}

    dlg = Dialogo(parent, "Categorías de cuentas", ancho=1100, alto=640, min_ancho=820, min_alto=480)
    w = _panel_consulta(
        dlg, "Categoría", list(etiquetas), next(iter(etiquetas)), "Categorías de cuentas", INFO_CATEGORIAS,
        ["Cuentas (4 dígitos)", "Subcuentas (6 dígitos)", "Auxiliares (9+ dígitos)", "Todas"],
        [("codigo", "Código"), ("nombre", "Nombre"), ("resultado", "Resultado")],
        {"codigo": 1, "nombre": 4, "resultado": 1.6},
    )
    arbol = w["arbol"]
    arbol.tag_configure("excluida", foreground=PALETTE["danger"])
    var_excluidas = tk.BooleanVar(value=False)
    ttk.Checkbutton(w["filtros"], text="Mostrar excluidas", variable=var_excluidas).pack(side="left", padx=6)
    listas, var_prefijo, entrada, botones = _listas_y_prefijo(
        w["izq"], (("incluir", "Incluir (prefijos)"), ("excluir", "Excluir (prefijos)")))

    def categoria():
        return etiquetas[w["selector"].get()]

    def refrescar_consulta(*_):
        arbol.delete(*arbol.get_children())
        cat, termino, incluidas = categoria(), w["buscar"].get().strip().lower(), 0
        for codigo, nombre in puc.items():
            if not _nivel_ok(codigo, w["nivel"].get()) or (termino and termino not in codigo and termino not in nombre.lower()):
                continue
            pertenece, prefijo = clasificar(codigo, cat, reglas)
            if not prefijo or (not pertenece and not var_excluidas.get()):
                continue
            incluidas += pertenece
            arbol.insert("", "end", values=(codigo, nombre, f"{'Incluida' if pertenece else 'Excluida'} por {prefijo}"),
                         tags=() if pertenece else ("excluida",))
        w["total"].configure(text=f"{incluidas} cuentas incluidas en {w['selector'].get()} con los filtros actuales.")

    def refrescar_listas(*_):
        for clave, lb in listas.items():
            lb.delete(0, "end")
            for p in sorted(reglas[categoria()][clave]):
                lb.insert("end", f"{p}  ·  {nombre_cuenta(p)}")
        refrescar_consulta()

    def agregar(clave):
        p = var_prefijo.get().strip()
        if not p.isdigit():
            messagebox.showwarning("Categorías", "El prefijo debe tener solo dígitos (ej. 2365).", parent=dlg.ventana)
            return
        r = reglas[categoria()]
        otra = "excluir" if clave == "incluir" else "incluir"
        if p in r[otra]:
            r[otra].remove(p)
        if p not in r[clave]:
            r[clave].append(p)
        var_prefijo.set("")
        refrescar_listas()

    def quitar():
        r = reglas[categoria()]
        for clave, lb in listas.items():
            for i in reversed(lb.curselection()):
                r[clave].remove(lb.get(i).split()[0])
        refrescar_listas()

    ttk.Button(botones, text="Agregar a Incluir", command=lambda: agregar("incluir")).pack(side="left")
    ttk.Button(botones, text="Agregar a Excluir", command=lambda: agregar("excluir")).pack(side="left", padx=6)
    ttk.Button(botones, text="Quitar seleccionado", command=quitar).pack(side="left")

    def solapamientos(limpio):
        excluyentes = [c for c in CATEGORIAS_EXCLUYENTES if c in limpio]
        cruces = []
        for codigo in puc:
            cats = [c for c in excluyentes if clasificar(codigo, c, limpio)[0]]
            if len(cats) > 1:
                cruces.append(f"{codigo} ({', '.join(ETIQUETAS_CATEGORIAS.get(c, c) for c in cats)})")
        return cruces

    def guardar():
        try:
            limpio = parametros_repo.validar_reglas(reglas)
        except ValueError as e:
            messagebox.showwarning("Categorías", str(e), parent=dlg.ventana)
            return
        cruces = solapamientos(limpio)
        if cruces and not messagebox.askyesno(
                "Categorías", f"{len(cruces)} cuentas quedan en más de una categoría excluyente, por ejemplo:\n\n"
                + "\n".join(cruces[:10]) + "\n\n¿Guardar de todas formas?", parent=dlg.ventana):
            return
        try:
            ruta = parametros_repo.guardar_reglas(limpio)
        except OSError as e:
            messagebox.showerror("Categorías", str(e), parent=dlg.ventana)
            return
        recargar_reglas()
        messagebox.showinfo("Categorías", f"Categorías guardadas en:\n{ruta}", parent=dlg.ventana)
        dlg.cerrar(True)

    def restaurar():
        if messagebox.askyesno("Categorías", "¿Restaurar las reglas de fábrica de todas las categorías? "
                               "(no se guarda hasta presionar Guardar)", parent=dlg.ventana):
            reglas.clear()
            reglas.update(parametros_repo.reglas_por_defecto())
            refrescar_listas()

    dlg.boton("Restaurar valores de fábrica", restaurar, lado="left")
    dlg.boton("Cancelar", dlg.cerrar)
    dlg.boton("Guardar", guardar)

    w["combo"].bind("<<ComboboxSelected>>", refrescar_listas)
    for var in (w["buscar"], w["nivel"], var_excluidas):
        var.trace_add("write", refrescar_consulta)
    arbol.bind("<Double-1>", lambda _e: arbol.selection() and var_prefijo.set(arbol.item(arbol.selection()[0], "values")[0]))
    entrada.bind("<Return>", lambda _e: agregar("incluir"))
    refrescar_listas()
    return bool(dlg.mostrar())


# ============================================================== cuentas por concepto de base mínima
def configurar_conceptos_retencion(parent):
    """Consulta y edición de las cuentas que definen el concepto de base mínima (compras, servicios...)."""
    from core.config import ETIQUETAS_CONCEPTOS, INFO_CONCEPTOS
    from data import parametros_repo
    from shared.cuentas import cuentas_puc, nombre_cuenta, pertenece_a_categoria

    conceptos = {k: dict(v) for k, v in parametros_repo.cargar_conceptos().items()}
    etiquetas = {v: k for k, v in ETIQUETAS_CONCEPTOS.items()}
    puc = cuentas_puc()
    universo = [(c, n, "gasto") for c, n in puc.items() if pertenece_a_categoria(c, "base")]
    universo += [(c, n, "retencion") for c, n in puc.items() if pertenece_a_categoria(c, "autorretenedor")]
    tipos = {"gasto": "Gasto/costo", "retencion": "Retención"}

    dlg = Dialogo(parent, "Cuentas por concepto de base mínima", ancho=1100, alto=620, min_ancho=820,
                  min_alto=480)
    w = _panel_consulta(
        dlg, "Concepto", list(etiquetas), ETIQUETAS_CONCEPTOS["compras"], "Conceptos de base mínima", INFO_CONCEPTOS,
        ["Cuentas (4 dígitos)", "Subcuentas (6 dígitos)", "Auxiliares (9+ dígitos)", "Todas"],
        [("codigo", "Código"), ("nombre", "Nombre"), ("tipo", "Tipo"), ("origen", "Asignada por")],
        {"codigo": 1, "nombre": 3.4, "tipo": 1.2, "origen": 1.5},
    )
    arbol = w["arbol"]
    listas, var_prefijo, entrada, botones = _listas_y_prefijo(
        w["izq"], (("gasto", "Cuentas de gasto/costo (prefijos)"), ("retencion", "Cuentas de retención (prefijos)")))
    lbl_defecto = ttk.Label(w["izq"], text="", foreground=PALETTE["text_muted"], wraplength=360)
    lbl_defecto.grid(row=99, column=0, sticky="w", pady=(6, 0))

    def concepto():
        return etiquetas[w["selector"].get()]

    def refrescar_consulta(*_):
        arbol.delete(*arbol.get_children())
        actual, termino, total = concepto(), w["buscar"].get().strip().lower(), 0
        for codigo, nombre, tipo in universo:
            if not _nivel_ok(codigo, w["nivel"].get()) or (termino and termino not in codigo and termino not in nombre.lower()):
                continue
            asignado, prefijo = parametros_repo.concepto_de(codigo, conceptos[tipo], "compras" if tipo == "gasto" else "")
            if asignado != actual:
                continue
            total += 1
            arbol.insert("", "end", values=(codigo, nombre, tipos[tipo], prefijo or "por defecto"))
        w["total"].configure(text=f"{total} cuentas usan la base de {w['selector'].get()} con los filtros actuales.")

    def refrescar_listas(*_):
        actual = concepto()
        for tipo, lb in listas.items():
            lb.delete(0, "end")
            for p in sorted(p for p, c in conceptos[tipo].items() if c == actual):
                lb.insert("end", f"{p}  ·  {nombre_cuenta(p)}")
        lbl_defecto.configure(text="Además, toda cuenta de base sin asignar se trata como compra."
                              if actual == "compras" else "")
        refrescar_consulta()

    def asignar(tipo):
        p = var_prefijo.get().strip()
        if not p.isdigit():
            messagebox.showwarning("Conceptos", "El prefijo debe tener solo dígitos (ej. 5135).", parent=dlg.ventana)
            return
        if tipo == "retencion" and concepto() == "mixto":
            messagebox.showwarning("Conceptos", "Una cuenta de retención no puede ser 'Mixto'.", parent=dlg.ventana)
            return
        previo = conceptos[tipo].get(p)
        if previo and previo != concepto() and not messagebox.askyesno(
                "Conceptos", f"{p} está asignado a {ETIQUETAS_CONCEPTOS[previo]}. ¿Moverlo a {w['selector'].get()}?",
                parent=dlg.ventana):
            return
        conceptos[tipo][p] = concepto()
        var_prefijo.set("")
        refrescar_listas()

    def quitar():
        for tipo, lb in listas.items():
            for i in reversed(lb.curselection()):
                conceptos[tipo].pop(lb.get(i).split()[0], None)
        refrescar_listas()

    ttk.Button(botones, text="Asignar como gasto", command=lambda: asignar("gasto")).pack(side="left")
    ttk.Button(botones, text="Asignar como retención", command=lambda: asignar("retencion")).pack(side="left", padx=6)
    ttk.Button(botones, text="Quitar seleccionado", command=quitar).pack(side="left")

    def guardar():
        try:
            ruta = parametros_repo.guardar_conceptos(parametros_repo.validar_conceptos(conceptos))
        except (ValueError, OSError) as e:
            messagebox.showwarning("Conceptos", str(e), parent=dlg.ventana)
            return
        messagebox.showinfo("Conceptos", f"Conceptos guardados en:\n{ruta}", parent=dlg.ventana)
        dlg.cerrar(True)

    def restaurar():
        if messagebox.askyesno("Conceptos", "¿Restaurar las asignaciones de fábrica? (no se guarda hasta presionar "
                               "Guardar)", parent=dlg.ventana):
            conceptos.clear()
            conceptos.update(parametros_repo.conceptos_por_defecto())
            refrescar_listas()

    dlg.boton("Restaurar valores de fábrica", restaurar, lado="left")
    dlg.boton("Cancelar", dlg.cerrar)
    dlg.boton("Guardar", guardar)

    w["combo"].bind("<<ComboboxSelected>>", refrescar_listas)
    for var in (w["buscar"], w["nivel"]):
        var.trace_add("write", refrescar_consulta)
    arbol.bind("<Double-1>", lambda _e: arbol.selection() and var_prefijo.set(arbol.item(arbol.selection()[0], "values")[0]))
    entrada.bind("<Return>", lambda _e: asignar("gasto"))
    refrescar_listas()
    return bool(dlg.mostrar())

# ============================================================== guía de alertas
def mostrar_guia_alertas(parent):
    """Explica cada alerta tributaria: cuándo aparece y qué significa."""
    from core.config import GUIA_ALERTAS, PERFIL_TRIBUTARIO_ACTIVO

    dlg = Dialogo(parent, "Guía de alertas tributarias", ancho=820, alto=640, min_ancho=560)
    marco = dlg.cuerpo
    ancho_texto = 740
    ttk.Label(marco, text="Las alertas no cambian la conciliación: señalan casos para revisar. Aparecen en la columna "
                          "ALERTA de AUD-COMP y Alertas de DIAN VS CONT; si hay varias se separan con \" / \". "
                          "Diferencias menores a $50 o al 0,1 % del valor se ignoran.",
              wraplength=ancho_texto, justify="left").pack(anchor="w", pady=(0, 10))

    for seccion, nota, alertas in GUIA_ALERTAS:
        if seccion.startswith("Perfil tributario") and PERFIL_TRIBUTARIO_ACTIVO:
            seccion = "Perfil tributario"
        ttk.Label(marco, text=seccion, font=("", 11, "bold"), foreground=PALETTE["primary"]).pack(anchor="w", pady=(12, 2))
        if nota:
            ttk.Label(marco, text=nota, wraplength=ancho_texto, justify="left",
                      foreground=PALETTE["text_muted"]).pack(anchor="w", pady=(0, 4))
        for alerta, cuando, significa in alertas:
            bloque = ttk.Frame(marco)
            bloque.pack(fill="x", pady=(6, 0), padx=(8, 0))
            ttk.Label(bloque, text=f"• {alerta}", font=("", 10, "bold"), wraplength=ancho_texto,
                      justify="left").pack(anchor="w")
            ttk.Label(bloque, text=f"Cuándo aparece: {cuando}", wraplength=ancho_texto - 16,
                      justify="left").pack(anchor="w", padx=(14, 0))
            ttk.Label(bloque, text=f"Qué significa: {significa}", wraplength=ancho_texto - 16,
                      justify="left").pack(anchor="w", padx=(14, 0))

    dlg.boton("Cerrar", dlg.cerrar, principal=True)
    return dlg.mostrar()