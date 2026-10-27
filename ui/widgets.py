import tkinter as tk
from tkinter import ttk

from core.config import CATEGORIAS_EXCLUYENTES, CATEGORY_COLORS, PREFIJOS_CATEGORIAS, SERIALES_EXENTOS
from ui.theme import PALETTE, fuente, fuente_titulo


def _desplazar(canvas, event):
    num = getattr(event, "num", None)
    if num == 4:
        canvas.yview_scroll(-3, "units")
    elif num == 5:
        canvas.yview_scroll(3, "units")
    else:
        canvas.yview_scroll(int(-1 * (event.delta / 120) * 3), "units")


class _CanvasDesplazable(ttk.Frame):
    """Frame con canvas + scrollbar vertical y rueda del mouse acotada al propio widget."""

    def __init__(self, parent, bg, estilo, height=None, **kwargs):
        super().__init__(parent, style=estilo, **kwargs)
        opciones = {"height": height} if height else {}
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0, **opciones)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style=estilo)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.actualizar_scrollregion())
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._window, width=e.width))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind_all(seq, self._on_wheel, add="+")

    def actualizar_scrollregion(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _contenedor_del_evento(self, event):
        w = self.winfo_containing(event.x_root, event.y_root)
        while w is not None and w is not self:
            if self._cede_a(w):
                return None
            w = w.master
        return w

    def _cede_a(self, widget):
        return False

    def _on_wheel(self, event):
        if self._contenedor_del_evento(event) is self:
            _desplazar(self.canvas, event)


class ScrollableTab(_CanvasDesplazable):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, PALETTE["bg"], "TFrame", **kwargs)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

    def _cede_a(self, widget):
        return isinstance(widget, ScrollableChecklist)


class ScrollableChecklist(_CanvasDesplazable):
    def __init__(self, parent, height=300, **kwargs):
        super().__init__(parent, PALETTE["surface"], "Card.TFrame", height=height, **kwargs)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

    def limpiar(self):
        for w in self.inner.winfo_children():
            w.destroy()

    def mensaje(self, texto):
        ttk.Label(self.inner, text=texto, style="SubheaderCard.TLabel").pack(anchor="w", padx=8, pady=12)

    def actualizar_scrollregion(self):
        self.canvas.update_idletasks()
        super().actualizar_scrollregion()


def card(parent, **pack_kwargs):
    outer = tk.Frame(parent, bg=PALETTE["border"])
    inner = ttk.Frame(outer, style="Card.TFrame", padding=20)
    inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
    outer.pack(**pack_kwargs)
    return inner


def encabezado_card(parent, titulo, subtitulo=None, pady_sub=(4, 16)):
    ttk.Label(parent, text=titulo, style="CardHeader.TLabel").pack(anchor="w")
    if subtitulo:
        ttk.Label(parent, text=subtitulo, style="SubheaderCard.TLabel").pack(anchor="w", pady=pady_sub)


def boton_primario(parent, text, command):
    btn = tk.Button(
        parent, text=text, command=command, font=fuente_titulo(10),
        bg=PALETTE["primary"], fg=PALETTE["primary_fg"],
        activebackground=PALETTE["primary_active"], activeforeground=PALETTE["primary_fg"],
        disabledforeground="#F3F4F6", relief="flat", bd=0, highlightthickness=0, padx=18, pady=8, cursor="hand2",
    )
    btn.bind("<Enter>", lambda e: btn.configure(bg=PALETTE["primary_hover"]) if btn["state"] != "disabled" else None)
    btn.bind("<Leave>", lambda e: btn.configure(bg=PALETTE["primary"]) if btn["state"] != "disabled" else None)
    return btn


def boton_secundario(parent, text, command, **kwargs):
    return ttk.Button(parent, text=text, style="Secondary.TButton", command=command, cursor="hand2", **kwargs)


def habilitar_boton(btn, habilitado):
    btn.configure(state="normal" if habilitado else "disabled",
                  bg=PALETTE["primary"] if habilitado else PALETTE["disabled"])


def leyenda_categorias(parent):
    for etiqueta, clave in (("IVA", "iva"), ("Base", "base"), ("Base 2", "base2"), ("Autorretenedor", "autorretenedor")):
        item = tk.Frame(parent, bg=PALETTE["surface"])
        item.pack(side=tk.LEFT, padx=(0, 18))
        tk.Frame(item, bg=CATEGORY_COLORS[clave], width=10, height=10).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(item, text=etiqueta, bg=PALETTE["surface"], fg=PALETTE["text_muted"], font=(fuente(), 9)).pack(side=tk.LEFT)


def chip_toggle(parent, color, var):
    chip = tk.Label(parent, text=" ", width=3, bg=PALETTE["border"], relief="flat", cursor="hand2", font=(fuente(), 8))

    def refrescar(*_):
        chip.configure(bg=color if var.get() else PALETTE["border"])

    chip.bind("<Button-1>", lambda _e: var.set(not var.get()))
    var.trace_add("write", refrescar)
    refrescar()
    return chip


def categoria_inicial(columna, categoria):
    col = str(columna).strip()
    if col.startswith(SERIALES_EXENTOS):
        return False
    prefijos = PREFIJOS_CATEGORIAS.get(categoria, ())
    return col.startswith(prefijos) if prefijos else False


class SerialesChecklist(ttk.Frame):
    """Buscador + lista de seriales (columnas) con chips por categoría (IVA/BASE/BASE 2/Autorretenedor)."""

    def __init__(self, parent, titulo, subtitulo, texto_vacio, height):
        super().__init__(parent, style="Card.TFrame")
        self.texto_vacio = texto_vacio
        self.vars = {}
        self._filas = {}
        self._columnas = []

        encabezado_card(self, titulo, subtitulo, pady_sub=(4, 0))

        fila_busqueda = ttk.Frame(self, style="Card.TFrame")
        fila_busqueda.pack(fill=tk.X, pady=(14, 8))
        ttk.Label(fila_busqueda, text="Buscar cuenta", style="FieldLabelOptional.TLabel").pack(side=tk.LEFT, padx=(0, 12))
        self.busqueda = tk.StringVar()
        self.busqueda.trace_add("write", lambda *_: self._filtrar())
        ttk.Entry(fila_busqueda, textvariable=self.busqueda).pack(side=tk.LEFT, fill=tk.X, expand=True)

        barra = ttk.Frame(self, style="Card.TFrame")
        barra.pack(fill=tk.X, pady=(0, 6))
        leyenda_categorias(barra)

        self.lista = ScrollableChecklist(self, height=height)
        self.lista.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.poblar([])

    def poblar(self, columnas):
        self.lista.limpiar()
        self.vars, self._filas, self._columnas = {}, {}, list(columnas)
        if not columnas:
            self.lista.mensaje(self.texto_vacio)
            self.lista.actualizar_scrollregion()
            return

        for col in columnas:
            fila = tk.Frame(self.lista.inner, bg=PALETTE["surface"])
            contenido = tk.Frame(fila, bg=PALETTE["surface"], padx=10, pady=7)
            contenido.pack(fill=tk.X)
            lbl = tk.Label(contenido, text=col, bg=PALETTE["surface"], fg=PALETTE["text"],
                           font=(fuente(), 10), anchor="w", width=55)
            lbl.pack(side=tk.LEFT, padx=(0, 24))

            vars_col = {}
            for cat, color in CATEGORY_COLORS.items():
                var = tk.BooleanVar(value=categoria_inicial(col, cat))
                chip_toggle(contenido, color, var).pack(side=tk.LEFT, padx=(6, 0))
                vars_col[cat] = var
            self._hacer_excluyentes(vars_col)
            self.vars[col] = vars_col
            self._filas[col] = (fila, contenido, lbl)

        if self.busqueda.get():
            self.busqueda.set("")
        else:
            self._filtrar()

    @staticmethod
    def _hacer_excluyentes(vars_col):
        grupo = [vars_col[c] for c in CATEGORIAS_EXCLUYENTES if c in vars_col]
        for var in grupo:
            def _trace(*_, activa=var):
                if activa.get():
                    for otra in grupo:
                        if otra is not activa:
                            otra.set(False)
            var.trace_add("write", _trace)

    def _filtrar(self):
        termino = self.busqueda.get().strip().lower()
        for fila, _, _ in self._filas.values():
            fila.pack_forget()
        visibles = 0
        for col in self._columnas:
            if col in self._filas and termino in col.lower():
                fila, contenido, lbl = self._filas[col]
                fondo = PALETTE["surface"] if visibles % 2 == 0 else PALETTE["surface_alt"]
                for w in (fila, contenido, lbl):
                    w.configure(bg=fondo)
                fila.pack(fill=tk.X)
                visibles += 1
        self.lista.actualizar_scrollregion()

    def seleccionados(self, categoria):
        return [col for col, v in self.vars.items() if v[categoria].get()]