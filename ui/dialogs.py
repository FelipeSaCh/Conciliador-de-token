import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from PIL import Image, ImageTk

from ui.theme import PALETTE, aplicar_icono
from ui.widgets import boton_secundario


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


def seleccionar_tipos_documento(parent, tipos, preseleccionados=None):
    preseleccionados = set(preseleccionados or []) & set(tipos) or set(tipos)
    resultado = {"valor": None}

    ventana = tk.Toplevel(parent)
    ventana.title("Tipos de documento")
    ventana.resizable(False, False)
    ventana.transient(parent)
    ventana.configure(bg=PALETTE["bg"])

    marco = ttk.Frame(ventana, padding=16)
    marco.pack(fill="both", expand=True)
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
            messagebox.showwarning("Tipos de documento", "Debe seleccionar al menos un tipo de documento.", parent=ventana)
            return
        resultado["valor"] = seleccion
        ventana.destroy()

    barra = ttk.Frame(marco)
    barra.pack(fill="x", pady=(12, 0))
    ttk.Button(barra, text="Cancelar", command=ventana.destroy).pack(side="right")
    ttk.Button(barra, text="Aceptar", command=aceptar).pack(side="right", padx=(0, 6))

    ventana.protocol("WM_DELETE_WINDOW", ventana.destroy)
    ventana.bind("<Return>", lambda _e: aceptar())
    ventana.bind("<Escape>", lambda _e: ventana.destroy())

    centrar(ventana, parent)
    ventana.grab_set()
    ventana.focus_set()
    parent.wait_window(ventana)
    return resultado["valor"]


class VentanaVistaPreviaPDF(tk.Toplevel):
    ZOOM = 1.6

    def __init__(self, parent, ruta_pdf):
        super().__init__(parent)
        import pymupdf as fitz

        self.title("Vista previa - Informe de IVA")
        self.geometry("900x700")
        self.configure(bg=PALETTE["bg"])
        aplicar_icono(self)

        self._fitz = fitz
        self.ruta_pdf = ruta_pdf
        self.documento = fitz.open(ruta_pdf)
        self.pagina_actual = 0
        self._imagen_tk = None

        barra = ttk.Frame(self, style="TFrame", padding=14)
        barra.pack(fill=tk.X)
        self.btn_anterior = boton_secundario(barra, "‹ Anterior", lambda: self._ir(-1))
        self.btn_anterior.pack(side=tk.LEFT)
        self.lbl_pagina = ttk.Label(barra, text="", style="Subheader.TLabel")
        self.lbl_pagina.pack(side=tk.LEFT, padx=16)
        self.btn_siguiente = boton_secundario(barra, "Siguiente ›", lambda: self._ir(1))
        self.btn_siguiente.pack(side=tk.LEFT)
        boton_secundario(barra, "Abrir con visor predeterminado", self._abrir_externo).pack(side=tk.RIGHT)

        contenedor = tk.Frame(self, bg=PALETTE["surface_alt"])
        contenedor.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(contenedor, bg=PALETTE["surface_alt"], highlightthickness=0)
        sy = ttk.Scrollbar(contenedor, orient="vertical", command=self.canvas.yview)
        sx = ttk.Scrollbar(contenedor, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120) * 3), "units"))

        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self._mostrar_pagina()

    def _mostrar_pagina(self):
        pix = self.documento[self.pagina_actual].get_pixmap(matrix=self._fitz.Matrix(self.ZOOM, self.ZOOM))
        imagen = Image.frombytes("RGBA" if pix.alpha else "RGB", [pix.width, pix.height], pix.samples)
        self._imagen_tk = ImageTk.PhotoImage(imagen)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self._imagen_tk)
        self.canvas.configure(scrollregion=(0, 0, pix.width, pix.height))

        total = len(self.documento)
        self.lbl_pagina.configure(text=f"Página {self.pagina_actual + 1} de {total}")
        self.btn_anterior.configure(state="normal" if self.pagina_actual > 0 else "disabled")
        self.btn_siguiente.configure(state="normal" if self.pagina_actual < total - 1 else "disabled")

    def _ir(self, delta):
        nueva = self.pagina_actual + delta
        if 0 <= nueva < len(self.documento):
            self.pagina_actual = nueva
            self._mostrar_pagina()

    def _abrir_externo(self):
        try:
            abrir_con_sistema(self.ruta_pdf)
        except Exception:
            messagebox.showinfo("Abrir PDF", f"El archivo se encuentra en:\n{self.ruta_pdf}")

    def _cerrar(self):
        self.documento.close()
        self.destroy()