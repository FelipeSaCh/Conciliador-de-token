from tkinter import ttk


class Pagina(ttk.Frame):
    """Página de la app. Recibe la app (contexto compartido) y reacciona a la carga/liberación del archivo."""

    titulo = ""
    descripcion = ""

    def __init__(self, master, app):
        super().__init__(master, style="TFrame")
        self.app = app
        self.ocupada = False
        self.construir()

    def encabezado(self, parent=None):
        destino = parent or self
        ttk.Label(destino, text=self.titulo, style="Header.TLabel").pack(anchor="w", pady=(0, 4 if self.descripcion else 20))
        if self.descripcion:
            ttk.Label(destino, text=self.descripcion, style="PageDesc.TLabel").pack(anchor="w", pady=(0, 20))

    def construir(self):
        raise NotImplementedError

    def on_archivo_cargado(self, hojas):
        pass

    def on_archivo_liberado(self):
        pass


def asignar_hojas(combo, var, hojas, opcional=False, reiniciar=True):
    """Actualiza las opciones del combo y, si la hoja actual ya no existe, la reinicia."""
    combo['values'] = ([""] if opcional else []) + list(hojas)
    if not reiniciar:
        return
    actual = var.get()
    if actual not in hojas and actual != "":
        var.set("" if opcional else (hojas[0] if hojas else ""))