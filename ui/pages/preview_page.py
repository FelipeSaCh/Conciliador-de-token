import tkinter as tk

from ui.pages.base import Pagina
from ui.preview_widget import VistaPreviaExcel


class PreviewPage(Pagina):
    titulo = "Vista previa"
    descripcion = "Consulta el contenido de las hojas sin modificar el archivo."

    def construir(self):
        self.encabezado()
        self.vista = VistaPreviaExcel(self)
        self.vista.pack(fill=tk.BOTH, expand=True)

    def on_archivo_cargado(self, hojas):
        self.vista.cargar_archivo(self.app.file_path.get(), hojas)

    def on_archivo_liberado(self):
        self.vista.limpiar()