import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.errors import ErrorSistema, ErrorUsuario
from services.reportes.informe_iva import GeneradorInformeIVA
from ui.dialogs import VentanaVistaPreviaPDF, seleccionar_tipos_documento
from ui.pages.base import Pagina, asignar_hojas
from ui.widgets import boton_primario, boton_secundario, card, encabezado_card, habilitar_boton


class IvaPage(Pagina):
    titulo = "Informe de IVA"
    descripcion = "Genera el informe de IVA desde la hoja de token y expórtalo a PDF."

    def construir(self):
        self.encabezado()
        self.generador = None
        self._exportando = False

        c = card(self, fill=tk.X, pady=(0, 20))
        encabezado_card(c, "Configuración del informe", "Selecciona la hoja de token principal.")
        fila = ttk.Frame(c, style="Card.TFrame")
        fila.pack(fill=tk.X)
        ttk.Label(fila, text="Hoja de token", style="FieldLabel.TLabel").pack(side=tk.LEFT, padx=(0, 30))
        self.hoja_var = tk.StringVar()
        self.combo = ttk.Combobox(fila, textvariable=self.hoja_var, width=50, state="readonly", cursor="hand2")
        self.combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        acciones = ttk.Frame(c, style="Card.TFrame")
        acciones.pack(fill=tk.X, pady=(24, 0))
        self.btn_generar = boton_primario(acciones, "Generar informe", self._on_generar)
        self.btn_generar.pack(side=tk.LEFT)
        self.lbl_estado = ttk.Label(acciones, text="", style="SubheaderCard.TLabel")
        self.lbl_estado.pack(side=tk.LEFT, padx=(20, 0))

        pdf = ttk.Frame(c, style="Card.TFrame")
        pdf.pack(fill=tk.X, pady=(16, 0))
        self.btn_pdf = boton_secundario(pdf, "Exportar a PDF", self._on_exportar_pdf, state="disabled")
        self.btn_pdf.pack(side=tk.LEFT)
        self.btn_vista = boton_secundario(pdf, "Vista previa PDF", self._on_vista_previa, state="disabled")
        self.btn_vista.pack(side=tk.LEFT, padx=(8, 0))

    def _botones_pdf(self, estado):
        self.btn_pdf.configure(state=estado)
        self.btn_vista.configure(state=estado)

    def on_archivo_cargado(self, hojas):
        asignar_hojas(self.combo, self.hoja_var, hojas)
        if self.hoja_var.get() not in hojas:
            self.hoja_var.set(hojas[0] if hojas else "")

    def on_archivo_liberado(self):
        self.combo['values'] = []
        self.hoja_var.set("")
        self.lbl_estado.configure(text="")
        self.generador = None
        self._botones_pdf("disabled")

    def _on_generar(self):
        if self.ocupada or not self.app.requiere_archivo():
            return
        hoja = self.hoja_var.get().strip()
        if not hoja:
            messagebox.showwarning("Hoja requerida", "Debes seleccionar la hoja de token principal.")
            return
        try:
            tipos, preseleccionados = GeneradorInformeIVA.obtener_tipos_documento(self.app.file_path.get(), hoja)
        except ErrorUsuario as e:
            messagebox.showwarning("Tipos de documento", str(e))
            return
        except ErrorSistema as e:
            messagebox.showerror("Error", str(e))
            return

        seleccion = seleccionar_tipos_documento(self.app, tipos, preseleccionados)
        if not seleccion:
            return

        self.ocupada = True
        habilitar_boton(self.btn_generar, False)
        self._botones_pdf("disabled")
        self.lbl_estado.configure(text="Generando informe...")
        self.generador = GeneradorInformeIVA(
            self.app.file_path.get(), hoja, tipos_documento=seleccion,
            progress_callback=self.app.ejecutor.en_ui(lambda m: self.lbl_estado.configure(text=m)),
        )
        self.app.ejecutor.lanzar(self.generador.generar, self._on_generado, self._on_error, "informe de IVA")

    def _fin_generacion(self):
        self.ocupada = False
        habilitar_boton(self.btn_generar, True)

    def _on_generado(self, resumen):
        self._fin_generacion()
        self.lbl_estado.configure(text=f"IVA a pagar: {resumen['totales']['iva_a_pagar']:,.2f}")
        self._botones_pdf("normal")
        messagebox.showinfo("Informe generado", "Informe generado en la hoja 'Informe IVA'.")

    def _on_error(self, mensaje, es_sistema):
        self._fin_generacion()
        self._exportando = False
        self.lbl_estado.configure(text=f"Error: {mensaje}")
        (messagebox.showerror if es_sistema else messagebox.showwarning)("Error", mensaje)

    def _informe_listo(self):
        if self.generador and self.generador.df_filtrado is not None:
            return True
        messagebox.showwarning("Informe requerido", "Primero debes generar el informe de IVA.")
        return False

    def _on_exportar_pdf(self):
        if self._exportando or not self._informe_listo():
            return
        ruta = filedialog.asksaveasfilename(
            title="Guardar informe de IVA en PDF", defaultextension=".pdf",
            initialfile=self.generador.nombre_sugerido_pdf(), filetypes=[("Archivo PDF", "*.pdf")],
        )
        if not ruta:
            return
        self._exportando = True
        self.btn_pdf.configure(state="disabled")
        self.lbl_estado.configure(text="Generando PDF...")
        self.app.ejecutor.lanzar(lambda: self.generador.generar_pdf(ruta), self._on_pdf_exportado,
                                 self._on_error_pdf, "exportación PDF de IVA")

    def _on_pdf_exportado(self, ruta):
        self._exportando = False
        self.btn_pdf.configure(state="normal")
        self.lbl_estado.configure(text=f"PDF exportado: {Path(ruta).name}")
        messagebox.showinfo("PDF", f"Informe exportado en:\n{ruta}")

    def _on_error_pdf(self, mensaje, es_sistema):
        self._exportando = False
        self.btn_pdf.configure(state="normal")
        self.lbl_estado.configure(text=f"Error: {mensaje}")
        (messagebox.showerror if es_sistema else messagebox.showwarning)("Error", mensaje)

    def _on_vista_previa(self):
        if not self._informe_listo():
            return
        try:
            import pymupdf  # noqa: F401
        except ImportError:
            messagebox.showerror("Dependencia faltante", "No se encontró PyMuPDF. Instálalo con: pip install pymupdf")
            return
        ruta = Path(tempfile.gettempdir()) / "vista_previa_informe_iva.pdf"
        try:
            self.generador.generar_pdf(str(ruta))
        except (ErrorUsuario, ErrorSistema) as e:
            messagebox.showerror("Error al generar vista previa", str(e))
            return
        VentanaVistaPreviaPDF(self.app, str(ruta))