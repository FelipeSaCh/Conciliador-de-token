import tkinter as tk
from tkinter import messagebox, ttk

from core.config import OPTIONAL_SHEETS, SHEET_LABELS, TOKEN_SHEETS
from services.token_service import FormateadorToken
from ui.pages.base import Pagina, asignar_hojas
from ui.widgets import boton_primario, card, encabezado_card, habilitar_boton


class TokenPage(Pagina):
    titulo = "Formatear token"
    descripcion = "Prepara la hoja de token (documentos electrónicos) para la conciliación."

    def construir(self):
        self.encabezado()
        c = card(self, fill=tk.X, pady=(0, 20))
        encabezado_card(c, "Hojas de origen", "Selecciona la hoja de Token y, si aplica, Contabilidad y Terceros.")

        grid = ttk.Frame(c, style="Card.TFrame")
        grid.pack(fill=tk.X, anchor="n")
        self.combos = {}
        for i, clave in enumerate(TOKEN_SHEETS):
            opcional = clave in OPTIONAL_SHEETS
            ttk.Label(grid, text=SHEET_LABELS[clave] + ("" if opcional else "  *"),
                      style="FieldLabelOptional.TLabel" if opcional else "FieldLabel.TLabel").grid(
                row=i, column=0, sticky="w", pady=10, padx=(0, 30))
            combo = ttk.Combobox(grid, textvariable=self.app.sheet_vars[clave], width=50, state="readonly", cursor="hand2")
            combo.grid(row=i, column=1, sticky="ew", pady=10)
            self.combos[clave] = combo
        grid.columnconfigure(1, weight=1)

        fila = ttk.Frame(c, style="Card.TFrame")
        fila.pack(fill=tk.X, pady=(24, 0))
        self.btn = boton_primario(fila, "Formatear token", self._on_formatear)
        self.btn.pack(side=tk.LEFT)
        self.lbl_estado = ttk.Label(fila, text="", style="SubheaderCard.TLabel")
        self.lbl_estado.pack(side=tk.LEFT, padx=(20, 0))

    def on_archivo_cargado(self, hojas):
        for clave, combo in self.combos.items():
            asignar_hojas(combo, self.app.sheet_vars[clave], hojas, clave in OPTIONAL_SHEETS, reiniciar=False)

    def on_archivo_liberado(self):
        for combo in self.combos.values():
            combo['values'] = []
        self.lbl_estado.configure(text="")

    def _on_formatear(self):
        if self.ocupada or not self.app.requiere_archivo():
            return
        principal = self.app.sheet_vars["principal"].get().strip()
        if not principal:
            messagebox.showwarning("Hoja requerida", "Debes indicar la hoja de Token.")
            return
        if principal not in self.app.hojas_disponibles:
            messagebox.showwarning("Hoja inválida", f"La hoja '{principal}' no existe en el archivo.")
            return
        if not messagebox.askyesno("Confirmar formateo",
                                   "Este proceso modificará la hoja de Token en el archivo seleccionado. ¿Deseas continuar?"):
            return

        self._set_ocupada(True)
        self.lbl_estado.configure(text="Formateando...")
        hojas = {clave: self.app.sheet_vars[clave].get().strip() for clave in TOKEN_SHEETS}
        formateador = FormateadorToken(
            self.app.file_path.get(), hojas,
            progress_callback=self.app.ejecutor.en_ui(lambda m: self.lbl_estado.configure(text=m)),
        )
        self.app.ejecutor.lanzar(formateador.ejecutar, self._on_exito, self._on_error, "formateo de token")

    def _set_ocupada(self, valor):
        self.ocupada = valor
        habilitar_boton(self.btn, not valor)

    def _on_exito(self, resumen):
        self._set_ocupada(False)
        self.lbl_estado.configure(text=f"Filas: {resumen['filas_procesadas']} | Personales: {resumen['filas_personales']}")
        self.app.recargar_vista_previa()
        messagebox.showinfo("Token formateado", "Token formateado correctamente.")

    def _on_error(self, mensaje, es_sistema):
        self._set_ocupada(False)
        self.lbl_estado.configure(text=f"Error: {mensaje}")
        (messagebox.showerror if es_sistema else messagebox.showwarning)("Error", mensaje)