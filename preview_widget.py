import tkinter as tk
from tkinter import ttk

import pandas as pd

from config import PREVIEW_MAX_COLS, PREVIEW_MAX_ROWS
from errors import logger


class VistaPreviaExcel(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._file_path = None
        self._cache = {}
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=(8, 4))

        ttk.Label(top, text="Hoja:").pack(side=tk.LEFT)
        self.combo_hojas = ttk.Combobox(top, state="readonly", width=40)
        self.combo_hojas.pack(side=tk.LEFT, padx=(6, 0))
        self.combo_hojas.bind("<<ComboboxSelected>>", self._on_hoja_seleccionada)

        self.lbl_info = ttk.Label(top, text="", foreground="#555")
        self.lbl_info.pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def cargar_archivo(self, file_path, nombres_hojas=None):
        self._file_path = file_path
        self._cache.clear()
        try:
            xls = pd.ExcelFile(file_path)
        except Exception as e:
            logger.error(f"No se pudo abrir el archivo para vista previa: {e}")
            self.lbl_info.configure(text="No se pudo leer el archivo")
            return

        hojas = nombres_hojas if nombres_hojas else xls.sheet_names
        self.combo_hojas['values'] = list(hojas)
        if hojas:
            self.combo_hojas.current(0)
            self._mostrar_hoja(hojas[0])
        else:
            self._limpiar_tabla()

    def _on_hoja_seleccionada(self, _event=None):
        hoja = self.combo_hojas.get()
        if hoja:
            self._mostrar_hoja(hoja)

    def _mostrar_hoja(self, hoja):
        if hoja in self._cache:
            df = self._cache[hoja]
        else:
            try:
                df = pd.read_excel(
                    self._file_path, sheet_name=hoja,
                    nrows=PREVIEW_MAX_ROWS, header=None
                )
                df = df.iloc[:, :PREVIEW_MAX_COLS]
                self._cache[hoja] = df
            except Exception as e:
                logger.error(f"No se pudo leer la hoja '{hoja}' para vista previa: {e}")
                self.lbl_info.configure(text=f"Error leyendo la hoja '{hoja}'")
                self._limpiar_tabla()
                return

        self._poblar_tabla(df)
        self.lbl_info.configure(
            text=f"Mostrando {len(df)} filas x {len(df.columns)} columnas (vista previa)"
        )

    def _limpiar_tabla(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []

    def _poblar_tabla(self, df):
        self._limpiar_tabla()
        columnas = [f"Col {i + 1}" for i in range(len(df.columns))]
        self.tree["columns"] = columnas
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor="w", stretch=False)

        for _, row in df.iterrows():
            valores = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            self.tree.insert("", tk.END, values=valores)
