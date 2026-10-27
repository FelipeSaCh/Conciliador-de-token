import tkinter as tk
from tkinter import ttk

import pandas as pd

from core.config import PREVIEW_MAX_COLS, PREVIEW_MAX_ROWS
from core.errors import logger


class VistaPreviaExcel(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._file_path = None
        self._cache = {}

        top = ttk.Frame(self)
        top.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(top, text="Hoja", style="Subheader.TLabel").pack(side=tk.LEFT)
        self.combo_hojas = ttk.Combobox(top, state="readonly", width=40)
        self.combo_hojas.pack(side=tk.LEFT, padx=(10, 0))
        self.combo_hojas.bind("<<ComboboxSelected>>", lambda _e: self.combo_hojas.get() and self._mostrar_hoja(self.combo_hojas.get()))
        self.lbl_info = ttk.Label(top, text="", style="Subheader.TLabel")
        self.lbl_info.pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def cargar_archivo(self, file_path, nombres_hojas):
        """`nombres_hojas` ya viene de la app: evita reabrir el libro solo para listar hojas."""
        self._file_path = file_path
        self._cache.clear()
        hojas = list(nombres_hojas or [])
        self.combo_hojas['values'] = hojas
        if hojas:
            self.combo_hojas.current(0)
            self._mostrar_hoja(hojas[0])
        else:
            self._limpiar_tabla()

    def limpiar(self):
        self._file_path = None
        self._cache.clear()
        self.combo_hojas.set("")
        self.combo_hojas['values'] = []
        self.lbl_info.configure(text="")
        self._limpiar_tabla()

    def _mostrar_hoja(self, hoja):
        df = self._cache.get(hoja)
        if df is None:
            try:
                df = pd.read_excel(self._file_path, sheet_name=hoja, nrows=PREVIEW_MAX_ROWS, header=None)
                df = self._cache.setdefault(hoja, df.iloc[:, :PREVIEW_MAX_COLS])
            except Exception as e:
                logger.error(f"No se pudo leer la hoja '{hoja}' para vista previa: {e}")
                self.lbl_info.configure(text=f"Error leyendo la hoja '{hoja}'")
                self._limpiar_tabla()
                return
        self._poblar_tabla(df)
        self.lbl_info.configure(text=f"{len(df)} filas × {len(df.columns)} columnas (vista previa)")

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
        for fila in df.itertuples(index=False, name=None):
            self.tree.insert("", tk.END, values=["" if pd.isna(v) else str(v) for v in fila])