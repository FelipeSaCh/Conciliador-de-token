import tkinter as tk
from tkinter import ttk

from core.config import (
    COLUMNAS_EXCLUIDAS_AUD_COMP, DEFAULT_SHEET_NAMES, INFO_PARAMETROS_UVT, INFO_PERIODICIDAD_IVA, INFO_TIPOS_CONTRIBUYENTE,
    JERARQUIA_SERIALES, OPTIONAL_SHEETS, PERFIL_TRIBUTARIO_ACTIVO, PERIODICIDADES_IVA, SHEET_LABELS, SHEET_ORDER,
    TIPOS_CONTRIBUYENTE,
)
from core.errors import logger
from data import parametros_repo, perfiles_repo
from ui.dialogs import configurar_categorias_cuentas, configurar_parametros_tributarios
from shared.cuentas import ordenar_por_prefijos
from shared.excel_io import cargar_hoja_encabezado_variable
from ui.pages.base import Pagina, asignar_hojas
from ui.theme import PALETTE, fuente
from ui.widgets import (
    SerialesChecklist, boton_primario, boton_secundario, card, chip_info, encabezado_card,
)

FILTROS_GRUPO = ["Solo Recibidos", "Solo Emitidos", "Ambos"]


class AuditoriaPage(Pagina):
    titulo = "Configurar auditoría"
    descripcion = "Relaciona las hojas del archivo y clasifica los seriales de IVA, base y autorretención."

    def construir(self):
        wrapper = ttk.Frame(self, style="TFrame")  # el desplazamiento lo da el contenedor de página (app)
        wrapper.pack(fill=tk.BOTH, expand=True)
        self.encabezado(wrapper)

        paneles = ttk.Frame(wrapper, style="TFrame")
        paneles.pack(fill=tk.BOTH, expand=True)
        col_izq = ttk.Frame(paneles, style="TFrame")
        col_izq.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 16))
        col_der = ttk.Frame(paneles, style="TFrame")
        col_der.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._construir_hojas(col_izq)
        self._construir_filtro(col_izq)
        self._construir_parametros(col_izq)
        if PERFIL_TRIBUTARIO_ACTIVO:
            self._construir_perfil(col_izq)

        self.seriales = {
            "aud_comp": SerialesChecklist(
                card(col_der, fill=tk.BOTH, expand=True), "Seriales de compras",
                "Marca a qué categoría pertenece cada columna de la hoja AUD-COMP.",
                "Selecciona un archivo y la hoja AUD-COMP.", 240),
            "Aud_dc": SerialesChecklist(
                card(col_der, fill=tk.BOTH, expand=True, pady=(16, 0)), "Seriales de devoluciones",
                "Marca a qué categoría pertenece cada columna de la hoja de auditoría de devoluciones.",
                "Selecciona un archivo y la hoja de Devoluciones.", 180),
        }
        for checklist in self.seriales.values():
            checklist.pack(fill=tk.BOTH, expand=True)

    def _construir_hojas(self, parent):
        c = card(parent, fill=tk.X)
        encabezado_card(c, "Estructura de hojas", "Relaciona cada sección requerida con la pestaña de tu archivo.")
        grid = ttk.Frame(c, style="Card.TFrame")
        grid.pack(fill=tk.X, anchor="n")

        self.combos = {}
        for i, clave in enumerate(SHEET_ORDER):
            opcional = clave in OPTIONAL_SHEETS
            ttk.Label(grid, text=SHEET_LABELS[clave] + ("" if opcional else "  *"),
                      style="FieldLabelOptional.TLabel" if opcional else "FieldLabel.TLabel").grid(
                row=i, column=0, sticky="w", pady=8, padx=(0, 20))
            combo = ttk.Combobox(grid, textvariable=self.app.sheet_vars[clave], width=35, state="readonly", cursor="hand2")
            combo.grid(row=i, column=1, sticky="ew", pady=8)
            self.combos[clave] = combo
            if clave in ("aud_comp", "Aud_dc"):
                combo.bind("<<ComboboxSelected>>", lambda _e, k=clave: self.refrescar_seriales(k))
            if opcional:
                self._boton_limpiar(grid, clave).grid(row=i, column=2, padx=(8, 0))
        grid.columnconfigure(1, weight=1)
        ttk.Label(c, text="* Campo obligatorio", style="SubheaderCard.TLabel").pack(anchor="w", pady=(12, 0))

    def _boton_limpiar(self, parent, clave):
        btn = tk.Button(parent, text="✕", command=lambda: self.app.sheet_vars[clave].set(""),
                        bg=PALETTE["surface"], fg=PALETTE["text_muted"], activebackground=PALETTE["surface"],
                        font=(fuente(), 9), relief="flat", bd=0, highlightthickness=0, cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.configure(fg=PALETTE["danger"]))
        btn.bind("<Leave>", lambda e: btn.configure(fg=PALETTE["text_muted"]))
        return btn

    def _construir_filtro(self, parent):
        c = card(parent, fill=tk.X, pady=(16, 0))
        encabezado_card(c, "Filtro de documentos", "Selecciona qué documentos del token procesar.", pady_sub=(4, 12))
        fila = ttk.Frame(c, style="Card.TFrame")
        fila.pack(fill=tk.X)
        ttk.Label(fila, text="Grupo a procesar  *", style="FieldLabel.TLabel").pack(side=tk.LEFT, padx=(0, 20))
        self.filtro_grupo = tk.StringVar(value="")
        ttk.Combobox(fila, textvariable=self.filtro_grupo, values=FILTROS_GRUPO, state="readonly",
                     width=25, cursor="hand2").pack(side=tk.LEFT, pady=8)

    def _construir_parametros(self, parent):
        c = card(parent, fill=tk.X, pady=(16, 0))
        encabezado_card(c, "Parámetros tributarios",
                        "UVT, bases mínimas de retención y categorías de cuentas de los seriales.", pady_sub=(4, 12))
        fila = ttk.Frame(c, style="Card.TFrame")
        fila.pack(fill=tk.X)
        self.lbl_uvt = ttk.Label(fila, style="FieldLabel.TLabel")
        self.lbl_uvt.pack(side=tk.LEFT)
        chip_info(fila, "Parámetros tributarios", INFO_PARAMETROS_UVT).pack(side=tk.LEFT, padx=(6, 0))
        boton_primario(c, "⚙  Configurar UVT y bases mínimas", self._abrir_parametros).pack(fill=tk.X, pady=(12, 0))
        boton_secundario(c, "Consultar y configurar categorías de cuentas", self._abrir_categorias).pack(
            fill=tk.X, pady=(8, 0))
        self._refrescar_resumen_uvt()

    def _refrescar_resumen_uvt(self):
        datos = parametros_repo.cargar()
        anio = max(datos["uvt"])
        valor = f"{datos['uvt'][anio]:,.0f}".replace(",", ".")
        desde, bases = datos["bases"][-1]
        self.lbl_uvt.configure(text=f"UVT {anio}: ${valor}  ·  Vigencia desde {desde}: "
                                    f"compras {bases['compras']:g} / servicios {bases['servicios']:g} UVT")

    def _abrir_parametros(self):
        if configurar_parametros_tributarios(self.winfo_toplevel()):
            self._refrescar_resumen_uvt()

    def _abrir_categorias(self):
        if configurar_categorias_cuentas(self.winfo_toplevel()):
            for clave in self.seriales:
                self.refrescar_seriales(clave)

    def _construir_perfil(self, parent):
        c = card(parent, fill=tk.X, pady=(16, 0))
        encabezado_card(c, "Perfil tributario de la empresa",
                        "Se guarda por NIT y se recupera al digitarlo de nuevo.", pady_sub=(4, 12))
        grid = ttk.Frame(c, style="Card.TFrame")
        grid.pack(fill=tk.X)
        self.perfil_nit = tk.StringVar()
        self.perfil_tipo = tk.StringVar()
        self.perfil_periodicidad = tk.StringVar()

        ttk.Label(grid, text="NIT empresa", style="FieldLabelOptional.TLabel").grid(row=0, column=0, sticky="w", pady=8)
        entrada = ttk.Entry(grid, textvariable=self.perfil_nit, width=20)
        entrada.grid(row=0, column=1, sticky="w", pady=8)
        entrada.bind("<FocusOut>", lambda _e: self._cargar_perfil())

        filas = (
            ("Tipo de contribuyente  *", self.perfil_tipo, list(TIPOS_CONTRIBUYENTE), INFO_TIPOS_CONTRIBUYENTE),
            ("Periodicidad IVA  *", self.perfil_periodicidad, list(PERIODICIDADES_IVA), INFO_PERIODICIDAD_IVA),
        )
        for i, (texto, var, valores, info) in enumerate(filas, start=1):
            ttk.Label(grid, text=texto, style="FieldLabel.TLabel").grid(row=i, column=0, sticky="w", pady=8, padx=(0, 20))
            ttk.Combobox(grid, textvariable=var, values=valores, state="readonly", width=40,
                         cursor="hand2").grid(row=i, column=1, sticky="ew", pady=8)
            chip_info(grid, texto.rstrip(" *"), info).grid(row=i, column=2, padx=(8, 0))
        grid.columnconfigure(1, weight=1)

    def _cargar_perfil(self):
        perfil = perfiles_repo.obtener(self.perfil_nit.get())
        if perfil:
            self.perfil_tipo.set(perfil.get("tipo", ""))
            self.perfil_periodicidad.set(perfil.get("periodicidad", ""))

    def obtener_perfil(self):
        if not PERFIL_TRIBUTARIO_ACTIVO:
            return {}
        return {"nit": self.perfil_nit.get().strip(), "tipo": self.perfil_tipo.get(),
                "periodicidad": self.perfil_periodicidad.get()}

    def refrescar_seriales(self, clave):
        ruta = self.app.file_path.get()
        hoja = self.app.sheet_vars[clave].get().strip()
        checklist = self.seriales[clave]
        if not ruta or not hoja or hoja not in self.app.hojas_disponibles:
            checklist.poblar([])
            return
        try:
            df = cargar_hoja_encabezado_variable(ruta, hoja)
        except Exception as e:
            logger.warning(f"No se pudieron leer las columnas de '{hoja}': {e}")
            checklist.poblar([])
            return

        columnas = [
            str(c).strip() for c in df.columns
            if str(c).strip() and not str(c).strip().startswith("Unnamed")
            and str(c).strip() not in COLUMNAS_EXCLUIDAS_AUD_COMP and not df[c].dropna().empty
        ]
        checklist.poblar(ordenar_por_prefijos(columnas, JERARQUIA_SERIALES))

    def obtener_seriales(self):
        """{'iva', 'base', 'base2', 'iva_dc', 'base_dc', 'base2_dc', 'ret', 'reteiva'}: seriales por categoría."""
        res = {}
        for clave, sufijo in (("aud_comp", ""), ("Aud_dc", "_dc")):
            for cat in ("iva", "base", "base2"):
                res[f"{cat}{sufijo}"] = self.seriales[clave].seleccionados(cat)
        res["ret"] = self.seriales["aud_comp"].seleccionados("autorretenedor")
        res["reteiva"] = self.seriales["aud_comp"].seleccionados("reteiva") if PERFIL_TRIBUTARIO_ACTIVO else []
        return res

    def on_archivo_cargado(self, hojas):
        for clave, combo in self.combos.items():
            asignar_hojas(combo, self.app.sheet_vars[clave], hojas, clave in OPTIONAL_SHEETS)

    def on_archivo_liberado(self):
        for clave, combo in self.combos.items():
            combo['values'] = []
            self.app.sheet_vars[clave].set("" if clave in OPTIONAL_SHEETS else DEFAULT_SHEET_NAMES[clave])
        for checklist in self.seriales.values():
            checklist.poblar([])