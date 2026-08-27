# informes_iva.py
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from errors import ErrorSistema, ErrorUsuario, logger

TIPOS_DOC_VALIDOS = {"Factura electrónica", "Nota credito"}
GRUPOS_VALIDOS = {"Emitido", "Recibido"}

COLUMNA_GRUPO = "Grupo"
COLUMNA_TIPO_DOC = "Tipo de documento"
COLUMNA_IVA = "IVA"

HOJA_INFORME_IVA = "Informe IVA"


class GeneradorInformeIVA:
    def __init__(self, ruta_archivo, hoja_token, progress_callback=None):
        self.ruta_archivo = ruta_archivo
        self.hoja_token = hoja_token
        self.progress_callback = progress_callback or (lambda msg: None)

    def _log(self, mensaje):
        logger.info(mensaje)
        self.progress_callback(mensaje)

    def generar(self):
        self._log(f"Leyendo hoja de token: {self.hoja_token}")
        df = self._cargar_hoja(self.ruta_archivo, self.hoja_token)

        for col in (COLUMNA_GRUPO, COLUMNA_TIPO_DOC, COLUMNA_IVA):
            if col not in df.columns:
                raise ErrorUsuario(
                    f"La hoja '{self.hoja_token}' no contiene la columna requerida '{col}'."
                )

        df[COLUMNA_GRUPO] = df[COLUMNA_GRUPO].astype(str).str.strip()
        df[COLUMNA_TIPO_DOC] = df[COLUMNA_TIPO_DOC].astype(str).str.strip()
        df[COLUMNA_IVA] = pd.to_numeric(df[COLUMNA_IVA], errors="coerce").fillna(0)

        df_filtrado = df[
            df[COLUMNA_TIPO_DOC].isin(TIPOS_DOC_VALIDOS)
            & df[COLUMNA_GRUPO].isin(GRUPOS_VALIDOS)
        ].copy()

        if df_filtrado.empty:
            raise ErrorUsuario(
                "No se encontraron registros que cumplan con los criterios de Grupo "
                "y Tipo de documento indicados."
            )

        self._log(f"Registros filtrados: {len(df_filtrado)}")

        df_filtrado["Signo"] = df_filtrado.apply(self._calcular_signo, axis=1)
        df_filtrado["IVA Ajustado"] = df_filtrado[COLUMNA_IVA] * df_filtrado["Signo"]

        resumen = self._construir_resumen(df_filtrado)

        self._log("Escribiendo informe en el archivo de Excel...")
        self._escribir_informe(df_filtrado, resumen)

        self._log("Informe de IVA generado correctamente.")
        return resumen

    @staticmethod
    def _calcular_signo(fila):
        grupo = fila[COLUMNA_GRUPO]
        tipo_doc = fila[COLUMNA_TIPO_DOC]

        if grupo == "Emitido":
            return 1 if tipo_doc == "Factura electrónica" else -1
        if grupo == "Recibido":
            return -1 if tipo_doc == "Factura electrónica" else 1
        return 0

    @staticmethod
    def _cargar_hoja(ruta, hoja):
        try:
            return pd.read_excel(ruta, sheet_name=hoja)
        except Exception as e:
            logger.error(f"No se pudo leer la hoja '{hoja}': {e}")
            raise ErrorSistema(f"No se pudo leer la hoja '{hoja}' del archivo: {e}")

    def _construir_resumen(self, df):
        resumen = {"grupos": {}, "totales": {}}

        for grupo in ("Emitido", "Recibido"):
            df_grupo = df[df[COLUMNA_GRUPO] == grupo]
            facturas = df_grupo[df_grupo[COLUMNA_TIPO_DOC] == "Factura electrónica"]
            notas = df_grupo[df_grupo[COLUMNA_TIPO_DOC] == "Nota credito"]

            resumen["grupos"][grupo] = {
                "facturas_cantidad": len(facturas),
                "facturas_iva": round(facturas[COLUMNA_IVA].sum(), 2),
                "notas_cantidad": len(notas),
                "notas_iva": round(notas[COLUMNA_IVA].sum(), 2),
                "iva_neto": round(df_grupo["IVA Ajustado"].sum(), 2),
            }

        iva_emitido = resumen["grupos"]["Emitido"]["iva_neto"]
        iva_recibido = resumen["grupos"]["Recibido"]["iva_neto"]

        resumen["totales"] = {
            "iva_emitido": iva_emitido,
            "iva_recibido": iva_recibido,
            "iva_a_pagar": round(iva_emitido + iva_recibido, 2),
            "total_registros": len(df),
        }
        return resumen

    def _escribir_informe(self, df, resumen):
        try:
            libro = load_workbook(self.ruta_archivo)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo para escribir el informe: {e}")

        if HOJA_INFORME_IVA in libro.sheetnames:
            del libro[HOJA_INFORME_IVA]
        hoja = libro.create_sheet(HOJA_INFORME_IVA, 0)

        estilo_titulo = Font(size=14, bold=True, color="1F2937")
        estilo_encabezado_seccion = Font(size=11, bold=True, color="FFFFFF")
        relleno_seccion = PatternFill("solid", fgColor="2563EB")
        estilo_encabezado_tabla = Font(bold=True, color="FFFFFF")
        relleno_encabezado_tabla = PatternFill("solid", fgColor="374151")
        relleno_total = PatternFill("solid", fgColor="D1FAE5")
        estilo_total = Font(bold=True)
        borde_fino = Border(*(Side(style="thin", color="D3D8E0"),) * 4)
        formato_moneda = '#,##0.00'
        relleno_ventas=PatternFill("solid", fgColor="edffc3")
        relleno_compras=PatternFill("solid", fgColor="E0E7FF")

        fila = 1
        hoja.cell(row=fila, column=1, value="INFORME DE IVA").font = estilo_titulo
        fila += 1
        hoja.cell(row=fila, column=1, value=f"Hoja origen: {self.hoja_token}").font = Font(italic=True, color="5B6472")
        fila += 2

        hoja.cell(row=fila, column=1, value="RESUMEN GENERAL")
        hoja.cell(row=fila, column=1).font = estilo_encabezado_seccion
        hoja.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=4)
        for c in range(1, 5):
            hoja.cell(row=fila, column=c).fill = relleno_seccion
        fila += 1

        encabezados_resumen = ["Concepto", "Facturas (cant. / IVA)", "Notas Crédito (cant. / IVA)", "IVA Neto"]
        for c, texto in enumerate(encabezados_resumen, start=1):
            celda = hoja.cell(row=fila, column=c, value=texto)
            celda.font = estilo_encabezado_tabla
            celda.fill = relleno_encabezado_tabla
            celda.border = borde_fino
        fila += 1

        nombres_conceptos = {
            "Emitido":"IVA DEVOLUCION DE VENTAS",
            "Recibido":"IVA DEVOLUCION DE COMPRAS"
        }
        for grupo in ("Emitido", "Recibido"):
            datos = resumen["grupos"][grupo]
            hoja.cell(row=fila, column=1, value=nombres_conceptos[grupo]).border = borde_fino
            hoja.cell(row=fila, column=2, value=f"{datos['facturas_cantidad']} / {datos['facturas_iva']:,.2f}").border = borde_fino
            hoja.cell(row=fila, column=3, value=f"{datos['notas_cantidad']} / {datos['notas_iva']:,.2f}").border = borde_fino
            c4 = hoja.cell(row=fila, column=4, value=datos["iva_neto"])
            c4.number_format = formato_moneda
            c4.border = borde_fino
            fila += 1

        totales = resumen["totales"]
        hoja.cell(row=fila, column=1, value="IVA generado").font = estilo_total
        hoja.cell(row=fila, column=1).fill = relleno_total
        celda_total = hoja.cell(row=fila, column=4, value=totales["iva_a_pagar"])
        celda_total.font = estilo_total
        celda_total.fill = relleno_total
        celda_total.number_format = formato_moneda
        for c in range(1, 5):
            hoja.cell(row=fila, column=c).border = borde_fino
        fila += 3

        hoja.cell(row=fila, column=1, value="DETALLE DE REGISTROS")
        hoja.cell(row=fila, column=1).font = estilo_encabezado_seccion
        hoja.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=len(df.columns))
        for c in range(1, len(df.columns) + 1):
            hoja.cell(row=fila, column=c).fill = relleno_seccion
        fila += 1

        columnas_detalle = list(df.columns)
        for c, nombre_col in enumerate(columnas_detalle, start=1):
            celda = hoja.cell(row=fila, column=c, value=nombre_col)
            celda.font = estilo_encabezado_tabla
            celda.fill = relleno_encabezado_tabla
            celda.border = borde_fino
        fila_inicio_detalle = fila
        fila += 1

        for _, registro in df.iterrows():
            if registro[COLUMNA_GRUPO] == "Emitido":
                relleno_fila = relleno_ventas
            elif registro[COLUMNA_GRUPO] == "Recibido":
                relleno_fila = relleno_compras
            else:
                relleno_fila = None

            for c, nombre_col in enumerate(columnas_detalle, start=1):
                celda = hoja.cell(row=fila, column=c, value=registro[nombre_col])
                celda.border = borde_fino
                if relleno_fila:
                    celda.fill = relleno_fila
                    
                if nombre_col in (COLUMNA_IVA, "IVA Ajustado"):
                    celda.number_format = formato_moneda
            fila += 1

        for c, nombre_col in enumerate(columnas_detalle, start=1):
            max_len = max(
                [len(str(nombre_col))] + [len(str(v)) for v in df[nombre_col].astype(str).tolist()]
            )
            hoja.column_dimensions[get_column_letter(c)].width = min(max(max_len + 2, 12), 45)

        hoja.freeze_panes = hoja.cell(row=fila_inicio_detalle + 1, column=1)

        try:
            libro.save(self.ruta_archivo)
        except Exception as e:
            raise ErrorSistema(f"No se pudo guardar el informe en el archivo: {e}")