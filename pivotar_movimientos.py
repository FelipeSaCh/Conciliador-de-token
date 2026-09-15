"""
pivote_movimientos.py

Módulo desacoplado para la transformación y reestructuración (pivotado)
de hojas de movimientos contables y exportación con formato profesional en Excel.
"""

from pathlib import Path
from typing import Callable, List, Optional, Union
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from errors import ErrorSistema, ErrorUsuario, logger


class TransformadorMovimientos:
    """Clase encargada de la lógica de negocio y transformación de datos
    para la reestructuración (pivotado) de hojas de movimientos contables.
    """

    COLUMNAS_REQUERIDAS = ["DEBITO", "CREDITO", "CUENTA", "NOM. CUENTA", "TIPO"]
    PREFIJOS_PERMITIDOS = ("FC", "GS", "DC", "DS", "GC", "CG","GI")
    HOJA_DESTINO = "MOVS_AUD"

    def __init__(
        self,
        ruta_entrada: Union[str, Path],
        hoja_origen: str,
        progress_callback: Optional[Callable[[str], None]] = None,
    ):
        self.ruta_entrada = Path(ruta_entrada)
        self.hoja_origen = hoja_origen
        self.progress_callback = progress_callback

    def _notificar(self, mensaje: str) -> None:
        if self.progress_callback:
            self.progress_callback(mensaje)
        logger.info(mensaje)

    def ejecutar(self) -> Path:
        """Ejecuta el flujo completo de lectura, filtrado, transformación, pivotado y diseño Excel."""
        try:
            df = self._cargar_y_validar()
            df_final, cols_indice = self._transformar_y_pivotar(df)
            self._guardar_con_formato(df_final, cols_indice)

            self._notificar(f"Guardado finalizado exitosamente en la hoja '{self.HOJA_DESTINO}'.")
            return self.ruta_entrada

        except ErrorUsuario:
            raise
        except Exception as e:
            logger.exception("Error durante la reestructuración de movimientos.")
            raise ErrorSistema(f"Error inesperado al procesar la hoja de movimientos: {e}") from e

    def _cargar_y_validar(self) -> pd.DataFrame:
        self._notificar("Leyendo la hoja de movimientos...")
        if not self.ruta_entrada.exists():
            raise ErrorUsuario(f"El archivo origen no existe: {self.ruta_entrada}")

        try:
            df = pd.read_excel(self.ruta_entrada, sheet_name=self.hoja_origen)
        except ValueError as e:
            raise ErrorUsuario(f"La hoja '{self.hoja_origen}' no existe en el archivo.") from e
        except Exception as e:
            raise ErrorSistema(f"No se pudo leer el archivo de Excel: {e}") from e

        # Verificar columnas obligatorias iniciales
        faltantes = [col for col in self.COLUMNAS_REQUERIDAS if col not in df.columns]
        if faltantes:
            raise ErrorUsuario(
                f"La hoja seleccionada no contiene las columnas requeridas: {', '.join(faltantes)}"
            )

        return df

    def _transformar_y_pivotar(self, df: pd.DataFrame) -> tuple[pd.DataFrame, List[str]]:
        self._notificar("Estandarizando nombres de columnas y limpiando basura...")
        df_calc = df.copy()

        # 1. Renombrar columnas a un estándar limpio
        renombres = {
            "TIPO": "Tipo",
            "NUMERO": "Número",
            "NO.EXT": "Num.Ext",
            "Nit/C.C.": "NIT",
            "NIT": "Nit/C.C.",
            "Nombre Emisor": "Tercero",
            "TERCERO": "Tercero",
            "AÑO": "Año",
            "MES": "Mes",
            "DIA": "Dia"
        }
        df_calc.rename(columns=renombres, inplace=True, errors="ignore")

        # 2. Filtrado de filas por prefijo en la columna 'Tipo'
        if "Tipo" not in df_calc.columns:
            raise ErrorUsuario("El archivo no contiene la columna 'TIPO' o 'Tipo'.")

        tipo_clean = df_calc["Tipo"].astype(str).str.strip().str.upper()
        mask_prefijo = tipo_clean.str.startswith(self.PREFIJOS_PERMITIDOS)
        df_calc = df_calc[mask_prefijo].copy()

        if df_calc.empty:
            raise ErrorUsuario(
                "No se encontraron registros que inicien con los prefijos permitidos (FC, GS, DC, DS, GC)."
            )

        # --- NUEVO: Ignorar por completo los registros anulados ---
        if "DETALLE" in df_calc.columns:
            mask_anulado = df_calc["DETALLE"].astype(str).str.contains("ANULADO", case=False, na=False)
            df_calc = df_calc[~mask_anulado].copy()

        # 3. Eliminar columnas que no se necesitan (SE ELIMINÓ 'DETALLE' DE LA LISTA)
        cols_a_borrar = ["CHEQ.NO", "ELABORO", "SUC.PTO","CCOSTO","SCCOSTO","REF1","REF2","REF3","REF4","No. AUTORIZACION(DATAFONO)"]
        df_calc.drop(columns=[c for c in cols_a_borrar if c in df_calc.columns], inplace=True, errors='ignore')

        # 4. Consolidar AÑO, MES, DIA en "Fecha" (Ej: Ago-23-2026)
        if all(c in df_calc.columns for c in ["Año", "Mes", "Dia"]):
            meses_espanol = {
                1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
                7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic',
                '1': 'Ene', '2': 'Feb', '3': 'Mar', '4': 'Abr', '5': 'May', '6': 'Jun',
                '7': 'Jul', '8': 'Ago', '9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dic',
                '01': 'Ene', '02': 'Feb', '03': 'Mar', '04': 'Abr', '05': 'May', '06': 'Jun',
                '07': 'Jul', '08': 'Ago', '09': 'Sep'
            }
            # Mapeo del mes, día a 2 dígitos y año sin decimales (.0)
            mes_str = df_calc["Mes"].map(meses_espanol).fillna(df_calc["Mes"].astype(str))
            dia_str = df_calc["Dia"].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
            ano_str = df_calc["Año"].astype(str).str.replace(r'\.0$', '', regex=True)
            
            df_calc["Fecha"] = mes_str + "-" + dia_str + "-" + ano_str
            df_calc.drop(columns=["Año", "Mes", "Dia"], inplace=True)

        # 5. Limpieza profunda de Num.Ext
        if "Num.Ext" in df_calc.columns:
            df_calc["Num.Ext"] = df_calc["Num.Ext"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_calc["Num.Ext"] = df_calc["Num.Ext"].replace(['nan', 'None', '<NA>'], '')

        self._notificar("Aplicando cálculos numéricos y unificando registros duplicados...")

        # 6. Preparar valores numéricos
        df_calc["DEBITO"] = pd.to_numeric(df_calc["DEBITO"], errors="coerce").fillna(0)
        df_calc["CREDITO"] = pd.to_numeric(df_calc["CREDITO"], errors="coerce").fillna(0)
        df_calc["VALOR_NETO"] = df_calc["DEBITO"] - df_calc["CREDITO"]
        df_calc["CUENTA_CONCAT"] = df_calc["CUENTA"].astype(str) + " - " + df_calc["NOM. CUENTA"].astype(str)

        # 7. Definir partición de columnas
        claves_agrupacion = ["Tipo", "Número", "Num.Ext"]
        claves_agrupacion = [c for c in claves_agrupacion if c in df_calc.columns]

        cols_sistema = ["CUENTA", "NOM. CUENTA", "CUENTA_CONCAT", "DEBITO", "CREDITO", "VALOR_NETO"]
        cols_descriptivas = [c for c in df_calc.columns if c not in claves_agrupacion and c not in cols_sistema]

        # --- NUEVO: Selección de DETALLE evaluando la prioridad de la cuenta ---
        df_detalle_elegido = None
        if "DETALLE" in cols_descriptivas:
            cols_descriptivas.remove("DETALLE") # Evitamos que el `.first()` genérico tome cualquiera
            
            df_det_temp = df_calc[claves_agrupacion + ["CUENTA", "DETALLE"]].copy()
            prefijos_detalle = ('6', '5', '7', '15')
            
            # Asignamos prioridad 1 a las cuentas solicitadas, prioridad 2 a las demás
            df_det_temp["Prioridad"] = df_det_temp["CUENTA"].astype(str).str.strip().apply(
                lambda x: 1 if x.startswith(prefijos_detalle) else 2
            )
            
            # Al ordenar, los movimientos con "Prioridad 1" quedan arriba en su grupo
            df_det_temp.sort_values(by=claves_agrupacion + ["Prioridad"], inplace=True)
            
            # Al eliminar duplicados, siempre conservaremos el "DETALLE" del movimiento prioritario
            df_detalle_elegido = df_det_temp.drop_duplicates(subset=claves_agrupacion, keep='first')[claves_agrupacion + ["DETALLE"]]

        # 8. Unificar importes (Suma) pivotando ÚNICAMENTE por la llave principal
        df_pivot = pd.pivot_table(
            df_calc,
            index=claves_agrupacion,
            columns="CUENTA_CONCAT",
            values="VALOR_NETO",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()
        df_pivot.columns.name = None

        # 9. Recuperar datos descriptivos tomando la primera ocurrencia (.first()) para evitar desdoblamientos
        if cols_descriptivas:
            df_desc = df_calc.groupby(claves_agrupacion, as_index=False)[cols_descriptivas].first()
            df_final = pd.merge(df_pivot, df_desc, on=claves_agrupacion, how="left")
        else:
            df_final = df_pivot
            
        # --- NUEVO: Fusionar el DETALLE elegido y reincorporarlo ---
        if df_detalle_elegido is not None:
            df_final = pd.merge(df_final, df_detalle_elegido, on=claves_agrupacion, how="left")
            cols_descriptivas.append("DETALLE")

        # 10. Ordenar columnas (Descriptivas a la izquierda, Cuentas monetarias reordenadas a la derecha)
        # 10.1 Definir el orden fijo de las columnas descriptivas (Se añadió DETALLE)
        orden_ideal = ["Tipo", "Número", "Num.Ext", "Fecha", "Tercero", "DOCRELA", "NIT", "DETALLE"]
        
        # Recolectar las columnas descriptivas que existen en df_final y mantener el orden ideal
        cols_indice_ordenadas = [c for c in orden_ideal if c in df_final.columns]
        # Agregar cualquier otra descriptiva sobrante que no estaba en el orden_ideal
        cols_indice_ordenadas += [c for c in df_final.columns if c in cols_descriptivas and c not in cols_indice_ordenadas]
        
        # 10.2 Obtener la lista de las columnas de cuentas (las pivotadas)
        cols_cuentas = [c for c in df_final.columns if c not in cols_indice_ordenadas]
        
        # 10.3 Lógica de reordenamiento de cuentas basado en prefijos prioritarios
        prefijos_prioridad = ('2365', '2367', '2368', '15', '6', '5', '22')
        
        cuentas_priorizadas = []
        cuentas_restantes = []

        # Separar las cuentas conservando el nombre original exacto (la variable 'c')
        for c in cols_cuentas:
            c_str = str(c).strip()
            if c_str.startswith(prefijos_prioridad):
                cuentas_priorizadas.append(c) # Guardamos 'c' original, no 'c_str'
            else:
                cuentas_restantes.append(c)   # Guardamos 'c' original
                
        # Función auxiliar para determinar la jerarquía
        def obtener_jerarquia(cuenta):
            cuenta_str = str(cuenta).strip()
            for i, prefijo in enumerate(prefijos_prioridad):
                if cuenta_str.startswith(prefijo):
                    return i
            return 999
            
        # Ordenar priorizadas por jerarquía y alfabéticamente (evaluando como string seguro)
        cuentas_priorizadas.sort(key=lambda x: (obtener_jerarquia(x), str(x).strip()))
        
        # Ordenar las restantes alfabéticamente
        cuentas_restantes.sort(key=lambda x: str(x).strip())
        
        # Unir las listas reordenadas
        cols_cuentas_reordenadas = cuentas_priorizadas + cuentas_restantes

        # 10.4 Ensamblar el DataFrame final con el nuevo orden total
        df_final = df_final[cols_indice_ordenadas + cols_cuentas_reordenadas]

        return df_final, cols_indice_ordenadas

    def _guardar_con_formato(self, df_final: pd.DataFrame, cols_indice: List[str]) -> None:
        self._notificar(f"Guardando resultados en la hoja '{self.HOJA_DESTINO}' del archivo original...")
        self._notificar("Aplicando estilos personalizados, bandas alternas y ajuste dinámico de celdas...")

        # Definición de paleta de colores y estilos visuales
        FILL_HEADER = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")  # Azul oscuro / Slate
        FILL_ROW_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        FILL_ROW_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")  # Gris muy claro

        FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        FONT_BODY = Font(name="Calibri", size=10, color="000000")

        BORDER_THIN = Side(border_style="thin", color="E2E8F0")
        BORDER_HEADER = Side(border_style="medium", color="0F172A")
        BORDER_BODY_ROW = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)
        BORDER_HEADER_CELL = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_HEADER, bottom=BORDER_HEADER)

        ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
        ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

        # Determinar si hay que mantener macros o vba basado en la extensión
        engine_kwargs = {'keep_vba': True} if self.ruta_entrada.suffix.lower() == '.xlsm' else {}

        # SE MODIFICÓ AQUÍ: 'mode=a' y 'if_sheet_exists=replace' para guardar en el mismo libro
        with pd.ExcelWriter(self.ruta_entrada, engine="openpyxl", mode='a', if_sheet_exists='replace', engine_kwargs=engine_kwargs) as writer:
            df_final.to_excel(writer, index=False, sheet_name=self.HOJA_DESTINO)
            worksheet = writer.sheets[self.HOJA_DESTINO]

            max_row = df_final.shape[0] + 1
            max_col = df_final.shape[1]
            num_cols_indice = len(cols_indice)

            # 1. Estilizado de Encabezados con altura compacta
            max_lineas_encabezado = 1

            for col_idx, col_nombre in enumerate(df_final.columns, start=1):
                col_letter = get_column_letter(col_idx)
                celda = worksheet.cell(row=1, column=col_idx)

                celda.fill = FILL_HEADER
                celda.font = FONT_HEADER
                celda.alignment = ALIGN_CENTER
                celda.border = BORDER_HEADER_CELL

                texto_col = str(col_nombre)
                longitud_texto = len(texto_col)

                if col_idx <= num_cols_indice:
                    # Columnas base: ajuste por contenido (Usando .str.len() que es seguro contra floats/NaN)
                    if not df_final.empty:
                        max_len_series = df_final[col_nombre].astype(str).str.len().max()
                        longitud_datos = int(max_len_series) if pd.notna(max_len_series) else 0
                    else:
                        longitud_datos = 0
                        
                    ancho_calculado = max(longitud_datos, longitud_texto) + 3
                    worksheet.column_dimensions[col_letter].width = min(max(ancho_calculado, 12), 45)
                    lineas = (longitud_texto // 35) + 1
                else:
                    # Columnas pivotadas (Cuentas): ancho más amplio para evitar filas de 4 líneas
                    worksheet.column_dimensions[col_letter].width = 30
                    lineas = (longitud_texto // 30) + 1

                if lineas > max_lineas_encabezado:
                    max_lineas_encabezado = lineas

            # Limitar la altura de la fila 1 entre 24 y 32 puntos como máximo
            worksheet.row_dimensions[1].height = min(max(24, max_lineas_encabezado * 14), 32)
            
            # 2. Estilizado de Filas de Datos (Banded Rows / Striped)
            for row_idx in range(2, max_row + 1):
                fill_row = FILL_ROW_EVEN if row_idx % 2 == 0 else FILL_ROW_ODD
                worksheet.row_dimensions[row_idx].height = 20

                for col_idx in range(1, max_col + 1):
                    celda = worksheet.cell(row=row_idx, column=col_idx)
                    celda.fill = fill_row
                    celda.font = FONT_BODY
                    celda.border = BORDER_BODY_ROW

                    if col_idx <= num_cols_indice:
                        celda.alignment = ALIGN_LEFT
                    else:
                        celda.alignment = ALIGN_RIGHT
                        celda.number_format = '"$" #,##0.00'

            # 3. Congelar fila de encabezados
            worksheet.freeze_panes = "A2"