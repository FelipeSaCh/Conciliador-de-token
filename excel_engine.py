import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from config import (
    COLUMNAS_DIAN_VS_CONT,
    CURRENCY_FORMAT,
    OUTPUT_SHEETS_TO_HIDE,
    ORDEN,
    RED_FILL_COLOR,
)
from errors import ColumnaFaltanteError, ErrorSistema, ErrorUsuario, HojaNoEncontradaError, logger


class ConciliadorExcel:
    COLUMNAS_REQUERIDAS_PRINCIPAL = ['NIT Emisor', 'Prefijo', 'Folio', 'Total', 'IVA']
    COLUMNAS_REQUERIDAS_CONTA = ['NIT']
    COLUMNAS_REQUERIDAS_TERCEROS = ['NIT']

    def __init__(self, file_path, sheet_names, seriales_iva=None, seriales_base=None, progress_callback=None):
        self.file_path = Path(file_path)
        self.sheet_names = sheet_names
        self.progress_callback = progress_callback or (lambda mensaje: None)
        
        # Guardamos los seriales dinámicos provenientes de la interfaz. 
        # Si no envían nada, usamos listas vacías por precaución.
        self.seriales_iva = seriales_iva if seriales_iva is not None else []
        self.seriales_base = seriales_base if seriales_base is not None else []

    def _reportar(self, mensaje):
        logger.info(mensaje)
        self.progress_callback(mensaje)

    def _validar_hojas_existen(self, hojas_excel):
        for clave in ('principal', 'contabilidad', 'terceros', 'aud_comp'):
            nombre = self.sheet_names.get(clave)
            if not nombre:
                raise ErrorUsuario(f"No se especificó el nombre de la hoja '{clave}'.")
            if nombre not in hojas_excel:
                raise HojaNoEncontradaError(nombre, hojas_excel)

    @staticmethod
    def _limpiar_encabezados(df, mayus=False):
        cols = df.columns.astype(str).str.replace(r'[\r\n\t]', '', regex=True).str.strip()
        df.columns = cols.str.upper() if mayus else cols
        return df

    @staticmethod
    def _cargar_hoja_con_encabezado_variable(path, sheet):
        try:
            df_temp = pd.read_excel(path, sheet_name=sheet, header=None, nrows=10)
            header_idx = 0
            for idx, row in df_temp.iterrows():
                if row.dropna().count() > 2:
                    header_idx = idx
                    break
            return pd.read_excel(path, sheet_name=sheet, header=header_idx)
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja '{sheet}' con encabezado variable: {e}") from e

    def _validar_columnas(self, df, columnas_requeridas, nombre_hoja):
        for col in columnas_requeridas:
            if col not in df.columns:
                raise ColumnaFaltanteError(nombre_hoja, col)

    def ejecutar(self):
        if not self.file_path.exists():
            raise ErrorUsuario(f"El archivo no existe: {self.file_path}")

        try:
            xls = pd.ExcelFile(self.file_path)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo Excel: {e}") from e

        hojas_excel = xls.sheet_names
        self._validar_hojas_existen(hojas_excel)

        self._reportar("Cargando hojas del archivo...")
        try:
            df_principal = pd.read_excel(self.file_path, sheet_name=self.sheet_names['principal'])
            df_conta = pd.read_excel(self.file_path, sheet_name=self.sheet_names['contabilidad'])
            df_tercer = pd.read_excel(self.file_path, sheet_name=self.sheet_names['terceros'])
        except Exception as e:
            raise ErrorSistema(f"Error leyendo las hojas principales: {e}") from e

        df_aud_comp = self._cargar_hoja_con_encabezado_variable(self.file_path, self.sheet_names['aud_comp'])

        nombre_auto = self.sheet_names.get('autorretenedores')
        if nombre_auto and nombre_auto in hojas_excel:
            try:
                df_autoretenedores = pd.read_excel(self.file_path, sheet_name=nombre_auto)
                df_autoretenedores = self._limpiar_encabezados(df_autoretenedores, mayus=True)
            except Exception as e:
                logger.warning(f"No se pudo leer la hoja de autorretenedores '{nombre_auto}': {e}")
                df_autoretenedores = pd.DataFrame(columns=['NIT', 'COMENTARIO'])
        else:
            df_autoretenedores = pd.DataFrame(columns=['NIT', 'COMENTARIO'])

        df_conta = self._limpiar_encabezados(df_conta, mayus=True)
        df_principal = self._limpiar_encabezados(df_principal)
        df_tercer = self._limpiar_encabezados(df_tercer, mayus=True)
        df_aud_comp = self._limpiar_encabezados(df_aud_comp)
        df_aud_comp = df_aud_comp.dropna(subset=df_aud_comp.columns[:6], how='all').copy()

        self._validar_columnas(df_principal, self.COLUMNAS_REQUERIDAS_PRINCIPAL, self.sheet_names['principal'])
        self._validar_columnas(df_conta, self.COLUMNAS_REQUERIDAS_CONTA, self.sheet_names['contabilidad'])
        self._validar_columnas(df_tercer, self.COLUMNAS_REQUERIDAS_TERCEROS, self.sheet_names['terceros'])

        self._reportar("Cruzando datos entre hojas...")
        df = self._cruzar_datos(df_principal, df_conta, df_tercer)

        self._reportar("Calculando bases, totales y consecutivos...")
        df, es_personales = self._calcular_campos(df)

        df_resultado = df[~es_personales][['Prefijo', 'BASE', 'Num.Ext', 'TIPO-DETALLE']]
        for col in ORDEN:
            if col not in df.columns:
                df[col] = ''
        df_auditoria = df[~es_personales][ORDEN].copy()

        self._reportar("Procesando hoja de auditoría de comprobantes...")
            # Pasamos los seriales guardados en la instancia de la clase
        df_res_auditoria = self._procesar_aud_comp(df_aud_comp, self.seriales_iva, self.seriales_base)

        self._reportar("Unificando y conciliando DIAN vs Contabilidad...")
        df_dian_vs_cont, parejas_incompletas = self._unificar_dian_vs_cont(
            df_auditoria, df_aud_comp, df_res_auditoria, df_autoretenedores
        )

        self._reportar("Escribiendo resultados en el archivo Excel...")
        self._escribir_excel(
            df_principal_cols=df_principal.columns,
            df=df,
            df_resultado=df_resultado,
            df_auditoria=df_auditoria,
            df_res_auditoria=df_res_auditoria,
            df_dian_vs_cont=df_dian_vs_cont,
            es_personales=es_personales,
            parejas_incompletas=parejas_incompletas,
        )

        self._reportar("Proceso completado con éxito.")
        return {
            "filas_procesadas": len(df),
            "filas_personales": int(es_personales.sum()),
            "filas_sin_pareja": int(sum(parejas_incompletas)),
        }

    def _cruzar_datos(self, df_principal, df_conta, df_tercer):
        df_principal = df_principal.copy()
        df_conta = df_conta.copy()
        df_tercer = df_tercer.copy()

        df_principal['NIT_Emisor_clean'] = (
            df_principal['NIT Emisor'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        )
        df_conta['NIT_clean'] = (
            df_conta['NIT'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        )
        df_tercer['NIT_clean'] = (
            df_tercer['NIT'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        )

        df_conta_unica = df_conta.drop_duplicates(subset=['NIT_clean'])

        cols_a_traer = ['NIT_clean']
        if 'TIPO-DETALLE' in df_conta_unica.columns:
            cols_a_traer.append('TIPO-DETALLE')
        if 'TIPO' in df_conta_unica.columns:
            cols_a_traer.append('TIPO')

        df = df_principal.merge(
            df_conta_unica[cols_a_traer],
            left_on='NIT_Emisor_clean',
            right_on='NIT_clean',
            how='left'
        )

        if 'TIPO-DETALLE' not in df.columns:
            df['TIPO-DETALLE'] = ''
        if 'TIPO' not in df.columns:
            df['TIPO'] = ''
        df['TIPO-DETALLE'] = df['TIPO-DETALLE'].fillna('')
        df['TIPO'] = df['TIPO'].fillna('')

        df['CONCEPTO'] = np.where(
            df['TIPO-DETALLE'].astype(str).str.strip() == '',
            'CONCEPTO NO CREADO',
            df['TIPO-DETALLE'].astype(str).str.strip()
        )

        nits_terceros_unicos = set(df_tercer['NIT_clean'].unique())
        df['TERCERO'] = np.where(
            df['NIT_Emisor_clean'].isin(nits_terceros_unicos) & (df['NIT_Emisor_clean'] != ''),
            'CREADO',
            'NO CREADO'
        )
        return df

    @staticmethod
    def _calcular_campos(df):
        prefijo_texto = df['Prefijo'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
        folio_texto = df['Folio'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
        num_ext_base = prefijo_texto + folio_texto

        prefijo_esta_vacio = df['Prefijo'].isna() | (df['Prefijo'].astype(str).str.strip() == '')

        df['Num.Ext'] = np.where(
            prefijo_esta_vacio,
            num_ext_base.str.zfill(10),
            num_ext_base
        )

        df['BASE'] = np.where(df['Total'] == 0, 0, (df['Total'] - df['IVA']))

        es_personales = df['TIPO'].astype(str).str.strip().str.upper() == 'PERSONALES'
        return df, es_personales

    @staticmethod
    def _procesar_aud_comp(df_aud_comp, seriales_iva, seriales_base):
        # Ahora usamos los parámetros 'seriales_iva' y 'seriales_base' en lugar de los globales
        cols_iva = [
            col for col in df_aud_comp.columns
            if any(re.search(rf'\b{s}', str(col)) for s in seriales_iva)
        ]
        cols_base = [
            col for col in df_aud_comp.columns
            if any(re.search(rf'\b{s}', str(col)) for s in seriales_base)
        ]

        df_res_auditoria = pd.DataFrame()
        df_res_auditoria['Num.Ext'] = (
            df_aud_comp['Num.Ext'].fillna('').astype(str) if 'Num.Ext' in df_aud_comp.columns else ''
        )
        df_res_auditoria['IVA'] = (
            df_aud_comp[cols_iva].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if cols_iva else 0
        )
        df_res_auditoria['BASE'] = (
            df_aud_comp[cols_base].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) * -1
            if cols_base else 0
        )
        return df_res_auditoria[['Num.Ext', 'BASE', 'IVA']]

    @staticmethod
    def _unificar_dian_vs_cont(df_auditoria, df_aud_comp, df_res_auditoria, df_autoretenedores):
        df_dian_prep = df_auditoria.copy()
        df_dian_prep['Fecha'] = ''
        df_dian_prep['Prioridad_Fila'] = 1
        for col in COLUMNAS_DIAN_VS_CONT:
            if col not in df_dian_prep.columns:
                df_dian_prep[col] = ''

        df_cont_prep = pd.DataFrame()
        df_cont_prep['Tipo de documento'] = df_aud_comp['Tipo'] if 'Tipo' in df_aud_comp.columns else ''

        if 'Número' in df_aud_comp.columns:
            numero_limpio = (
                df_aud_comp['Número'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            )
            df_cont_prep['CUFE/CUDE'] = numero_limpio.apply(lambda x: x.zfill(10) if x.isdigit() else x)
        else:
            df_cont_prep['CUFE/CUDE'] = ''

        df_cont_prep['Num.Ext'] = df_aud_comp['Num.Ext'] if 'Num.Ext' in df_aud_comp.columns else ''
        df_cont_prep['Fecha'] = df_aud_comp['Fecha'] if 'Fecha' in df_aud_comp.columns else ''
        df_cont_prep['NIT Emisor'] = df_aud_comp['Nit/C.C.'] if 'Nit/C.C.' in df_aud_comp.columns else ''
        df_cont_prep['Nombre Emisor'] = df_aud_comp['Tercero'] if 'Tercero' in df_aud_comp.columns else ''
        df_cont_prep['BASE'] = df_res_auditoria['BASE']
        df_cont_prep['IVA'] = df_res_auditoria['IVA']
        df_cont_prep['Prioridad_Fila'] = 2
        for col in COLUMNAS_DIAN_VS_CONT:
            if col not in df_cont_prep.columns:
                df_cont_prep[col] = ''

        df_dian_prep['Num.Ext_Clean'] = df_dian_prep['Num.Ext'].fillna('').astype(str).str.strip()
        df_cont_prep['Num.Ext_Clean'] = df_cont_prep['Num.Ext'].fillna('').astype(str).str.strip()

        df_unificado = pd.concat([df_dian_prep, df_cont_prep], ignore_index=True)
        df_unificado['Conteo_Pareja'] = df_unificado.groupby('Num.Ext_Clean')['Num.Ext_Clean'].transform('count')

        df_unificado['NIT_Emisor_Clean'] = (
            df_unificado['NIT Emisor'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        )
        df_unificado['NIT_Emisor_Group'] = df_unificado.groupby('Num.Ext_Clean')['NIT_Emisor_Clean'].transform('max')

        if not df_autoretenedores.empty:
            col_nit_auto = next((c for c in df_autoretenedores.columns if 'NIT' in c), None)
            col_coment_auto = next((c for c in df_autoretenedores.columns if 'COMENT' in c), None)

            if col_nit_auto and col_coment_auto:
                df_auto_clean = df_autoretenedores[[col_nit_auto, col_coment_auto]].copy()
                df_auto_clean['NIT_Clean'] = (
                    df_auto_clean[col_nit_auto].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                )
                df_auto_clean = df_auto_clean.drop_duplicates(subset=['NIT_Clean'])

                df_unificado = df_unificado.merge(
                    df_auto_clean[['NIT_Clean', col_coment_auto]],
                    left_on='NIT_Emisor_Group',
                    right_on='NIT_Clean',
                    how='left'
                )

                cond_comentario = (df_unificado['Prioridad_Fila'] == 2) & (df_unificado[col_coment_auto].notna())
                df_unificado['NIT Receptor'] = np.where(
                    cond_comentario, df_unificado[col_coment_auto], df_unificado['NIT Receptor']
                )
                df_unificado = df_unificado.drop(columns=['NIT_Clean', col_coment_auto], errors='ignore')

        df_unificado['Tipo_Doc_Clean'] = df_unificado['Tipo de documento'].fillna('').astype(str).str.strip().str.upper()
        df_unificado['Tipo_Doc_Group'] = np.where(df_unificado['Tipo_Doc_Clean'] == 'FC1', 1, 2)
        df_unificado['Tipo_Doc_Group'] = df_unificado.groupby('Num.Ext_Clean')['Tipo_Doc_Group'].transform('min')

        df_unificado = df_unificado.sort_values(
            by=['Tipo_Doc_Group', 'NIT_Emisor_Group', 'Num.Ext_Clean', 'Prioridad_Fila'],
            ascending=[True, False, True, True]
        )

        parejas_incompletas = (df_unificado['Conteo_Pareja'] == 1).tolist()
        df_dian_vs_cont = df_unificado[COLUMNAS_DIAN_VS_CONT]
        return df_dian_vs_cont, parejas_incompletas

    def _escribir_excel(
        self, df_principal_cols, df, df_resultado, df_auditoria, df_res_auditoria,
        df_dian_vs_cont, es_personales, parejas_incompletas
    ):
        try:
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                cols_originales = [col for col in df_principal_cols if col != 'NIT_Emisor_clean']
                df_export_principal = df[list(cols_originales) + ['CONCEPTO', 'TERCERO']]

                nombre_principal = self.sheet_names['principal']
                df_export_principal.to_excel(writer, index=False, sheet_name=nombre_principal)
                df_resultado.to_excel(writer, index=False, sheet_name='Resultados')
                df_auditoria.to_excel(writer, index=False, sheet_name='auditoria')
                df_res_auditoria.to_excel(writer, index=False, sheet_name='resultados-auditoria')
                df_dian_vs_cont.to_excel(writer, index=False, sheet_name='DIAN VS CONT')

                wb = writer.book
                sheet_principal = writer.sheets[nombre_principal]
                sheet_resultados = writer.sheets['Resultados']
                sheet_auditoria = writer.sheets['auditoria']
                sheet_res_auditoria = writer.sheets['resultados-auditoria']
                sheet_dian_vs_cont = writer.sheets['DIAN VS CONT']

                red_fill = PatternFill(start_color=RED_FILL_COLOR, end_color=RED_FILL_COLOR, fill_type="solid")

                for row_idx, es_pers in enumerate(es_personales, start=2):
                    if es_pers:
                        for col_idx in range(1, len(df_export_principal.columns) + 1):
                            sheet_principal.cell(row=row_idx, column=col_idx).fill = red_fill

                for row_idx, es_incompleta in enumerate(parejas_incompletas, start=2):
                    if es_incompleta:
                        for col_idx in range(1, len(COLUMNAS_DIAN_VS_CONT) + 1):
                            sheet_dian_vs_cont.cell(row=row_idx, column=col_idx).fill = red_fill

                for sheet_target, df_target in [
                    (sheet_resultados, df_resultado),
                    (sheet_auditoria, df_auditoria),
                    (sheet_res_auditoria, df_res_auditoria),
                    (sheet_dian_vs_cont, df_dian_vs_cont)
                ]:
                    for col_name in ['BASE', 'IVA', 'Total']:
                        if col_name in df_target.columns:
                            col_idx = df_target.columns.get_loc(col_name) + 1
                            for row in range(2, len(df_target) + 2):
                                sheet_target.cell(row=row, column=col_idx).number_format = CURRENCY_FORMAT

                max_row = len(df_dian_vs_cont) + 1
                max_col_letter = get_column_letter(len(COLUMNAS_DIAN_VS_CONT))
                tabla = Table(displayName="TablaDianVsCont", ref=f"A1:{max_col_letter}{max_row}")
                tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                sheet_dian_vs_cont.add_table(tabla)

                for col in sheet_dian_vs_cont.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    sheet_dian_vs_cont.column_dimensions[col_letter].width = max(max_len + 3, 12)

                for sheetname in OUTPUT_SHEETS_TO_HIDE:
                    if sheetname in wb.sheetnames:
                        wb[sheetname].sheet_state = 'hidden'
        except (ErrorUsuario, ErrorSistema):
            raise
        except Exception as e:
            raise ErrorSistema(f"Error escribiendo resultados en el archivo Excel: {e}") from e
