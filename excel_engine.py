# auditoria_engine.py
import re
import difflib
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries
from manager_autorretenedores import GestorAutorretenedores

from config import COLUMNAS_DIAN_VS_CONT, OUTPUT_SHEETS_TO_HIDE, ORDEN, RED_FILL_COLOR, CARACTERES_ESPECIALES, DUPLICADO_FILL_COLOR
from errors import ErrorSistema, ErrorUsuario, HojaNoEncontradaError, logger


class ConciliadorAuditoria:
    def __init__(self, file_path, sheet_names, seriales_iva=None, seriales_base=None,
                 seriales_base2=None, seriales_ret=None, seriales_iva_dc=None,
                 seriales_base_dc=None, seriales_base2_dc=None, progress_callback=None):
        self.file_path = Path(file_path)
        self.sheet_names = sheet_names
        self.progress_callback = progress_callback or (lambda mensaje: None)
        self.seriales_iva = seriales_iva if seriales_iva is not None else []
        self.seriales_base = seriales_base if seriales_base is not None else []
        self.seriales_base2 = seriales_base2 if seriales_base2 is not None else []
        self.seriales_ret = seriales_ret if seriales_ret is not None else []
        self.seriales_iva_dc = seriales_iva_dc if seriales_iva_dc is not None else []
        self.seriales_base_dc = seriales_base_dc if seriales_base_dc is not None else []
        self.seriales_base2_dc = seriales_base2_dc if seriales_base2_dc is not None else []

    def _reportar(self, mensaje):
        logger.info(mensaje)
        self.progress_callback(mensaje)

    def _validar_hojas_existen(self, hojas_excel):
        for clave in ('principal', 'aud_comp'):
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
    def _shift_table_ref(ref, pivot_col, delta):
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        if pivot_col <= max_col:
            max_col += delta
        if delta > 0 and pivot_col <= min_col:
            min_col += delta
        return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

    @staticmethod
    def _cargar_hoja_con_encabezado_variable(path, sheet):
        try:
            df_temp = pd.read_excel(path, sheet_name=sheet, header=None, nrows=10)
            # Encuentra el índice de la primera fila que tenga más de 2 valores no nulos
            filas_validas = df_temp.dropna(thresh=3)
            header_idx = filas_validas.index[0] if not filas_validas.empty else 0
            
            return pd.read_excel(path, sheet_name=sheet, header=header_idx)
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja '{sheet}' con encabezado variable: {e}") from e

    @staticmethod
    def _normalizar_num_ext(serie, patron_regex):
        """
        Normaliza Num.Ext para comparación/emparejamiento:
        limpia caracteres especiales y espacios. 
        Se eliminan SOLO los ceros a la izquierda para evitar colisiones 
        y falsos duplicados (ej: COTE21107 vs COTE21170).
        """
        return (
            serie
            .fillna('')
            .astype(str)
            .str.replace(patron_regex, '', regex=True)
            .str.strip()
            .str.replace(r'^0+', '', regex=True) # Elimina únicamente ceros a la izquierda
        )

    @staticmethod
    def _detectar_duplicados(df, columna_num_ext):
        """
        Detecta registros duplicados basados en Num.Ext (limpio)
        """
        patron_regex = f"[{re.escape(''.join(CARACTERES_ESPECIALES))}]"

        num_ext_clean = ConciliadorAuditoria._normalizar_num_ext(df[columna_num_ext], patron_regex)

        mask_no_vacio = num_ext_clean != ''
        mask_duplicados = num_ext_clean.duplicated(keep=False) & mask_no_vacio
        
        return mask_duplicados

    @staticmethod
    def _extraer_concepto_personal(valor_tipo):
        """
        Extrae el concepto específico de personal del valor TIPO
        Si el valor contiene "PERSONAL", extrae la parte relevante
        """
        if pd.isna(valor_tipo) or valor_tipo == '':
            return 'Gasto Personal'
        
        valor_str = str(valor_tipo).strip().upper()
        
        # Buscar el texto que contiene "PERSONAL"
        if 'PERSONAL' in valor_str:
            if ' - ' in valor_str:
                partes = valor_str.split(' - ')
                for parte in partes:
                    if 'PERSONAL' not in parte:
                        return parte.strip()
            
            if valor_str.startswith('PERSONAL'):
                resto = valor_str.replace('PERSONAL', '').strip()
                if resto:
                    return resto
            
            if valor_str.endswith('PERSONAL'):
                resto = valor_str.replace('PERSONAL', '').strip()
                if resto:
                    return resto
            
            return 'PERSONAL'
        
        return 'Gasto Personal'

    def ejecutar(self):
        if not self.file_path.exists():
            raise ErrorUsuario(f"El archivo no existe: {self.file_path}")

        try:
            xls = pd.ExcelFile(self.file_path)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo Excel: {e}") from e

        hojas_excel = xls.sheet_names
        self._validar_hojas_existen(hojas_excel)

        self._reportar("Cargando hoja de Token procesado...")
        try:
            df_full = pd.read_excel(self.file_path, sheet_name=self.sheet_names['principal'])

            ultimo_indice_valido = df_full.dropna(how='all').index.max()
            if pd.notna(ultimo_indice_valido):
                df_full = df_full.loc[:ultimo_indice_valido].copy()

            df_full = self._limpiar_encabezados(df_full)
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja principal: {e}") from e

        if 'Prefijo' not in df_full.columns:
            df_full['Prefijo'] = ''

        if 'Num.Ext' not in df_full.columns:
            col_folio = next((c for c in df_full.columns if str(c).strip().upper() == 'FOLIO'), None)
            if not col_folio:
                raise ErrorUsuario(
                    "La hoja de Token no contiene 'Num.Ext' ni la columna 'Folio' para construirlo."
                )
            prefijo_texto = df_full['Prefijo'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
            folio_texto = df_full[col_folio].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
            df_full['Num.Ext'] = prefijo_texto + folio_texto

        if 'BASE' not in df_full.columns:
            if 'Total' in df_full.columns and 'IVA' in df_full.columns:
                total_num = pd.to_numeric(df_full['Total'], errors='coerce').fillna(0)
                iva_num = pd.to_numeric(df_full['IVA'], errors='coerce').fillna(0)
                df_full['BASE'] = np.where(total_num == 0, 0, total_num - iva_num)
            else:
                df_full['BASE'] = 0

        if 'CONCEPTO' not in df_full.columns:
            df_full['CONCEPTO'] = ''

        cols_impuestos = [
            "ICA", "IC", "INC", "Timbre", "INC Timbre", "INC Bolsas", "IN Carbono",
            "IN Combustibles", "IC Datos", "ICL", "INPP", "IBUA", "ICUI",
            "Rete IVA", "Rete Renta", "Rete ICA"
        ]
        cols_existentes = [c for c in df_full.columns if str(c).strip().upper() in [x.upper() for x in cols_impuestos]]

        df_full['OTROS IMPUESTOS'] = df_full[cols_existentes].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)

        patron_regex = f"[{re.escape(''.join(CARACTERES_ESPECIALES))}]"
        df_full['Num.Ext_Clean_Temp'] = self._normalizar_num_ext(df_full['Num.Ext'], patron_regex)
        mapa_otros_imp = df_full.groupby('Num.Ext_Clean_Temp')['OTROS IMPUESTOS'].sum().to_dict()
        if '' in mapa_otros_imp:
            del mapa_otros_imp['']

# --- FILTRO ROBUSTO PARA ELIMINAR FILAS FANTASMAS EN EL TOKEN ---
        cols_token = [c for c in ['CONCEPTO', 'TERCERO', 'TIPO', 'BASE', 'Num.Ext'] if c in df_full.columns]
        if cols_token:
            mask_token_vacias = df_full[cols_token].apply(
                lambda c: c.fillna('').astype(str).str.strip().str.lower().replace(['nan', 'none', 'nat', 'null'], '') == ''
            ).all(axis=1)
            df_full = df_full[~mask_token_vacias].reset_index(drop=True)

        self._reportar("Cargando hoja de auditoría de comprobantes...")
        df_aud_comp = self._cargar_hoja_con_encabezado_variable(self.file_path, self.sheet_names['aud_comp'])
        df_aud_comp = self._limpiar_encabezados(df_aud_comp)
        
        # --- NUEVO: FILTRO ROBUSTO PARA COMPROBANTES ---
        if len(df_aud_comp.columns) >= 6:
            cols_comp = df_aud_comp.columns[:6]
            mask_comp_vacias = df_aud_comp[cols_comp].apply(
                lambda c: c.fillna('').astype(str).str.strip().str.lower().replace(['nan', 'none', 'nat', 'null'], '') == ''
            ).all(axis=1)
            df_aud_comp = df_aud_comp[~mask_comp_vacias].reset_index(drop=True).copy()

        try:
            datos_auto = GestorAutorretenedores.obtener_todos()
            if datos_auto:
                registros = [
                    {
                        'NIT': str(d.get('nit', d.get('NIT', ''))).strip(),
                        'COMENTARIO': str(
                            d.get('COMENT/COMENTARIO', d.get('nombre', d.get('NOMBRE', '')))
                        ).strip(),
                    }
                    for d in datos_auto
                ]
                df_autoretenedores = pd.DataFrame(registros, columns=['NIT', 'COMENTARIO'])
            else:
                df_autoretenedores = pd.DataFrame(columns=['NIT', 'COMENTARIO'])
        except Exception as e:
            logger.warning(f"No se pudo cargar el listado de autorretenedores desde JSON: {e}")
            df_autoretenedores = pd.DataFrame(columns=['NIT', 'COMENTARIO'])

        self._reportar("Filtrando registros y mapeando motivos de diferencias...")
        nombre_dc = self.sheet_names.get('Aud_dc')
        df_aud_dc = None
        if nombre_dc and nombre_dc in hojas_excel:
            try:
                df_aud_dc = self._cargar_hoja_con_encabezado_variable(self.file_path, nombre_dc)
                df_aud_dc = self._limpiar_encabezados(df_aud_dc)
                
                # --- NUEVO: FILTRO ROBUSTO PARA DEVOLUCIONES ---
                if len(df_aud_dc.columns) >= 6:
                    cols_dc = df_aud_dc.columns[:6]
                    mask_dc_vacias = df_aud_dc[cols_dc].apply(
                        lambda c: c.fillna('').astype(str).str.strip().str.lower().replace(['nan', 'none', 'nat', 'null'], '') == ''
                    ).all(axis=1)
                    df_aud_dc = df_aud_dc[~mask_dc_vacias].reset_index(drop=True).copy()
            except Exception as e:
                logger.warning(f"No se pudo leer la hoja de auditoría de devoluciones '{nombre_dc}': {e}")
                df_aud_dc = None
        
# --- 1. IDENTIFICACIÓN PRIORITARIA DE PERSONALES ---
        # Primero que todo, preparamos la columna Novedad
        df_full['Novedad'] = ''

        # Marcamos los personales INMEDIATAMENTE para blindarlos
        es_personales_full = df_full['CONCEPTO'].astype(str).str.strip().str.upper().str.contains('PERSONAL', na=False)
        df_full.loc[es_personales_full, 'Novedad'] = 'DIAN: ' + df_full.loc[es_personales_full, 'CONCEPTO'].astype(str)
        self._reportar(f"Se identificaron y protegieron {es_personales_full.sum()} registros Personales primero.")

# --- 2. EXCLUSIÓN DE EMITIDOS Y APPLICATION RESPONSE ---
        col_grupo = next((c for c in df_full.columns if str(c).strip().lower() == 'grupo'), None)
        col_doc = next((c for c in df_full.columns if str(c).strip().lower() == 'tipo de documento'), None)

        if col_grupo:
            # Detectamos todos los "emitidos"
            mask_emi = df_full[col_grupo].astype(str).str.contains('emitido', case=False, na=False)
            
            if col_doc:
                # Detectamos cuáles son Documento Soporte
                mask_soporte = df_full[col_doc].astype(str).str.contains('soporte', case=False, na=False)
                # Excluimos los emitidos ÚNICAMENTE si NO son Documentos Soporte
                mask_a_eliminar = mask_emi & ~mask_soporte
            else:
                mask_a_eliminar = mask_emi
                
            # Eliminamos los emitidos (que no son soporte) SOLO si no fueron marcados previamente como personales
            df_full = df_full[~(mask_a_eliminar & (df_full['Novedad'] == ''))]

        if col_doc:
            mask_app_response = df_full[col_doc].astype(str).str.contains('Application response', case=False, na=False)
            # Eliminamos application response SOLO si no fueron marcados como personales
            df_full = df_full[~(mask_app_response & (df_full['Novedad'] == ''))]

        # Reseteamos el índice de la copia temporal tras borrar filas
        df_full = df_full.reset_index(drop=True)

        # RECALCULAMOS la máscara de personales (necesario tras el reset_index)
        es_personales_full = df_full['CONCEPTO'].astype(str).str.strip().str.upper().str.contains('PERSONAL', na=False)

        # --- 3. DETECCIÓN DE DUPLICADOS EN TOKEN ---
        if 'Num.Ext' in df_full.columns:
            mask_duplicados_token = self._detectar_duplicados(df_full, 'Num.Ext')
            # Aplicamos "DUPLICADO" solo a los que tienen Novedad vacío (respeta a los Personales)
            df_full.loc[mask_duplicados_token & (df_full['Novedad'] == ''), 'Novedad'] = 'DIAN - DUPLICADO'
            self._reportar(f"Se detectaron {mask_duplicados_token.sum()} registros duplicados en Token (excluyendo Personales)")

        es_personales_full = df_full['CONCEPTO'].astype(str).str.strip().str.upper().str.contains('PERSONAL', na=False)
        mask_personales_sin_detalle = es_personales_full & (df_full['Novedad'] == '')
        df_full.loc[mask_personales_sin_detalle, 'Novedad'] = 'DIAN: ' + df_full.loc[mask_personales_sin_detalle, 'CONCEPTO'].astype(str)

        mask_ignorar = pd.Series(False, index=df_full.index)

        col_tipo_doc = next((c for c in df_full.columns if str(c).strip().lower() == 'tipo de documento'), None)
        if col_tipo_doc:
            mask_app = df_full[col_tipo_doc].astype(str).str.contains('application response', case=False, na=False)
            mask_ignorar |= mask_app
            df_full.loc[mask_app & (df_full['Novedad'] == ''), 'Novedad'] = 'DIAN - Application Response (Evento)'
            
        mask_dian_sin_num = df_full['Num.Ext'].fillna('').astype(str).str.strip() == ''
        df_full.loc[mask_dian_sin_num & (df_full['Novedad'] == ''), 'Novedad'] = 'DIAN - Sin Num.Ext'

        df_proc = df_full[~mask_ignorar].copy()
        es_personales_proc = es_personales_full[~mask_ignorar]

        col_concepto_res = 'TIPO-DETALLE' if 'TIPO-DETALLE' in df_proc.columns else 'CONCEPTO'
        df_resultado = df_proc[~es_personales_proc][['Prefijo', 'BASE', 'Num.Ext', col_concepto_res]]
        if col_concepto_res != 'TIPO-DETALLE':
            df_resultado = df_resultado.rename(columns={col_concepto_res: 'TIPO-DETALLE'})

        for col in ORDEN:
            if col not in df_full.columns:
                df_full[col] = ''
        
        columnas_base = ORDEN + ['Novedad']
        if 'OTROS IMPUESTOS' not in columnas_base:
            columnas_base.append('OTROS IMPUESTOS')
        df_auditoria = df_full[columnas_base].copy()

        self._reportar("Procesando hoja de auditoría de comprobantes...")
        df_res_auditoria = self._procesar_aud_comp(
            df_aud_comp, self.seriales_iva, self.seriales_base, self.seriales_base2, self.seriales_ret,
            mapa_otros_imp=mapa_otros_imp  # <--- NUEVO
        )

        self._reportar("Unificando y conciliando DIAN vs Contabilidad...")
        df_dian_vs_cont, parejas_incompletas, dif_base, dif_iva, tiene_caracter_especial, dif_nit, nombre_emisor_group, prefix_doc_group, df_res_dc, box_color_list = self._unificar_dian_vs_cont(
            df_auditoria, df_aud_comp, df_res_auditoria, df_autoretenedores,
            df_aud_dc=df_aud_dc, seriales_iva_dc=self.seriales_iva_dc,
            seriales_base_dc=self.seriales_base_dc, seriales_base2_dc=self.seriales_base2_dc,
            mapa_otros_imp=mapa_otros_imp
        )

        self._reportar("Escribiendo resultados en el archivo Excel...")
        self._escribir_excel(
            df_resultado=df_resultado,
            df_auditoria=df_auditoria.drop(columns=['Novedad'], errors='ignore'),
            df_res_auditoria=df_res_auditoria,
            df_dian_vs_cont=df_dian_vs_cont,
            parejas_incompletas=parejas_incompletas,
            df_autoretenedores=df_autoretenedores,
            dif_base=dif_base,
            dif_iva=dif_iva,
            tiene_caracter_especial=tiene_caracter_especial,
            dif_nit=dif_nit,
            nombre_emisor_group=nombre_emisor_group,
            prefix_doc_group=prefix_doc_group,
            df_res_dc=df_res_dc,            
            nombre_dc=nombre_dc,
            box_color_list=box_color_list # <--- NUEVA LISTA ENVIADA
        )
        self._reportar("Proceso completado con éxito.")
        return {
            "filas_procesadas": len(df_full),
            "filas_personales": int(es_personales_full.sum()),
            "filas_sin_pareja": int(sum(parejas_incompletas)),
        }
    @staticmethod
    def _procesar_aud_comp(df_aud_comp, seriales_iva, seriales_base, seriales_base2, seriales_ret=None, es_devolucion=False, mapa_otros_imp=None):
        cols_iva = [col for col in seriales_iva if col in df_aud_comp.columns]
        cols_base = [col for col in seriales_base if col in df_aud_comp.columns]
        cols_base2 = [col for col in seriales_base2 if col in df_aud_comp.columns]

        df_res_auditoria = pd.DataFrame(index=df_aud_comp.index)

        df_res_auditoria['Num.Ext'] = (
            df_aud_comp['Num.Ext'].fillna('').astype(str) if 'Num.Ext' in df_aud_comp.columns else ''
        )
        
        # SIEMPRE negativo para los valores de auditoría (contabilidad y devoluciones)
        multiplicador = -1

        df_res_auditoria['IVA'] = (
            df_aud_comp[cols_iva].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).abs() * multiplicador if cols_iva else 0
        )
        df_res_auditoria['BASE'] = (
            df_aud_comp[cols_base].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).abs() * multiplicador if cols_base else 0
        )
        df_res_auditoria['BASE_2'] = (
            df_aud_comp[cols_base2].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).abs() * multiplicador if cols_base2 else 0
        )

        if mapa_otros_imp is not None:
            patron_regex = f"[{re.escape(''.join(CARACTERES_ESPECIALES))}]"
            num_ext_clean_aud = ConciliadorAuditoria._normalizar_num_ext(df_res_auditoria['Num.Ext'], patron_regex)
            df_res_auditoria['OTROS IMPUESTOS'] = num_ext_clean_aud.map(mapa_otros_imp).fillna(0)
        else:
            df_res_auditoria['OTROS IMPUESTOS'] = 0

        return df_res_auditoria[['Num.Ext', 'BASE', 'BASE_2', 'IVA', 'OTROS IMPUESTOS']]
    @staticmethod

    def _conciliar_con_devoluciones(df_unificado, df_aud_dc, seriales_iva_dc, seriales_base_dc, seriales_base2_dc, patron_regex, mapa_otros_imp=None):
        
        # En la primera línea, donde se llama a _procesar_aud_comp:
        df_res_dc = ConciliadorAuditoria._procesar_aud_comp(
            df_aud_dc, seriales_iva_dc or [], seriales_base_dc or [], seriales_base2_dc or [], es_devolucion=True, mapa_otros_imp=mapa_otros_imp
        )

        df_dc_prep = pd.DataFrame()
        df_dc_prep['Tipo de documento'] = df_aud_dc['Tipo'] if 'Tipo' in df_aud_dc.columns else ''
        if 'Número' in df_aud_dc.columns:
            df_dc_prep['CUFE/CUDE'] = df_aud_dc['Número'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        else:
            df_dc_prep['CUFE/CUDE'] = ''

        # --- NUEVO: Extraer Detalle original de la auditoría de devoluciones ---
        col_det_dc_orig = next((c for c in df_aud_dc.columns if str(c).strip().upper() == 'DETALLE'), None)
        df_dc_prep['Detalle'] = df_aud_dc[col_det_dc_orig] if col_det_dc_orig else ''

        df_dc_prep['Num.Ext'] = df_aud_dc['Num.Ext'] if 'Num.Ext' in df_aud_dc.columns else ''
        df_dc_prep['Num.Ext_Original'] = df_dc_prep['Num.Ext']
        df_dc_prep['Num.Ext_Clean'] = ConciliadorAuditoria._normalizar_num_ext(df_dc_prep['Num.Ext'], patron_regex)
        df_dc_prep['Tiene_Caracter_Especial'] = (
            df_dc_prep['Num.Ext_Original'].fillna('').astype(str).str.contains(patron_regex, regex=True, na=False)
        )

        mask_dup_dc = False
        if 'Num.Ext' in df_aud_dc.columns:
            num_ext_clean_dc = ConciliadorAuditoria._normalizar_num_ext(df_aud_dc['Num.Ext'], patron_regex)
            mask_dup_dc = num_ext_clean_dc.duplicated(keep=False) & (num_ext_clean_dc != '')

        df_dc_prep['Novedad'] = np.where(
            mask_dup_dc, 'DEVOL - DUPLICADO',
            np.where(df_dc_prep['Num.Ext_Clean'] == '', 'DEVOL - Sin Num.Ext', '')
        )

        df_dc_prep['Fecha'] = df_aud_dc['Fecha'] if 'Fecha' in df_aud_dc.columns else ''
        df_dc_prep['NIT Emisor'] = df_aud_dc['Nit/C.C.'] if 'Nit/C.C.' in df_aud_dc.columns else ''
        df_dc_prep['Nombre Emisor'] = df_aud_dc['Tercero'] if 'Tercero' in df_aud_dc.columns else ''
        df_dc_prep['BASE'] = -pd.to_numeric(df_res_dc['BASE'], errors='coerce').fillna(0).abs()
        df_dc_prep['BASE_2'] = -pd.to_numeric(df_res_dc['BASE_2'], errors='coerce').fillna(0).abs()
        df_dc_prep['IVA'] = -pd.to_numeric(df_res_dc['IVA'], errors='coerce').fillna(0).abs()
        df_dc_prep['OTROS IMPUESTOS'] = -pd.to_numeric(df_res_dc['OTROS IMPUESTOS'], errors='coerce').fillna(0).abs()
        df_dc_prep['Prioridad_Fila'] = 3
        df_dc_prep['Conteo_Pareja'] = 1
        df_dc_prep['Fue_Fuzzy_Match'] = False

        for col in COLUMNAS_DIAN_VS_CONT:
            if col not in df_dc_prep.columns:
                df_dc_prep[col] = ''

        mask_dian_pendiente = (df_unificado['Prioridad_Fila'] == 1) & (df_unificado['Novedad'] == 'DIAN - Sin Pareja en Contabilidad')
        dian_idx = df_unificado[mask_dian_pendiente].index.tolist()
        dc_matchable_mask = df_dc_prep['Novedad'] == ''
        dc_idx = df_dc_prep[dc_matchable_mask].index.tolist()

        emparejados_dian = set()
        emparejados_dc = set()

        dc_por_clave = {}
        for i in dc_idx:
            clave = df_dc_prep.at[i, 'Num.Ext_Clean']
            if clave:
                dc_por_clave.setdefault(clave, []).append(i)

        for d_idx in dian_idx:
            clave = df_unificado.at[d_idx, 'Num.Ext_Clean']
            candidatos = [c for c in dc_por_clave.get(clave, []) if c not in emparejados_dc]
            if candidatos:
                c_idx = candidatos[0]
                emparejados_dian.add(d_idx)
                emparejados_dc.add(c_idx)
                df_unificado.at[d_idx, 'Novedad'] = ''
                df_dc_prep.at[c_idx, 'Novedad'] = ''

        dian_restantes = [i for i in dian_idx if i not in emparejados_dian]
        dc_restantes = [i for i in dc_idx if i not in emparejados_dc]

        if dian_restantes and dc_restantes:
            nit_dian = {i: str(df_unificado.at[i, 'NIT Emisor']).replace('.0', '').strip() for i in dian_restantes}
            nit_dc = {i: str(df_dc_prep.at[i, 'NIT Emisor']).replace('.0', '').strip() for i in dc_restantes}

            matches = []
            for d_idx in dian_restantes:
                base_d = pd.to_numeric(df_unificado.at[d_idx, 'BASE'], errors='coerce')
                base_d = 0.0 if pd.isna(base_d) else base_d
                str_d = df_unificado.at[d_idx, 'Num.Ext_Clean']
                for c_idx in dc_restantes:
                    if not nit_dian[d_idx] or nit_dian[d_idx] != nit_dc[c_idx]:
                        continue
                    base_c = pd.to_numeric(df_dc_prep.at[c_idx, 'BASE'], errors='coerce')
                    base_c = 0.0 if pd.isna(base_c) else base_c
                    if abs(base_d - base_c) > 50.0:
                        continue
                    str_c = df_dc_prep.at[c_idx, 'Num.Ext_Clean']
                    sim = difflib.SequenceMatcher(None, str_d, str_c).ratio()
                    if str_d.replace('0', '') == str_c.replace('0', '') and str_d.replace('0', '') != '':
                        sim = max(sim, 0.95)
                    if (str_d in str_c or str_c in str_d) and len(str_d) > 2 and len(str_c) > 2:
                        sim = max(sim, 0.90)
                    if sim > 0.40:
                        matches.append((sim, d_idx, c_idx))

            matches.sort(key=lambda x: x[0], reverse=True)
            for sim, d_idx, c_idx in matches:
                if d_idx not in emparejados_dian and c_idx not in emparejados_dc:
                    nueva_clave = f"{df_unificado.at[d_idx, 'Num.Ext_Clean']}_DEVOL_{d_idx}"
                    df_unificado.at[d_idx, 'Num.Ext_Clean'] = nueva_clave
                    df_dc_prep.at[c_idx, 'Num.Ext_Clean'] = nueva_clave
                    df_unificado.at[d_idx, 'Novedad'] = ''
                    df_dc_prep.at[c_idx, 'Novedad'] = ''
                    df_unificado.at[d_idx, 'Fue_Fuzzy_Match'] = True
                    df_dc_prep.at[c_idx, 'Fue_Fuzzy_Match'] = True
                    emparejados_dian.add(d_idx)
                    emparejados_dc.add(c_idx)

        for d_idx in dian_idx:
            if d_idx not in emparejados_dian:
                df_unificado.at[d_idx, 'Novedad'] = 'DIAN - Sin Pareja en Contabilidad ni Devoluciones'

        mask_dc_sin_pareja = dc_matchable_mask & (~df_dc_prep.index.isin(emparejados_dc))
        df_dc_prep.loc[mask_dc_sin_pareja, 'Novedad'] = 'DEVOL - Sin Pareja en DIAN'

        df_unificado = pd.concat([df_unificado, df_dc_prep], ignore_index=True)

        mask_matchable_final = df_unificado['Novedad'] == ''
        counts_final = df_unificado[mask_matchable_final].groupby('Num.Ext_Clean')['Num.Ext_Clean'].transform('count')
        df_unificado.loc[mask_matchable_final, 'Conteo_Pareja'] = counts_final

        return df_unificado, df_res_dc
    @staticmethod
    def _unificar_dian_vs_cont(df_auditoria, df_aud_comp, df_res_auditoria, df_autoretenedores,
                                df_aud_dc=None, seriales_iva_dc=None, seriales_base_dc=None, seriales_base2_dc=None, mapa_otros_imp=None):
        df_dian_prep = df_auditoria.copy()
        df_dian_prep['Fecha'] = ''
        df_dian_prep['Prioridad_Fila'] = 1
        
        # GARANTIZAR POSITIVOS PARA DIAN (TOKEN)
        df_dian_prep['BASE'] = pd.to_numeric(df_dian_prep['BASE'], errors='coerce').fillna(0).abs()
        if 'IVA' in df_dian_prep.columns:
            df_dian_prep['IVA'] = pd.to_numeric(df_dian_prep['IVA'], errors='coerce').fillna(0).abs()
        if 'BASE_2' in df_dian_prep.columns:
            df_dian_prep['BASE_2'] = pd.to_numeric(df_dian_prep['BASE_2'], errors='coerce').fillna(0).abs()
        
        if 'Num.Ext' in df_dian_prep.columns:
            df_dian_prep['Num.Ext_Original'] = df_dian_prep['Num.Ext']
        else:
            df_dian_prep['Num.Ext_Original'] = ''
            
        df_dian_prep['Detalle'] = ''
        
        for col in COLUMNAS_DIAN_VS_CONT:
            if col not in df_dian_prep.columns:
                df_dian_prep[col] = ''

        df_cont_prep = pd.DataFrame()
        df_cont_prep['Tipo de documento'] = df_aud_comp['Tipo'] if 'Tipo' in df_aud_comp.columns else ''
        if 'Número' in df_aud_comp.columns:
            numero_limpio = (
                df_aud_comp['Número'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            )
            df_cont_prep['CUFE/CUDE'] = numero_limpio
        else:
            df_cont_prep['CUFE/CUDE'] = ''
            
        # --- NUEVO: Extraer Detalle original de la auditoría de comprobantes ---
        col_det_orig = next((c for c in df_aud_comp.columns if str(c).strip().upper() == 'DETALLE'), None)
        df_cont_prep['Detalle'] = df_aud_comp[col_det_orig] if col_det_orig else ''

        patron_regex = f"[{re.escape(''.join(CARACTERES_ESPECIALES))}]"

        df_cont_prep['Num.Ext'] = df_aud_comp['Num.Ext'] if 'Num.Ext' in df_aud_comp.columns else ''
        df_cont_prep['Num.Ext_Original'] = df_cont_prep['Num.Ext']
        df_cont_prep['Num.Ext_Clean'] = ConciliadorAuditoria._normalizar_num_ext(df_cont_prep['Num.Ext'], patron_regex)
        
        df_cont_prep['Tiene_Caracter_Especial'] = (
            df_cont_prep['Num.Ext_Original']
            .fillna('')
            .astype(str)
            .str.contains(patron_regex, regex=True, na=False)
        )
        
        # --- DETECCIÓN DE DUPLICADOS EN AUDITORÍA ---
        mask_duplicados_aud = False
        if 'Num.Ext' in df_aud_comp.columns:
            num_ext_clean_aud = ConciliadorAuditoria._normalizar_num_ext(df_aud_comp['Num.Ext'], patron_regex)
            mask_no_vacio_aud = num_ext_clean_aud != ''
            mask_duplicados_aud = num_ext_clean_aud.duplicated(keep=False) & mask_no_vacio_aud
        
        df_cont_prep['Novedad'] = np.where(
            mask_duplicados_aud, 
            'CONT - DUPLICADO', 
            np.where(df_cont_prep['Num.Ext_Clean'] == '', 'CONT - Sin Num.Ext', '')
        )

        df_cont_prep['Fecha'] = df_aud_comp['Fecha'] if 'Fecha' in df_aud_comp.columns else ''
        df_cont_prep['NIT Emisor'] = df_aud_comp['Nit/C.C.'] if 'Nit/C.C.' in df_aud_comp.columns else ''
        df_cont_prep['Nombre Emisor'] = df_aud_comp['Tercero'] if 'Tercero' in df_aud_comp.columns else ''
        df_cont_prep['BASE'] = -pd.to_numeric(df_res_auditoria['BASE'], errors='coerce').fillna(0).abs()
        df_cont_prep['BASE_2'] = -pd.to_numeric(df_res_auditoria['BASE_2'], errors='coerce').fillna(0).abs()
        df_cont_prep['IVA'] = -pd.to_numeric(df_res_auditoria['IVA'], errors='coerce').fillna(0).abs()
        df_cont_prep['OTROS IMPUESTOS'] = -pd.to_numeric(df_res_auditoria['OTROS IMPUESTOS'], errors='coerce').fillna(0).abs()
        
        df_cont_prep['Prioridad_Fila'] = 2
        
        for col in COLUMNAS_DIAN_VS_CONT:
            if col not in df_cont_prep.columns:
                df_cont_prep[col] = ''

        df_dian_prep['Num.Ext_Clean'] = ConciliadorAuditoria._normalizar_num_ext(df_dian_prep['Num.Ext'], patron_regex)
        
        df_dian_prep['Tiene_Caracter_Especial'] = (
            df_dian_prep['Num.Ext_Original']
            .fillna('')
            .astype(str)
            .str.contains(patron_regex, regex=True, na=False)
        )

        df_dian_prep['Num.Ext'] = df_dian_prep['Num.Ext_Original']
        df_cont_prep['Num.Ext'] = df_cont_prep['Num.Ext_Original']

        df_unificado = pd.concat([df_dian_prep, df_cont_prep], ignore_index=True)

        col_nombre_emisor = 'Nombre Emisor' if 'Nombre Emisor' in df_unificado.columns else 'TERCERO'
        mask_totales = df_unificado[col_nombre_emisor].fillna('').astype(str).str.upper().str.contains('TOTALES')
        mask_ceros = df_unificado['Num.Ext_Clean'].fillna('').astype(str).str.contains('0000000000')
        
        df_unificado = df_unificado[~(mask_totales | mask_ceros)].reset_index(drop=True)
        
        # --- LÓGICA DE EMPAREJAMIENTO EXCLUSIVA PARA REGISTROS APTOS ---
        df_unificado['Conteo_Pareja'] = 1
        mask_matchable = (df_unificado['Num.Ext_Clean'] != '') & (df_unificado['Novedad'] == '')
        
        counts = df_unificado[mask_matchable].groupby('Num.Ext_Clean')['Num.Ext_Clean'].transform('count')
        df_unificado.loc[mask_matchable, 'Conteo_Pareja'] = counts

 # --- NUEVO: EMPAREJAMIENTO FLEXIBLE (FUZZY MATCHING) ---
        # Identificamos los que quedaron sin pareja en el primer intento
        mask_huerfano = mask_matchable & (df_unificado['Conteo_Pareja'] == 1)
        
        # CORRECCIÓN: Agrupamos los huérfanos SOLO por NIT. 
        # (Comparar valores exactos como texto causaba que fallara si había decimales de diferencia)
        nit_clean_fuzzy = df_unificado['NIT Emisor'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_unificado['Fuzzy_Key'] = nit_clean_fuzzy
        
        # Pre-calculamos las bases para compararlas matemáticamente
        df_unificado['Base_Num_Fuzzy'] = pd.to_numeric(df_unificado['BASE'], errors='coerce').fillna(0)
        df_unificado['Iva_Num_Fuzzy'] = pd.to_numeric(df_unificado['IVA'], errors='coerce').fillna(0) # <-- LÍNEA NUEVA
        df_unificado['Fue_Fuzzy_Match'] = False
        
        huerfanos_idx = df_unificado[mask_huerfano].index
        
        if len(huerfanos_idx) > 0:
            for _, group in df_unificado.loc[huerfanos_idx].groupby('Fuzzy_Key'):
                dian_idxs = group[group['Prioridad_Fila'] == 1].index.tolist()
                cont_idxs = group[group['Prioridad_Fila'] == 2].index.tolist()
                
                # Si hay al menos un huérfano de DIAN y uno de CONT en este mismo proveedor
                if dian_idxs and cont_idxs:
                    matches = []
                    for d_idx in dian_idxs:
                        for c_idx in cont_idxs:
                            str_d = df_unificado.at[d_idx, 'Num.Ext_Clean']
                            str_c = df_unificado.at[c_idx, 'Num.Ext_Clean']
                            
                            base_d = df_unificado.at[d_idx, 'Base_Num_Fuzzy']
                            base_c = df_unificado.at[c_idx, 'Base_Num_Fuzzy']
                            iva_d = df_unificado.at[d_idx, 'Iva_Num_Fuzzy']
                            iva_c = df_unificado.at[c_idx, 'Iva_Num_Fuzzy']
                            
                            # 1. Tolerancia matemática (DIAN es positivo, CONT es negativo, por eso se suman)
                            # Permitimos el emparejamiento si cuadra la BASE *O* si cuadra el IVA
                            if abs(base_d + base_c) > 50.0 and abs(iva_d + iva_c) > 50.0:
                                continue # Si el valor no cuadra ni en base ni en IVA, ignoramos y seguimos
                            # 2. Porcentaje de similitud básica de los folios
                            sim = difflib.SequenceMatcher(None, str_d, str_c).ratio()
                            
                            # 3. Bonificación si al quitar TODOS los ceros internos quedan idénticos
                            if str_d.replace('0', '') == str_c.replace('0', '') and str_d.replace('0', '') != '':
                                sim = max(sim, 0.95)
                                
                            # 4. Bonificación si un folio contiene al otro por completo
                            if (str_d in str_c or str_c in str_d) and len(str_d) > 2 and len(str_c) > 2:
                                sim = max(sim, 0.90)
                                
                            # Consideramos que son pareja si tienen al menos 40% de similitud
                            if sim > 0.40:
                                matches.append((sim, d_idx, c_idx))
                    
                    # Emparejamos los más similares primero
                    matches.sort(key=lambda x: x[0], reverse=True)
                    emparejados_d = set()
                    emparejados_c = set()
                    
                    for sim, d_idx, c_idx in matches:
                        if d_idx not in emparejados_d and c_idx not in emparejados_c:
                            # Asignamos una clave nueva idéntica para que el sistema los una
                            base_str = df_unificado.at[d_idx, 'Num.Ext_Clean']
                            nueva_clave = f"{base_str}_FUZZY_{d_idx}"
                            
                            df_unificado.at[d_idx, 'Num.Ext_Clean'] = nueva_clave
                            df_unificado.at[c_idx, 'Num.Ext_Clean'] = nueva_clave
                            df_unificado.at[d_idx, 'Fue_Fuzzy_Match'] = True
                            df_unificado.at[c_idx, 'Fue_Fuzzy_Match'] = True
                            
                            emparejados_d.add(d_idx)
                            emparejados_c.add(c_idx)
                            
        # Limpiamos las columnas auxiliares
        df_unificado = df_unificado.drop(columns=['Base_Num_Fuzzy', 'Fuzzy_Key'])
        
        # Recalcular el conteo final con las nuevas parejas formadas a la fuerza
        counts_re = df_unificado[mask_matchable].groupby('Num.Ext_Clean')['Num.Ext_Clean'].transform('count')
        df_unificado.loc[mask_matchable, 'Conteo_Pareja'] = counts_re
        # --- FIN EMPAREJAMIENTO FLEXIBLE ---

        mask_single = mask_matchable & (df_unificado['Conteo_Pareja'] == 1)
        is_dian = df_unificado['Prioridad_Fila'] == 1
        is_cont = df_unificado['Prioridad_Fila'] == 2
        
        df_unificado.loc[mask_single & is_dian, 'Novedad'] = 'DIAN - Sin Pareja en Contabilidad'
        df_unificado.loc[mask_single & is_cont, 'Novedad'] = 'CONT - Sin Pareja en DIAN'
        df_res_dc = None
        if df_aud_dc is not None and not df_aud_dc.empty:
            df_unificado, df_res_dc = ConciliadorAuditoria._conciliar_con_devoluciones(
                df_unificado, df_aud_dc, seriales_iva_dc, seriales_base_dc, seriales_base2_dc, patron_regex, mapa_otros_imp # <--- NUEVO
            )

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
        df_unificado['Es_Incompleta'] = df_unificado['Novedad'] != ''

        # Homologar Nombre Emisor
        df_unificado['Nombre_Emisor_Clean'] = df_unificado['Nombre Emisor'].fillna('').astype(str).str.strip().str.upper()
        df_unificado['Nombre_Emisor_Group'] = df_unificado.groupby('Num.Ext_Clean')['Nombre_Emisor_Clean'].transform('max')

        mask_cont = df_unificado['Prioridad_Fila'].isin([2, 3])
        df_unificado['Temp_Prefix'] = np.where(mask_cont, df_unificado['Tipo de documento'].astype(str).str.strip().str.upper().str[:2], '')
        df_unificado['Prefix_Doc_Group'] = df_unificado.groupby('Num.Ext_Clean')['Temp_Prefix'].transform('max')

        df_unificado['Sort_Matched_Tipo'] = np.where(~df_unificado['Es_Incompleta'], df_unificado['Prefix_Doc_Group'], '')
        df_unificado['Sort_Matched_Emisor'] = np.where(~df_unificado['Es_Incompleta'], df_unificado['Nombre_Emisor_Group'], '')
        df_unificado['Sort_Unmatched_Detalle'] = np.where(df_unificado['Es_Incompleta'], df_unificado['Novedad'], '')

        df_unificado = df_unificado.sort_values(
            by=['Es_Incompleta', 'Sort_Matched_Tipo', 'Sort_Matched_Emisor', 'Sort_Unmatched_Detalle', 'Nombre_Emisor_Group', 'Num.Ext_Clean', 'Prioridad_Fila'],
            ascending=[True, True, True, True, True, True, True]
        )

# --- REVISIÓN ROBUSTA DE DIFERENCIAS EN PAREJAS Y BORDES (NUEVO) ---
        base_val_numeric = pd.to_numeric(df_unificado['BASE'], errors='coerce').fillna(0)
        iva_val_numeric = pd.to_numeric(df_unificado['IVA'], errors='coerce').fillna(0)

        df_unificado['Temp_Base_Num'] = base_val_numeric
        df_unificado['Temp_IVA_Num'] = iva_val_numeric
        df_unificado['Temp_Total_Suma'] = base_val_numeric + iva_val_numeric  # SUMA DE 4 VARIABLES (2 Bases + 2 Ivas)
        
        mask_pairs = df_unificado['Conteo_Pareja'] >= 2
        
        group_base_sum = df_unificado[mask_pairs].groupby('Num.Ext_Clean')['Temp_Base_Num'].transform('sum')
        group_iva_sum = df_unificado[mask_pairs].groupby('Num.Ext_Clean')['Temp_IVA_Num'].transform('sum')
        group_total_sum = df_unificado[mask_pairs].groupby('Num.Ext_Clean')['Temp_Total_Suma'].transform('sum')

        UMBRAL_TOLERANCIA = 50.0
        df_unificado['Diff_BASE'] = False
        df_unificado['Diff_IVA'] = False
        df_unificado['Box_Color'] = '' # Variable para el color del cuadro alrededor de las 4 celdas
        
        # Diferencias individuales para pintar de rojo (se mantiene actual)
        df_unificado.loc[mask_pairs, 'Diff_BASE'] = group_base_sum.abs() > UMBRAL_TOLERANCIA
        df_unificado.loc[mask_pairs, 'Diff_IVA'] = group_iva_sum.abs() > UMBRAL_TOLERANCIA
        
        # Validación nueva: Evaluar si la suma general está en margen de tolerancia
        df_unificado.loc[mask_pairs, 'Box_Color'] = np.where(group_total_sum.abs() <= UMBRAL_TOLERANCIA, 'BLACK', 'YELLOW')

        df_unificado = df_unificado.drop(columns=['Temp_Base_Num', 'Temp_IVA_Num', 'Temp_Total_Suma'])

        # --- REVISIÓN DE TERCERO (NIT) EN PAREJAS ---
        df_unificado['Diff_NIT'] = False
        nit_nunique = df_unificado[mask_pairs].groupby('Num.Ext_Clean')['NIT_Emisor_Clean'].transform('nunique')
        mask_diff_nit = mask_pairs & (nit_nunique > 1)
        
        df_unificado.loc[mask_diff_nit, 'Novedad'] = np.where(
            df_unificado.loc[mask_diff_nit, 'Novedad'] == '',
            'No coincide tercero',
            df_unificado.loc[mask_diff_nit, 'Novedad'] + ' / No coincide tercero'
        )
        df_unificado['Diff_NIT'] = mask_diff_nit

        # --- NUEVO: REVISIÓN DE EMPAREJAMIENTO FLEXIBLE (DIF. FORMATO) ---
        mask_fue_fuzzy = df_unificado['Fue_Fuzzy_Match'] == True
        
        df_unificado.loc[mask_fue_fuzzy, 'Novedad'] = np.where(
            df_unificado.loc[mask_fue_fuzzy, 'Novedad'] == '',
            'Dif. formato de factura',
            df_unificado.loc[mask_fue_fuzzy, 'Novedad'] + ' / Dif. formato de factura'
        )
        
        # Pinta la celda de rojo al compartir la variable de caracteres especiales
        df_unificado['Tiene_Caracter_Especial'] = df_unificado['Tiene_Caracter_Especial'] | mask_fue_fuzzy

        # Extraer listas finales para el excel
        parejas_incompletas = df_unificado['Es_Incompleta'].tolist()
        dif_base_list = df_unificado['Diff_BASE'].tolist()
        dif_iva_list = df_unificado['Diff_IVA'].tolist()
        tiene_caracter_especial_list = df_unificado['Tiene_Caracter_Especial'].tolist()
        dif_nit_list = df_unificado['Diff_NIT'].tolist()
        nombre_emisor_group_list = df_unificado['Nombre_Emisor_Group'].tolist()
        prefix_doc_group_list = df_unificado['Prefix_Doc_Group'].tolist()

        columnas_finales = list(COLUMNAS_DIAN_VS_CONT)
        if 'OTROS IMPUESTOS' not in columnas_finales:
            columnas_finales.append('OTROS IMPUESTOS') # Añadimos la columna final
        if 'Novedad' not in columnas_finales:
            columnas_finales.append('Novedad')
        if 'Detalle' not in columnas_finales:
            columnas_finales.append('Detalle')

        df_dian_vs_cont = df_unificado[columnas_finales]
        box_color_list = df_unificado['Box_Color'].tolist()
        return df_dian_vs_cont, parejas_incompletas, dif_base_list, dif_iva_list, tiene_caracter_especial_list, dif_nit_list, nombre_emisor_group_list, prefix_doc_group_list, df_res_dc, box_color_list
    def _escribir_excel(
            self, df_resultado, df_auditoria, df_res_auditoria, df_dian_vs_cont,
            parejas_incompletas, df_autoretenedores=None, dif_base=None, dif_iva=None,
            tiene_caracter_especial=None, dif_nit=None, nombre_emisor_group=None,
            prefix_doc_group=None, df_res_dc=None, nombre_dc=None, box_color_list=None
        ):
        if prefix_doc_group is None: prefix_doc_group = []
        if df_autoretenedores is None:
            df_autoretenedores = pd.DataFrame(columns=['NIT', 'COMENTARIO'])
        if dif_base is None: dif_base = []
        if dif_iva is None: dif_iva = []
        if tiene_caracter_especial is None: tiene_caracter_especial = []
        if dif_nit is None: dif_nit = []
        if nombre_emisor_group is None: nombre_emisor_group = []
        if box_color_list is None: box_color_list = []

        try:
            engine_kwargs = {'keep_vba': True} if self.file_path.suffix.lower() == '.xlsm' else {}
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace',
                                 engine_kwargs=engine_kwargs) as writer:

                df_resultado.to_excel(writer, index=False, sheet_name='Resultados')
                df_auditoria.to_excel(writer, index=False, sheet_name='auditoria')
                df_dian_vs_cont.to_excel(writer, index=False, sheet_name='DIAN VS CONT')

                wb = writer.book
                sheet_resultados = writer.sheets['Resultados']
                sheet_auditoria = writer.sheets['auditoria']
                sheet_dian_vs_cont = writer.sheets['DIAN VS CONT']

                # --- 1. PREPARAR DICCIONARIO DE AUTORRETENEDORES (usa siempre COMENT/COMENTARIO como fuente) ---
                dict_auto = {}
                col_nit_auto = next((c for c in df_autoretenedores.columns if 'NIT' in str(c).upper()), None)
                col_coment_auto = next(
                    (c for c in df_autoretenedores.columns if 'COMENT' in str(c).upper()),
                    None
                )
                if col_coment_auto is None:
                    col_coment_auto = next(
                        (c for c in df_autoretenedores.columns if 'NOMBRE' in str(c).upper()),
                        None
                    )

                if col_nit_auto and col_coment_auto:
                    for _, r in df_autoretenedores.iterrows():
                        nit_str = str(r[col_nit_auto]).replace('.0', '').strip()
                        if nit_str and str(nit_str).lower() != 'nan':
                            dict_auto[nit_str] = str(r[col_coment_auto]).strip()

                # --- 2. FUNCIÓN INTERNA PARA INYECTAR COLUMNAS EN CUALQUIER HOJA ---
                def inyectar_hoja(sheet_name, df_res):
                    if not sheet_name or sheet_name not in wb.sheetnames or df_res is None or df_res.empty:
                        return
                    
                    sheet_orig = wb[sheet_name]
                    header_row = 0
                    col_tercero = col_nit = col_tipo = 0

                    for r in range(1, 20):
                        for c in range(1, sheet_orig.max_column + 1):
                            val = str(sheet_orig.cell(row=r, column=c).value).strip().upper()
                            if val == 'TERCERO': col_tercero = c
                            elif val in ['NIT/C.C.', 'NIT']: col_nit = c
                            elif val == 'TIPO': col_tipo = c

                        if col_tercero > 0:
                            header_row = r
                            break

                    if header_row > 0 and col_tercero > 0:
                        # Limpiar si ya fueron inyectadas previamente
                        num_injected_found = 0
                        for i in range(1, 6):
                            val = str(sheet_orig.cell(row=header_row, column=col_tercero + i).value).strip().upper()
                            if val in ['AUTORRETENCION', 'BASE', 'BASE_2', 'IVA', 'OTROS IMPUESTOS']:
                                num_injected_found += 1
                            else:
                                break
                        
                        if num_injected_found > 0:
                            sheet_orig.delete_cols(col_tercero + 1, amount=num_injected_found)

                        row_data_list = []
                        df_res_records = df_res.to_dict('records')
                        max_row_orig = sheet_orig.max_row
                        max_col_orig = sheet_orig.max_column

                        for idx, excel_row_idx in enumerate(range(header_row + 1, max_row_orig + 1)):
                            cell_values = [sheet_orig.cell(row=excel_row_idx, column=c).value for c in range(1, max_col_orig + 1)]
                            calc_data = df_res_records[idx] if idx < len(df_res_records) else {'BASE': 0, 'BASE_2': 0, 'IVA': 0, 'OTROS IMPUESTOS': 0}

                            nit_val = str(cell_values[col_nit - 1]).replace('.0', '').strip() if col_nit > 0 else ''
                            comentario_auto = dict_auto.get(nit_val, '')

                            tipo_val = str(cell_values[col_tipo - 1]).strip().upper() if col_tipo > 0 else ''
                            tercero_val = str(cell_values[col_tercero - 1]).strip().upper() if col_tercero > 0 else ''

                            row_data_list.append({
                                'original_values': cell_values,
                                'base': calc_data.get('BASE', 0),
                                'base_2': calc_data.get('BASE_2', 0),
                                'iva': calc_data.get('IVA', 0),
                                'otros': calc_data.get('OTROS IMPUESTOS', 0),
                                'auto': comentario_auto,
                                'tipo': tipo_val,
                                'tercero': tercero_val
                            })

                        max_base_por_tercero = {}
                        for row in row_data_list:
                            t_str = str(row.get('tercero', '')).strip().upper()
                            try:
                                b_val = float(row.get('base', 0))
                            except (ValueError, TypeError):
                                b_val = 0.0
                            
                            if t_str not in max_base_por_tercero or abs(b_val) > abs(max_base_por_tercero[t_str]):
                                max_base_por_tercero[t_str] = b_val

                        def sort_key(row):
                            auto_str = str(row.get('auto', '')).upper().strip()
                            tipo_str = str(row.get('tipo', '')).upper().strip()
                            tercero_str = str(row.get('tercero', '')).strip().upper()
                            
                            if 'AUTORRETENEDOR' in auto_str: 
                                prioridad_auto = 1
                            elif auto_str != '' or 'REGIMEN SIMPLE' in tipo_str: 
                                prioridad_auto = 2
                            else: 
                                prioridad_auto = 3
                                
                            try: 
                                base_val = float(row.get('base', 0))
                            except (ValueError, TypeError): 
                                base_val = 0.0
                                
                            if tipo_str.startswith('FC'):
                                prioridad_doc = 1
                            elif tipo_str.startswith('DC'):
                                prioridad_doc = 3
                            else:
                                prioridad_doc = 2
                                
                            if prioridad_doc in (1, 2):
                                agrupacion_tercero = ""
                            else:
                                agrupacion_tercero = tercero_str
                                
                            orden_base = -abs(base_val)
                            
                            return (prioridad_auto, prioridad_doc, agrupacion_tercero, orden_base, tercero_str)

                        row_data_list.sort(key=sort_key)
                        idx_insert = col_tercero + 1
                        sheet_orig.insert_cols(idx_insert, amount=5)

                        c_AUTORRETENCION = sheet_orig.cell(row=header_row, column=idx_insert)
                        c_AUTORRETENCION.value = 'AUTORRETENCION'
                        c_AUTORRETENCION.fill = PatternFill(start_color="cc99ff", end_color="cc99ff", fill_type="solid")

                        c_base = sheet_orig.cell(row=header_row, column=idx_insert + 1)
                        c_base.value = 'BASE'
                        c_base.fill = PatternFill(start_color="339966", end_color="339966", fill_type="solid")

                        c_base2 = sheet_orig.cell(row=header_row, column=idx_insert + 2)
                        c_base2.value = 'BASE_2'
                        c_base2.fill = PatternFill(start_color="33EAEA", end_color="33EAEA", fill_type="solid") 

                        c_iva = sheet_orig.cell(row=header_row, column=idx_insert + 3)
                        c_iva.value = 'IVA'
                        c_iva.fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

                        c_otros = sheet_orig.cell(row=header_row, column=idx_insert + 4)
                        c_otros.value = 'OTROS IMPUESTOS'
                        c_otros.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

                        for i, row_dict in enumerate(row_data_list):
                            excel_row_idx = header_row + 1 + i
                            orig_vals = row_dict['original_values']

                            for c in range(1, idx_insert):
                                sheet_orig.cell(row=excel_row_idx, column=c).value = orig_vals[c-1]

                            sheet_orig.cell(row=excel_row_idx, column=idx_insert).value = row_dict['auto']

                            c_b = sheet_orig.cell(row=excel_row_idx, column=idx_insert + 1)
                            c_b.value = row_dict['base']
                            c_b.number_format = '#,##0.00'

                            c_b2 = sheet_orig.cell(row=excel_row_idx, column=idx_insert + 2)
                            c_b2.value = row_dict['base_2']
                            c_b2.number_format = '#,##0.00'

                            c_i = sheet_orig.cell(row=excel_row_idx, column=idx_insert + 3)
                            c_i.value = row_dict['iva']
                            c_i.number_format = '#,##0.00'

                            c_o = sheet_orig.cell(row=excel_row_idx, column=idx_insert + 4)
                            c_o.value = row_dict['otros']
                            c_o.number_format = '#,##0.00'
                            
                            for c in range(idx_insert, len(orig_vals) + 1):
                                sheet_orig.cell(row=excel_row_idx, column=c + 5).value = orig_vals[c-1]

                # --- 3. APLICAR LA INYECCIÓN A AMBAS HOJAS ---
                inyectar_hoja(self.sheet_names['aud_comp'], df_res_auditoria)
                inyectar_hoja(nombre_dc, df_res_dc)

                # Formateo general numérico
                for sheet_target, df_target in [
                    (sheet_resultados, df_resultado),
                    (sheet_auditoria, df_auditoria),
                    (sheet_dian_vs_cont, df_dian_vs_cont)
                ]:
                    for col_name in ['BASE', 'BASE_2', 'IVA','OTROS IMPUESTOS', 'Total']:
                        if col_name in df_target.columns:
                            col_idx = df_target.columns.get_loc(col_name) + 1
                            for row in range(2, len(df_target) + 2):
                                sheet_target.cell(row=row, column=col_idx).number_format = '#,##0.00'
                                
                # --- 4. ESTILOS VISUALES Y CORRECCIÓN DE COLORES EN DIAN VS CONT ---
                max_row = len(df_dian_vs_cont) + 1
                max_col = len(df_dian_vs_cont.columns)

                sheet_dian_vs_cont.freeze_panes = 'A2'

                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")

                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                for col_idx in range(1, max_col + 1):
                    cell = sheet_dian_vs_cont.cell(row=1, column=col_idx)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = thin_border

                fill_red_alert = PatternFill(start_color=RED_FILL_COLOR, end_color=RED_FILL_COLOR, fill_type="solid")

                COLOR_MOTIVOS = {
                    'DIAN - Gasto Personal': 'FFCCFF',                
                    'DIAN - Application Response (Evento)': 'D9D9D9', 
                    'DIAN - Sin Num.Ext': 'F2DCDB',                   
                    'CONT - Sin Num.Ext': 'F2DCDB',                   
                    'DIAN - Sin Pareja en Contabilidad': 'FFF2CC',    
                    'CONT - Sin Pareja en DIAN': 'E2EFDA',            
                    'DIAN - DUPLICADO': 'FF6B6B',                     
                    'CONT - DUPLICADO': 'FF6B6B',
                    'DIAN - Sin Pareja en Contabilidad ni Devoluciones': 'FFF2CC',
                    'DEVOL - Sin Pareja en DIAN': 'E2EFDA',
                    'DEVOL - Sin Num.Ext': 'F2DCDB',
                }

                col_fecha_idx = df_dian_vs_cont.columns.get_loc('Fecha') + 1 if 'Fecha' in df_dian_vs_cont.columns else None
                col_motivo_idx = df_dian_vs_cont.columns.get_loc('Novedad') + 1 if 'Novedad' in df_dian_vs_cont.columns else None
                col_doc = df_dian_vs_cont.columns.get_loc('Tipo de documento') + 1 if 'Tipo de documento' in df_dian_vs_cont.columns else None    
                
                col_base_idx_dian = df_dian_vs_cont.columns.get_loc('BASE') + 1 if 'BASE' in df_dian_vs_cont.columns else None
                col_iva_idx_dian = df_dian_vs_cont.columns.get_loc('IVA') + 1 if 'IVA' in df_dian_vs_cont.columns else None
                col_num_ext_idx = df_dian_vs_cont.columns.get_loc('Num.Ext') + 1 if 'Num.Ext' in df_dian_vs_cont.columns else None
                col_nit_idx_dian = df_dian_vs_cont.columns.get_loc('NIT Emisor') + 1 if 'NIT Emisor' in df_dian_vs_cont.columns else None
                col_base2_idx_dian = df_dian_vs_cont.columns.get_loc('BASE_2') + 1 if 'BASE_2' in df_dian_vs_cont.columns else None

                for row_idx in range(2, max_row + 1):
                    es_vacio = True

# Definir fuente roja reutilizable
                font_red = Font(name="Arial", size=10, color="FF0000")
                font_default = Font(name="Arial", size=10)

                for row_idx in range(2, max_row + 1):
                    es_vacio = True

                    if col_fecha_idx:
                        val_fecha = sheet_dian_vs_cont.cell(row=row_idx, column=col_fecha_idx).value
                        if val_fecha is not None and str(val_fecha).strip() != '':
                            es_vacio = False

                    motivo = sheet_dian_vs_cont.cell(row=row_idx, column=col_motivo_idx).value if col_motivo_idx else ''
                    
                    color_llenado_hex = "627FF0" 
                    color_vacio_hex = "A6B7F5"   
                    
                    es_incompleta = parejas_incompletas[row_idx - 2] if (row_idx - 2) < len(parejas_incompletas) else True
                    
                    if not es_incompleta and (row_idx - 2) < len(prefix_doc_group):
                        prefix = prefix_doc_group[row_idx - 2]
                        if prefix == 'FC':
                            color_llenado_hex, color_vacio_hex = 'B8E3BD', 'DAEDDC'
                        elif prefix == 'GC':
                            color_llenado_hex, color_vacio_hex = 'ADB3DB', 'D5D8F0'
                        elif prefix == 'DS':
                            color_llenado_hex, color_vacio_hex = 'E6BE9A', 'F2D7BF'
                        elif prefix == 'DC':
                            color_llenado_hex, color_vacio_hex = 'cddbb4', 'e7f5ce'
                        elif prefix == 'CG':
                            color_llenado_hex, color_vacio_hex = 'f5d6e7', 'f9e6f2'
                            
                    current_fill_llenado = PatternFill(start_color=color_llenado_hex, end_color=color_llenado_hex, fill_type="solid")
                    current_fill_vacio = PatternFill(start_color=color_vacio_hex, end_color=color_vacio_hex, fill_type="solid")

                    if motivo and str(motivo).strip() != '':
                        motivo_str = str(motivo).strip()
                        if 'DUPLICADO' in motivo_str:
                            color_hex = DUPLICADO_FILL_COLOR
                        elif 'Gasto Personal' in motivo_str:
                            color_hex = 'FFCCFF'
                        elif 'No coincide tercero' in motivo_str:
                            color_hex = color_llenado_hex
                        else:
                            color_hex = COLOR_MOTIVOS.get(motivo_str, 'FFC7CE') 
                        fill_base = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                    else:
                        fill_base = current_fill_vacio if es_vacio else current_fill_llenado

                    # --- NUEVA EVALUACIÓN: FUENTE ROJA PARA NOTAS DE CRÉDITO ---
                    tipo_doc_val = sheet_dian_vs_cont.cell(row=row_idx, column=col_doc).value if col_doc else ''
                    es_nota_credito = 'nota de crédito electrónica' in str(tipo_doc_val).strip().lower()
                    current_font = font_red if es_nota_credito else font_default

                    for col_idx in range(1, max_col + 1):
                        cell = sheet_dian_vs_cont.cell(row=row_idx, column=col_idx)
                        cell.fill = fill_base
                        cell.font = current_font  # <--- Aplica la fuente según corresponda
                        cell.border = thin_border
                    es_dif_base = dif_base[row_idx - 2] if (row_idx - 2) < len(dif_base) else False
                    es_dif_iva = dif_iva[row_idx - 2] if (row_idx - 2) < len(dif_iva) else False

                    if es_dif_base and col_base_idx_dian:
                        sheet_dian_vs_cont.cell(row=row_idx, column=col_base_idx_dian).fill = fill_red_alert

                    if es_dif_iva and col_iva_idx_dian:
                        sheet_dian_vs_cont.cell(row=row_idx, column=col_iva_idx_dian).fill = fill_red_alert

                    # --- NUEVO: APLICAR FADC46 DIRECTAMENTE SOBRE LAS CELDAS ---
                    box_color = box_color_list[row_idx - 2] if (row_idx - 2) < len(box_color_list) else ''
                    if box_color == 'YELLOW':
                        fadc46_fill = PatternFill(start_color='FADC46', end_color='FADC46', fill_type='solid')
                        if col_base_idx_dian:
                            sheet_dian_vs_cont.cell(row=row_idx, column=col_base_idx_dian).fill = fadc46_fill
                        if col_iva_idx_dian:
                            sheet_dian_vs_cont.cell(row=row_idx, column=col_iva_idx_dian).fill = fadc46_fill

                    if col_num_ext_idx and row_idx - 2 < len(tiene_caracter_especial):
                        if tiene_caracter_especial[row_idx - 2]:
                            sheet_dian_vs_cont.cell(row=row_idx, column=col_num_ext_idx).fill = fill_red_alert
                            
                    if col_nit_idx_dian and row_idx - 2 < len(dif_nit):
                        if dif_nit[row_idx - 2]:
                            sheet_dian_vs_cont.cell(row=row_idx, column=col_nit_idx_dian).fill = fill_red_alert                    

# =========================================================================
                # DIBUJO DEL RECUADRO Y RELLENO PARA LAS 4 CELDAS (BASE e IVA)
                # =========================================================================
                if col_base_idx_dian and col_iva_idx_dian:
                    col_start = min(col_base_idx_dian, col_iva_idx_dian)
                    col_end = max(col_base_idx_dian, col_iva_idx_dian)
                    
                    black_side = Side(style='medium', color='000000')
                    fadc46_side = Side(style='medium', color='FADC46')
                    
                    i = 0
                    while i < len(box_color_list):
                        color_flag = box_color_list[i]
                        if color_flag in ['BLACK', 'FADC46'] and i + 1 < len(box_color_list) and box_color_list[i+1] == color_flag:
                            thick_side = black_side if color_flag == 'BLACK' else fadc46_side
                            
                            row_1 = i + 2
                            row_2 = i + 3
                            
                            for c in range(col_start, col_end + 1):
                                c_top = sheet_dian_vs_cont.cell(row=row_1, column=c)
                                c_bot = sheet_dian_vs_cont.cell(row=row_2, column=c)
                                
                                t_top = thick_side
                                t_bot = thin_border.bottom
                                t_left = thick_side if c == col_start else thin_border.left
                                t_right = thick_side if c == col_end else thin_border.right
                                c_top.border = Border(top=t_top, bottom=t_bot, left=t_left, right=t_right)
                                
                                b_top = thin_border.top
                                b_bot = thick_side
                                b_left = thick_side if c == col_start else thin_border.left
                                b_right = thick_side if c == col_end else thin_border.right
                                c_bot.border = Border(top=b_top, bottom=b_bot, left=b_left, right=b_right)
                                
                            i += 2
                        else:
                            i += 1

                mask_conciliados = ~np.array(parejas_incompletas)
                sum_base = pd.to_numeric(df_dian_vs_cont.loc[mask_conciliados, 'BASE'], errors='coerce').fillna(0).sum() if 'BASE' in df_dian_vs_cont.columns else 0
                sum_iva = pd.to_numeric(df_dian_vs_cont.loc[mask_conciliados, 'IVA'], errors='coerce').fillna(0).sum() if 'IVA' in df_dian_vs_cont.columns else 0
                sum_base2 = pd.to_numeric(df_dian_vs_cont.loc[mask_conciliados, 'BASE_2'], errors='coerce').fillna(0).sum() if 'BASE_2' in df_dian_vs_cont.columns else 0

                boundary_added = False

                for row_idx in range(max_row, 2, -1):
                    idx_curr = row_idx - 2
                    idx_prev = row_idx - 3

                    curr_incompleta = parejas_incompletas[idx_curr]
                    prev_incompleta = parejas_incompletas[idx_prev]

                    if curr_incompleta == True and prev_incompleta == False:
                        sheet_dian_vs_cont.insert_rows(row_idx, amount=5)
                        boundary_added = True

                        row_sum = row_idx + 2 

                        if col_base_idx_dian and col_base_idx_dian > 1:
                            lbl = sheet_dian_vs_cont.cell(row=row_sum, column=col_base_idx_dian - 1)
                            lbl.value = "TOTAL CONCILIADOS:"
                            lbl.font = Font(bold=True)
                            lbl.alignment = Alignment(horizontal='right')

                        for val, col_idx in [(sum_base, col_base_idx_dian), (sum_iva, col_iva_idx_dian), (sum_base2, col_base2_idx_dian)]:
                            if col_idx:
                                c = sheet_dian_vs_cont.cell(row=row_sum, column=col_idx)
                                c.value = val
                                c.number_format = '#,##0.00'
                                c.font = Font(bold=True)

                    elif curr_incompleta == False and prev_incompleta == False:
                        curr_emisor = nombre_emisor_group[idx_curr]
                        prev_emisor = nombre_emisor_group[idx_prev]
                        
                        curr_prefix = prefix_doc_group[idx_curr]
                        prev_prefix = prefix_doc_group[idx_prev]

                        if curr_emisor != prev_emisor or curr_prefix != prev_prefix:
                            sheet_dian_vs_cont.insert_rows(row_idx, amount=1)

                if not boundary_added and len(parejas_incompletas) > 0 and not all(parejas_incompletas):
                    end_row = sheet_dian_vs_cont.max_row + 1
                    sheet_dian_vs_cont.insert_rows(end_row, amount=5)
                    row_sum = end_row + 2

                    if col_base_idx_dian and col_base_idx_dian > 1:
                        lbl = sheet_dian_vs_cont.cell(row=row_sum, column=col_base_idx_dian - 1)
                        lbl.value = "TOTAL CONCILIADOS:"
                        lbl.font = Font(bold=True)
                        lbl.alignment = Alignment(horizontal='right')

                    for val, col_idx in [(sum_base, col_base_idx_dian), (sum_iva, col_iva_idx_dian), (sum_base2, col_base2_idx_dian)]:
                        if col_idx:
                            c = sheet_dian_vs_cont.cell(row=row_sum, column=col_idx)
                            c.value = val
                            c.number_format = '#,##0.00'
                            c.font = Font(bold=True)

                for idx, col_name in enumerate(df_dian_vs_cont.columns):
                    col_letter = get_column_letter(idx + 1)
                    col_str = str(col_name).strip().upper()

                    if col_str == "CUFE/CUDE":
                        sheet_dian_vs_cont.column_dimensions[col_letter].width = 6   # ~45px
                    elif col_str == "NOMBRE EMISOR":
                        sheet_dian_vs_cont.column_dimensions[col_letter].width = 26  # ~187px
                    else:
                        max_len_header = len(str(col_name))
                        max_len_data = df_dian_vs_cont[col_name].astype(str).str.len().max()
                        max_len_data = 0 if pd.isna(max_len_data) else int(max_len_data)
                        ancho_final = max(max_len_data, max_len_header, 12) + 3
                        sheet_dian_vs_cont.column_dimensions[col_letter].width = ancho_final

                for sheetname in OUTPUT_SHEETS_TO_HIDE:
                    if sheetname in wb.sheetnames:
                        wb[sheetname].sheet_state = 'hidden'

                if not any(ws.sheet_state == 'visible' for ws in wb.worksheets):
                    wb.worksheets[0].sheet_state = 'visible'

        except (ErrorUsuario, ErrorSistema):
            raise
        except Exception as e:
            raise ErrorSistema(f"Error escribiendo resultados en el archivo Excel: {e}") from e