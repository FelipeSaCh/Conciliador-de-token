"""Carga y normalización de las fuentes: Token (DIAN), AUD-COMP (contabilidad) y devoluciones."""
import numpy as np
import pandas as pd

from core.config import COLUMNAS_DIAN_VS_CONT, IMPUESTOS_ADICIONALES, ORDEN
from core.errors import ErrorUsuario
from shared.dataframes import a_numero, buscar_columna, columna_o_vacio, contiene, quitar_filas_vacias, sumar_columnas
from shared.texto import limpiar_codigo, mascara_duplicados, normalizar_num_ext, tiene_caracter_especial

PRIORIDAD_DIAN, PRIORIDAD_CONT, PRIORIDAD_DEVOL = 1, 2, 3


def completar_columnas_token(df):
    if 'Prefijo' not in df.columns:
        df['Prefijo'] = ''

    if 'Num.Ext' not in df.columns:
        col_folio = buscar_columna(df.columns, 'FOLIO')
        if not col_folio:
            raise ErrorUsuario("La hoja de Token no contiene 'Num.Ext' ni la columna 'Folio' para construirlo.")
        df['Num.Ext'] = limpiar_codigo(df['Prefijo'], strip=False) + limpiar_codigo(df[col_folio], strip=False)

    if 'BASE' not in df.columns:
        if 'Total' in df.columns and 'IVA' in df.columns:
            total, iva = a_numero(df['Total']), a_numero(df['IVA'])
            df['BASE'] = np.where(total == 0, 0, total - iva)
        else:
            df['BASE'] = 0

    if 'CONCEPTO' not in df.columns:
        df['CONCEPTO'] = ''

    for destino, origenes in IMPUESTOS_ADICIONALES.items():
        objetivos = {x.upper() for x in origenes}
        df[destino] = sumar_columnas(df, [c for c in df.columns if str(c).strip().upper() in objetivos])
    return df


def mapas_impuestos(df):
    """{columna: {Num.Ext normalizado: total}} para OTROS IMPUESTOS y RETENCIONES."""
    claves = normalizar_num_ext(df['Num.Ext'])
    mapas = {}
    for col in IMPUESTOS_ADICIONALES:
        mapa = df.groupby(claves)[col].sum().to_dict()
        mapa.pop('', None)
        mapas[col] = mapa
    return mapas


def es_personal(df):
    return df['CONCEPTO'].astype(str).str.strip().str.upper().str.contains('PERSONAL', na=False)


def marcar_novedades_token(df, filtro_grupo):
    """Marca personales, aplica el filtro de grupo, elimina Application Response y marca duplicados / sin Num.Ext.

    Devuelve (df, es_personales, mask_ignorar, cantidad_duplicados).
    """
    df['Novedad'] = ''
    personales = es_personal(df)
    df.loc[personales, 'Novedad'] = 'DIAN - Gasto Personal: ' + df.loc[personales, 'CONCEPTO'].astype(str)

    col_grupo = buscar_columna(df.columns, 'grupo')
    col_doc = buscar_columna(df.columns, 'tipo de documento')
    sin_novedad = df['Novedad'] == ''
    eliminar = pd.Series(False, index=df.index)

    if col_grupo:
        if filtro_grupo == "Solo Recibidos":
            eliminar = contiene(df[col_grupo], 'emitido')
            if col_doc:
                eliminar &= ~contiene(df[col_doc], 'soporte')
        elif filtro_grupo == "Solo Emitidos":
            eliminar = contiene(df[col_grupo], 'recibido')

    if col_doc:
        eliminar |= contiene(df[col_doc], 'Application response')

    df = df[~(eliminar & sin_novedad)].reset_index(drop=True)
    personales = es_personal(df)

    duplicados = mascara_duplicados(df['Num.Ext'])
    df.loc[duplicados & (df['Novedad'] == ''), 'Novedad'] = 'DIAN - DUPLICADO'

    mask_ignorar = contiene(df[col_doc], 'application response') if col_doc else pd.Series(False, index=df.index)

    sin_num = df['Num.Ext'].fillna('').astype(str).str.strip() == ''
    df.loc[sin_num & (df['Novedad'] == ''), 'Novedad'] = 'DIAN - Sin Num.Ext'
    return df, personales, mask_ignorar, int(duplicados.sum())


def construir_resultado(df, personales, mask_ignorar):
    df_proc = df[~mask_ignorar]
    col_concepto = 'TIPO-DETALLE' if 'TIPO-DETALLE' in df_proc.columns else 'CONCEPTO'
    res = df_proc[~personales[~mask_ignorar]][['Prefijo', 'BASE', 'Num.Ext', col_concepto]]
    return res.rename(columns={col_concepto: 'TIPO-DETALLE'})


def construir_auditoria(df):
    for col in ORDEN:
        if col not in df.columns:
            df[col] = ''
    return df[ORDEN + ['Novedad', *IMPUESTOS_ADICIONALES]].copy()


def limpiar_hoja_auditoria(df):
    """Quita filas fantasma evaluando las 6 primeras columnas de AUD-COMP / devoluciones."""
    return quitar_filas_vacias(df, df.columns[:6]).copy() if len(df.columns) >= 6 else df


def procesar_seriales(df_aud, seriales_iva, seriales_base, seriales_base2, mapas_impuestos=None):
    """Suma por fila las columnas (seriales) de cada categoría. Contabilidad siempre en negativo."""
    res = pd.DataFrame(index=df_aud.index)
    res['Num.Ext'] = df_aud['Num.Ext'].fillna('').astype(str) if 'Num.Ext' in df_aud.columns else ''

    for destino, seriales in (('IVA', seriales_iva), ('BASE', seriales_base), ('BASE_2', seriales_base2)):
        cols = [c for c in (seriales or []) if c in df_aud.columns]
        res[destino] = -sumar_columnas(df_aud, cols).abs() if cols else 0

    claves = normalizar_num_ext(res['Num.Ext'])
    for col in IMPUESTOS_ADICIONALES:
        res[col] = claves.map(mapas_impuestos[col]).fillna(0) if mapas_impuestos else 0
    return res[['Num.Ext', 'BASE', 'BASE_2', 'IVA', *IMPUESTOS_ADICIONALES]]


def preparar_dian(df_auditoria):
    df = df_auditoria.copy()
    df['Fecha'] = ''
    df['Prioridad_Fila'] = PRIORIDAD_DIAN
    for col in ('BASE', 'IVA', 'BASE_2'):
        if col in df.columns:
            df[col] = a_numero(df[col]).abs()
    df['Num.Ext_Original'] = df['Num.Ext'] if 'Num.Ext' in df.columns else ''
    df['Detalle'] = ''
    for col in COLUMNAS_DIAN_VS_CONT:
        if col not in df.columns:
            df[col] = ''
    df['Num.Ext_Clean'] = normalizar_num_ext(df['Num.Ext'])
    df['Tiene_Caracter_Especial'] = tiene_caracter_especial(df['Num.Ext_Original'])
    return df


def preparar_contable(df_aud, df_res, prioridad, origen):
    """Normaliza AUD-COMP (origen 'CONT') o devoluciones (origen 'DEVOL') al esquema DIAN VS CONT."""
    df = pd.DataFrame()
    df['Tipo de documento'] = columna_o_vacio(df_aud, 'Tipo')
    df['CUFE/CUDE'] = limpiar_codigo(df_aud['Número']) if 'Número' in df_aud.columns else ''
    col_detalle = buscar_columna(df_aud.columns, 'DETALLE')
    df['Detalle'] = df_aud[col_detalle] if col_detalle else ''

    df['Num.Ext'] = columna_o_vacio(df_aud, 'Num.Ext')
    df['Num.Ext_Original'] = df['Num.Ext']
    df['Num.Ext_Clean'] = normalizar_num_ext(df['Num.Ext'])
    df['Tiene_Caracter_Especial'] = tiene_caracter_especial(df['Num.Ext_Original'])

    duplicados = mascara_duplicados(df_aud['Num.Ext']) if 'Num.Ext' in df_aud.columns else False
    df['Novedad'] = np.where(
        duplicados, f'{origen} - DUPLICADO',
        np.where(df['Num.Ext_Clean'] == '', f'{origen} - Sin Num.Ext', '')
    )

    df['Fecha'] = columna_o_vacio(df_aud, 'Fecha')
    df['NIT Emisor'] = columna_o_vacio(df_aud, 'Nit/C.C.')
    df['Nombre Emisor'] = columna_o_vacio(df_aud, 'Tercero')
    for col in ('BASE', 'BASE_2', 'IVA', *IMPUESTOS_ADICIONALES):
        df[col] = -a_numero(df_res[col]).abs()
    df['Prioridad_Fila'] = prioridad

    for col in COLUMNAS_DIAN_VS_CONT:
        if col not in df.columns:
            df[col] = ''
    return df
