"""Emparejamiento DIAN vs Contabilidad vs Devoluciones y cálculo de diferencias."""
import difflib
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from core.config import COLUMNAS_DIAN_VS_CONT, UMBRAL_TOLERANCIA
from shared.dataframes import a_numero
from shared.texto import limpiar_codigo, normalizar_nit
from services.conciliacion.preparacion import (
    PRIORIDAD_CONT, PRIORIDAD_DEVOL, PRIORIDAD_DIAN, preparar_contable, preparar_dian, procesar_seriales,
)

SIMILITUD_MINIMA = 0.40


@dataclass
class ResultadoConciliacion:
    df_completo: pd.DataFrame
    df_dian_vs_cont: pd.DataFrame
    df_res_dc: pd.DataFrame = None
    incompletas: list = field(default_factory=list)
    dif_base: list = field(default_factory=list)
    dif_iva: list = field(default_factory=list)
    caracter_especial: list = field(default_factory=list)
    dif_nit: list = field(default_factory=list)
    nombre_emisor_group: list = field(default_factory=list)
    prefix_doc_group: list = field(default_factory=list)
    box_color: list = field(default_factory=list)


def similitud_folios(a, b):
    sim = difflib.SequenceMatcher(None, a, b).ratio()
    sin_ceros_a = a.replace('0', '')
    if sin_ceros_a == b.replace('0', '') and sin_ceros_a != '':
        sim = max(sim, 0.95)
    if (a in b or b in a) and len(a) > 2 and len(b) > 2:
        sim = max(sim, 0.90)
    return sim


def emparejar_codicioso(candidatos):
    """candidatos: [(similitud, idx_a, idx_b)]. Empareja de mayor a menor similitud sin repetir índices."""
    usados_a, usados_b, pares = set(), set(), []
    for _, a, b in sorted(candidatos, key=lambda x: x[0], reverse=True):
        if a not in usados_a and b not in usados_b:
            usados_a.add(a)
            usados_b.add(b)
            pares.append((a, b))
    return pares


def _recontar_parejas(df, mask):
    df.loc[mask, 'Conteo_Pareja'] = df[mask].groupby('Num.Ext_Clean')['Num.Ext_Clean'].transform('count')


def _agregar_novedad(df, mask, texto):
    actual = df.loc[mask, 'Novedad']
    df.loc[mask, 'Novedad'] = np.where(actual == '', texto, actual + f' / {texto}')


def _emparejar_flexible(df, mask_huerfano):
    """Huérfanos DIAN/CONT del mismo NIT: se unen por similitud de folio si cuadra la BASE o el IVA."""
    df['Fue_Fuzzy_Match'] = False
    huerfanos = df[mask_huerfano]
    if huerfanos.empty:
        return

    nit = limpiar_codigo(df['NIT Emisor'])
    base = a_numero(df['BASE'])
    iva = a_numero(df['IVA'])
    folio = df['Num.Ext_Clean']

    for _, grupo in huerfanos.groupby(nit[mask_huerfano]):
        dian = grupo.index[grupo['Prioridad_Fila'] == PRIORIDAD_DIAN]
        cont = grupo.index[grupo['Prioridad_Fila'] == PRIORIDAD_CONT]
        if not len(dian) or not len(cont):
            continue
        candidatos = []
        for d in dian:
            for c in cont:
                # DIAN es positivo y CONT negativo: se suman
                if abs(base[d] + base[c]) > UMBRAL_TOLERANCIA and abs(iva[d] + iva[c]) > UMBRAL_TOLERANCIA:
                    continue
                sim = similitud_folios(folio[d], folio[c])
                if sim > SIMILITUD_MINIMA:
                    candidatos.append((sim, d, c))
        for d, c in emparejar_codicioso(candidatos):
            clave = f"{df.at[d, 'Num.Ext_Clean']}_FUZZY_{d}"
            df.loc[[d, c], 'Num.Ext_Clean'] = clave
            df.loc[[d, c], 'Fue_Fuzzy_Match'] = True


def conciliar_con_devoluciones(df_unificado, df_aud_dc, seriales_iva_dc, seriales_base_dc, seriales_base2_dc, mapa_otros_imp=None):
    df_res_dc = procesar_seriales(df_aud_dc, seriales_iva_dc, seriales_base_dc, seriales_base2_dc, mapa_otros_imp)
    df_dc = preparar_contable(df_aud_dc, df_res_dc, PRIORIDAD_DEVOL, 'DEVOL')
    df_dc['Conteo_Pareja'] = 1
    df_dc['Fue_Fuzzy_Match'] = False

    dian_idx = df_unificado.index[
        (df_unificado['Prioridad_Fila'] == PRIORIDAD_DIAN) & (df_unificado['Novedad'] == 'DIAN - Sin Pareja en Contabilidad')
    ].tolist()
    dc_matchable = df_dc['Novedad'] == ''
    dc_idx = df_dc.index[dc_matchable].tolist()

    emp_dian, emp_dc = set(), set()

    # 1. Coincidencia exacta de Num.Ext
    dc_por_clave = {}
    for i in dc_idx:
        clave = df_dc.at[i, 'Num.Ext_Clean']
        if clave:
            dc_por_clave.setdefault(clave, []).append(i)
    for d in dian_idx:
        libres = [c for c in dc_por_clave.get(df_unificado.at[d, 'Num.Ext_Clean'], []) if c not in emp_dc]
        if libres:
            emp_dian.add(d)
            emp_dc.add(libres[0])
            df_unificado.at[d, 'Novedad'] = ''
            df_dc.at[libres[0], 'Novedad'] = ''

    # 2. Coincidencia flexible (mismo NIT, base dentro de tolerancia, folio similar)
    dian_rest = [i for i in dian_idx if i not in emp_dian]
    dc_rest = [i for i in dc_idx if i not in emp_dc]
    if dian_rest and dc_rest:
        nit_simple = lambda v: str(v).replace('.0', '').strip()
        candidatos = []
        for d in dian_rest:
            nit_d = nit_simple(df_unificado.at[d, 'NIT Emisor'])
            if not nit_d:
                continue
            base_d = pd.to_numeric(df_unificado.at[d, 'BASE'], errors='coerce')
            base_d = 0.0 if pd.isna(base_d) else base_d
            folio_d = df_unificado.at[d, 'Num.Ext_Clean']
            for c in dc_rest:
                if nit_d != nit_simple(df_dc.at[c, 'NIT Emisor']):
                    continue
                base_c = pd.to_numeric(df_dc.at[c, 'BASE'], errors='coerce')
                base_c = 0.0 if pd.isna(base_c) else base_c
                if abs(base_d - base_c) > UMBRAL_TOLERANCIA:
                    continue
                sim = similitud_folios(folio_d, df_dc.at[c, 'Num.Ext_Clean'])
                if sim > SIMILITUD_MINIMA:
                    candidatos.append((sim, d, c))
        for d, c in emparejar_codicioso(candidatos):
            clave = f"{df_unificado.at[d, 'Num.Ext_Clean']}_DEVOL_{d}"
            df_unificado.at[d, 'Num.Ext_Clean'] = clave
            df_dc.at[c, 'Num.Ext_Clean'] = clave
            df_unificado.at[d, 'Novedad'] = ''
            df_dc.at[c, 'Novedad'] = ''
            df_unificado.at[d, 'Fue_Fuzzy_Match'] = True
            df_dc.at[c, 'Fue_Fuzzy_Match'] = True
            emp_dian.add(d)
            emp_dc.add(c)

    for d in dian_idx:
        if d not in emp_dian:
            df_unificado.at[d, 'Novedad'] = 'DIAN - Sin Pareja en Contabilidad ni Devoluciones'
    df_dc.loc[dc_matchable & ~df_dc.index.isin(emp_dc), 'Novedad'] = 'DEVOL - Sin Pareja en DIAN'

    df_unificado = pd.concat([df_unificado, df_dc], ignore_index=True)
    _recontar_parejas(df_unificado, df_unificado['Novedad'] == '')
    return df_unificado, df_res_dc


def _aplicar_autorretenedores(df, df_auto):
    if df_auto.empty:
        return df
    col_nit = next((c for c in df_auto.columns if 'NIT' in c), None)
    col_com = next((c for c in df_auto.columns if 'COMENT' in c), None)
    if not (col_nit and col_com):
        return df

    auto = df_auto[[col_nit, col_com]].copy()
    auto['NIT_Clean'] = auto[col_nit].map(normalizar_nit)
    auto = auto[auto['NIT_Clean'] != ''].drop_duplicates(subset=['NIT_Clean'])

    df = df.merge(auto[['NIT_Clean', col_com]], left_on='NIT_Emisor_Group', right_on='NIT_Clean', how='left')
    cond = (df['Prioridad_Fila'] == PRIORIDAD_CONT) & df[col_com].notna()
    df['NIT Receptor'] = np.where(cond, df[col_com], df['NIT Receptor'])
    return df.drop(columns=['NIT_Clean', col_com], errors='ignore')


def _marcar_diferencias_valores(df, mask_pairs):
    base = a_numero(df['BASE'])
    iva = a_numero(df['IVA'])
    claves = df.loc[mask_pairs, 'Num.Ext_Clean']
    suma_base = base[mask_pairs].groupby(claves).transform('sum')
    suma_iva = iva[mask_pairs].groupby(claves).transform('sum')
    suma_total = (base + iva)[mask_pairs].groupby(claves).transform('sum')

    df['Diff_BASE'] = False
    df['Diff_IVA'] = False
    df['Box_Color'] = ''
    df.loc[mask_pairs, 'Diff_BASE'] = suma_base.abs() > UMBRAL_TOLERANCIA
    df.loc[mask_pairs, 'Diff_IVA'] = suma_iva.abs() > UMBRAL_TOLERANCIA
    df.loc[mask_pairs, 'Box_Color'] = np.where(suma_total.abs() <= UMBRAL_TOLERANCIA, 'BLACK', 'YELLOW')


def _marcar_diferencias_nit(df, mask_pairs):
    """La pareja es válida si el NIT contable coincide con el emisor o el receptor DIAN."""
    df['Diff_NIT'] = False
    receptor = df['NIT Receptor'].map(normalizar_nit) if 'NIT Receptor' in df.columns else pd.Series('', index=df.index)

    for _, grupo in df[mask_pairs].groupby('Num.Ext_Clean'):
        idx_dian = grupo.index[grupo['Prioridad_Fila'] == PRIORIDAD_DIAN]
        idx_cont = grupo.index[grupo['Prioridad_Fila'].isin([PRIORIDAD_CONT, PRIORIDAD_DEVOL])]
        if not len(idx_dian) or not len(idx_cont):
            continue
        validos = set(grupo.loc[idx_dian, 'NIT_Emisor_Clean']) | {r for r in receptor[idx_dian] if r != ''}
        invalidos = [c for c in idx_cont if grupo.at[c, 'NIT_Emisor_Clean'] not in validos]
        if invalidos:
            df.loc[invalidos + list(idx_dian), 'Diff_NIT'] = True

    _agregar_novedad(df, df['Diff_NIT'], 'No coincide tercero')


def unificar_dian_vs_cont(df_auditoria, df_aud_comp, df_res_auditoria, df_autorretenedores,
                          df_aud_dc=None, seriales_iva_dc=None, seriales_base_dc=None, seriales_base2_dc=None,
                          mapa_otros_imp=None):
    df_dian = preparar_dian(df_auditoria)
    df_cont = preparar_contable(df_aud_comp, df_res_auditoria, PRIORIDAD_CONT, 'CONT')
    df = pd.concat([df_dian, df_cont], ignore_index=True)

    col_nombre = 'Nombre Emisor' if 'Nombre Emisor' in df.columns else 'TERCERO'
    mask_totales = df[col_nombre].fillna('').astype(str).str.upper().str.contains('TOTALES')
    mask_ceros = df['Num.Ext_Clean'].fillna('').astype(str).str.contains('0000000000')
    df = df[~(mask_totales | mask_ceros)].reset_index(drop=True)

    # Emparejamiento exacto + flexible solo para registros aptos
    df['Conteo_Pareja'] = 1
    mask_matchable = (df['Num.Ext_Clean'] != '') & (df['Novedad'] == '')
    _recontar_parejas(df, mask_matchable)
    _emparejar_flexible(df, mask_matchable & (df['Conteo_Pareja'] == 1))
    _recontar_parejas(df, mask_matchable)

    mask_single = mask_matchable & (df['Conteo_Pareja'] == 1)
    df.loc[mask_single & (df['Prioridad_Fila'] == PRIORIDAD_DIAN), 'Novedad'] = 'DIAN - Sin Pareja en Contabilidad'
    df.loc[mask_single & (df['Prioridad_Fila'] == PRIORIDAD_CONT), 'Novedad'] = 'CONT - Sin Pareja en DIAN'

    df_res_dc = None
    if df_aud_dc is not None and not df_aud_dc.empty:
        df, df_res_dc = conciliar_con_devoluciones(
            df, df_aud_dc, seriales_iva_dc, seriales_base_dc, seriales_base2_dc, mapa_otros_imp
        )

    df['NIT_Emisor_Clean'] = df['NIT Emisor'].map(normalizar_nit)
    df['NIT_Emisor_Group'] = df.groupby('Num.Ext_Clean')['NIT_Emisor_Clean'].transform('max')
    df = _aplicar_autorretenedores(df, df_autorretenedores)

    df['Es_Incompleta'] = df['Novedad'] != ''
    df['Nombre_Emisor_Group'] = (
        df['Nombre Emisor'].fillna('').astype(str).str.strip().str.upper().groupby(df['Num.Ext_Clean']).transform('max')
    )
    es_contable = df['Prioridad_Fila'].isin([PRIORIDAD_CONT, PRIORIDAD_DEVOL])
    prefijo = pd.Series(
        np.where(es_contable, df['Tipo de documento'].astype(str).str.strip().str.upper().str[:2], ''), index=df.index
    )
    df['Prefix_Doc_Group'] = prefijo.groupby(df['Num.Ext_Clean']).transform('max')

    completa = ~df['Es_Incompleta']
    df['Sort_Matched_Tipo'] = np.where(completa, df['Prefix_Doc_Group'], '')
    df['Sort_Matched_Emisor'] = np.where(completa, df['Nombre_Emisor_Group'], '')
    df['Sort_Unmatched_Detalle'] = np.where(~completa, df['Novedad'], '')
    df = df.sort_values(by=['Es_Incompleta', 'Sort_Matched_Tipo', 'Sort_Matched_Emisor', 'Sort_Unmatched_Detalle',
                            'Nombre_Emisor_Group', 'Num.Ext_Clean', 'Prioridad_Fila'])

    mask_pairs = df['Conteo_Pareja'] >= 2
    _marcar_diferencias_valores(df, mask_pairs)
    _marcar_diferencias_nit(df, mask_pairs)

    fuzzy = df['Fue_Fuzzy_Match'] == True  # noqa: E712 (puede contener NaN)
    _agregar_novedad(df, fuzzy, 'Dif. formato de factura')
    df['Tiene_Caracter_Especial'] = df['Tiene_Caracter_Especial'] | fuzzy

    columnas_finales = list(COLUMNAS_DIAN_VS_CONT) + ['OTROS IMPUESTOS', 'Novedad', 'Detalle']
    return ResultadoConciliacion(
        df_completo=df,
        df_dian_vs_cont=df[columnas_finales],
        df_res_dc=df_res_dc,
        incompletas=df['Es_Incompleta'].tolist(),
        dif_base=df['Diff_BASE'].tolist(),
        dif_iva=df['Diff_IVA'].tolist(),
        caracter_especial=df['Tiene_Caracter_Especial'].tolist(),
        dif_nit=df['Diff_NIT'].tolist(),
        nombre_emisor_group=df['Nombre_Emisor_Group'].tolist(),
        prefix_doc_group=df['Prefix_Doc_Group'].tolist(),
        box_color=df['Box_Color'].tolist(),
    )
