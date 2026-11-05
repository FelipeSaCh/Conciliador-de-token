"""Alertas tributarias: retenciones vs. perfil de la empresa y autorretenedores, IVA, estado DIAN, periodo y soporte."""
import re
from datetime import date

import numpy as np
import pandas as pd

from core.config import (
    COMENTARIOS_NO_SUJETOS, CONCEPTOS_ESTRICTOS, PATRON_IMPUESTO_EN_COSTO,
    TARIFA_MAXIMA_IVA, TARIFA_MAXIMA_RETENCION, TARIFA_RETEIVA, TIPOS_CONTRIBUYENTE, TOLERANCIA_RELATIVA,
    TOLERANCIA_TARIFA_RETENCION, UMBRAL_TOLERANCIA,
)
from data import parametros_repo
from shared.cuentas import nombre_cuenta
from shared.dataframes import a_numero, buscar_columna, contiene, sumar_columnas
from shared.texto import normalizar_nit

COL_ALERTAS = 'Alertas'
NO_RESPONSABLE_IVA = "No responsable de IVA"
PERIODOS_IVA = {"Bimestral": (2, 3), "Cuatrimestral": (4, 1)}  # (meses por periodo, periodos siguientes permitidos)
PATRON_NOTA_CREDITO = r'cr[eé]dito'


def agregar_alerta(serie, mask, texto):
    """texto: str o Series (un texto por fila)."""
    mask = mask.reindex(serie.index, fill_value=False).fillna(False).astype(bool)
    t = texto[mask].astype(str) if isinstance(texto, pd.Series) else texto
    serie.loc[mask] = np.where(serie[mask] == '', t, serie[mask] + ' / ' + t)
    return serie


def _pesos(valor):
    return f"${valor:,.0f}".replace(',', '.')


def _porcentaje(valor):
    return f"{valor * 100:.2f} %".replace('.', ',')


def diccionario_autorretenedores(df_auto):
    cols = df_auto.columns
    col_nit = next((c for c in cols if 'NIT' in str(c).upper()), None)
    col_com = next((c for c in cols if 'COMENT' in str(c).upper()), None) or \
        next((c for c in cols if 'NOMBRE' in str(c).upper()), None)
    if not (col_nit and col_com):
        return {}
    dic = {}
    for nit, comentario in zip(df_auto[col_nit], df_auto[col_com]):
        nit = normalizar_nit(nit)
        if nit:
            dic[nit] = str(comentario).strip() or 'AUTORRETENEDOR'
    return dic


def agentes_retencion(perfil):
    """(agente de retención en la fuente, agente de retención de IVA) según el tipo de contribuyente."""
    return TIPOS_CONTRIBUYENTE.get((perfil or {}).get('tipo'), (True, False))


def tolerancia(valor):
    return np.maximum(UMBRAL_TOLERANCIA, np.abs(valor) * TOLERANCIA_RELATIVA)


def _concepto_cuenta(columna, mapa=None, defecto='compras'):
    mapa = parametros_repo.cargar_conceptos()['gasto'] if mapa is None else mapa
    return parametros_repo.concepto_de(columna, mapa, defecto)[0]


def _tarifa_cuenta(columna):
    """Tarifa indicada en el nombre de la cuenta: 'Compras (2.5%)' -> 0.025."""
    for texto in (str(columna), nombre_cuenta(columna)):
        m = re.search(r'\(\s*(\d+(?:[.,]\d+)?)\s*%\s*\)', texto)
        if m:
            return float(m.group(1).replace(',', '.')) / 100
    return np.nan


def _base_minima(fecha, concepto, parametros=None):
    """(valor en pesos, texto del criterio) de la base mínima vigente a la fecha del comprobante."""
    parametros = parametros or parametros_repo.cargar()
    fecha = fecha if pd.notna(fecha) else pd.Timestamp(date.today())
    uvt = parametros_repo.uvt_vigente(parametros, fecha.year)
    bases = parametros['bases']
    vigente = next((b for desde, b in reversed(bases) if fecha >= pd.Timestamp(desde)), bases[0][1])
    uvts = vigente.get(concepto, vigente['compras'])
    uvts = int(uvts) if float(uvts).is_integer() else uvts
    return uvts * uvt, f"{concepto} {uvts} UVT = {_pesos(uvts * uvt)}"


def a_fecha(serie):
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    texto = serie.map(lambda v: v.isoformat() if hasattr(v, 'isoformat') else str(v)).str.strip()
    iso = texto.str.match(r'^\d{4}-\d{1,2}-\d{1,2}')
    fechas = pd.Series(pd.NaT, index=serie.index, dtype='datetime64[ns]')
    fechas[iso] = pd.to_datetime(texto[iso].str[:10], errors='coerce', format='%Y-%m-%d')
    fechas[~iso] = pd.to_datetime(texto[~iso], errors='coerce', dayfirst=True, format='mixed')
    return fechas


def _suma_abs(df, columnas):
    cols = [c for c in columnas if c in df.columns]
    return sumar_columnas(df, cols).abs() if cols else pd.Series(0.0, index=df.index)


def alertas_retencion(df_aud, seriales_base, seriales_iva, seriales_ret, seriales_reteiva, dict_auto, perfil):
    """Por fila de AUD-COMP: DataFrame con RETEFUENTE, RETEIVA y ALERTA."""
    idx = df_aud.index
    retencion, reteiva, iva = (_suma_abs(df_aud, s) for s in (seriales_ret, seriales_reteiva, seriales_iva))
    alertas = pd.Series('', index=idx)
    resultado = pd.DataFrame({'RETEFUENTE': retencion, 'RETEIVA': reteiva, 'ALERTA': alertas})

    if (perfil or {}).get('periodicidad') == NO_RESPONSABLE_IVA:
        agregar_alerta(alertas, iva > UMBRAL_TOLERANCIA, 'IVA descontable en empresa no responsable de IVA')

    cols_base = [c for c in seriales_base if c in df_aud.columns]
    if not cols_base or df_aud.empty:
        resultado['ALERTA'] = alertas
        return resultado

    valores_base = df_aud[cols_base].apply(pd.to_numeric, errors='coerce').fillna(0).abs()
    base = valores_base.sum(axis=1)
    # Base de retención: sin impuestos llevados al costo (consumo, saludables, IVA mayor valor): no son ingreso del proveedor
    cols_impuesto = [c for c in cols_base if re.search(PATRON_IMPUESTO_EN_COSTO, str(c), re.I)]
    base_ret = base - valores_base[cols_impuesto].sum(axis=1) if cols_impuesto else base
    texto_base = base_ret.map(_pesos) + np.where(base_ret < base - UMBRAL_TOLERANCIA,
                                                 ' (sin ' + (base - base_ret).map(_pesos) + ' de impuestos en el costo)', '')
    conceptos = parametros_repo.cargar_conceptos()
    concepto_gasto = valores_base.idxmax(axis=1).map(lambda c: _concepto_cuenta(c, conceptos['gasto']))

    col_nit = buscar_columna(df_aud.columns, 'Nit/C.C.', 'NIT')
    nit = df_aud[col_nit].map(normalizar_nit) if col_nit else pd.Series('', index=idx)
    col_tipo = buscar_columna(df_aud.columns, 'Tipo', 'TIPO')
    tipo = df_aud[col_tipo].fillna('').astype(str).str.upper() if col_tipo else pd.Series('', index=idx)
    col_fecha = buscar_columna(df_aud.columns, 'Fecha', 'FECHA')
    fecha = a_fecha(df_aud[col_fecha]) if col_fecha else pd.Series(pd.NaT, index=idx)
    comentario = nit.map(dict_auto).fillna('').astype(str).str.upper()

    con_base = (base > 0) & (nit != '')
    agregar_alerta(alertas, (nit != '') & (iva > UMBRAL_TOLERANCIA) & (base <= UMBRAL_TOLERANCIA), 'IVA sin base gravada')

    # Concepto: cuenta de retención usada > historial del tercero > cuenta de gasto
    cols_ret = [c for c in seriales_ret if c in df_aud.columns]
    concepto_ret, col_ret = pd.Series('', index=idx), pd.Series(None, index=idx, dtype=object)
    unica = pd.Series(False, index=idx)
    if cols_ret:
        valores_ret = df_aud[cols_ret].apply(pd.to_numeric, errors='coerce').fillna(0).abs()
        unica = (valores_ret > UMBRAL_TOLERANCIA).sum(axis=1) == 1
        col_ret = valores_ret.idxmax(axis=1).where(retencion > UMBRAL_TOLERANCIA)
        concepto_ret = col_ret.map(lambda c: _concepto_cuenta(c, conceptos['retencion'], '') if isinstance(c, str) else '')
    con_concepto = concepto_ret != ''
    historial = concepto_ret[con_concepto].groupby(nit[con_concepto]).agg(lambda s: s.mode().iat[0])
    concepto = concepto_ret.where(con_concepto, nit.map(historial)).fillna(concepto_gasto)
    concepto = concepto.where(concepto != '', concepto_gasto)
    origen = pd.Series(np.where(con_concepto, 'cuenta de retención',
                                np.where(nit.isin(historial.index), 'historial del tercero', 'cuenta de gasto')), index=idx)

    parametros = parametros_repo.cargar()
    alta = [_base_minima(f, 'compras' if c == 'mixto' else c, parametros) for f, c in zip(fecha, concepto)]
    baja = [_base_minima(f, 'servicios' if c == 'mixto' else c, parametros) for f, c in zip(fecha, concepto)]
    base_alta, criterio_alta = (pd.Series(v, index=idx) for v in zip(*alta))
    base_baja, criterio_baja = (pd.Series(v, index=idx) for v in zip(*baja))
    supera = base_ret >= base_alta
    simple = tipo.str.contains('SIMPLE', na=False) | comentario.str.contains('SIMPLE', na=False)
    agente_rf, agente_iva = agentes_retencion(perfil)

    if seriales_ret:
        con_ret = retencion > UMBRAL_TOLERANCIA
        if agente_rf:
            exime = comentario.map(lambda c: any(k in c for k in COMENTARIOS_NO_SUJETOS))
            no_sujeto = (nit.isin(dict_auto.keys()) & exime) | simple
            observado = nit.isin(dict_auto.keys()) & ~exime & (comentario != '')
            retuvo_antes = nit.isin(set(nit[con_ret & con_base])) & (nit != '')
            agregar_alerta(alertas, con_base & ~no_sujeto & ~observado & supera & ~con_ret,
                           'Sin retención: base ' + texto_base + ' supera mínimo de ' + criterio_alta
                           + ' [concepto por ' + origen + ']'
                           + np.where(retuvo_antes, '; a este tercero sí se le retuvo en otros comprobantes', ''))
            agregar_alerta(alertas, con_base & no_sujeto & con_ret,
                           'Retención de ' + retencion.map(_pesos) + ' practicada a no sujeto ('
                           + np.where(simple & ~nit.isin(dict_auto.keys()), 'RÉGIMEN SIMPLE', comentario) + ')')
            agregar_alerta(alertas, con_base & observado,
                           'Tercero con observación en listado de autorretenedores: ' + comentario
                           + (' (retención ' + retencion.map(_pesos) + ')').where(con_ret, ' (sin retención)'))
            agregar_alerta(alertas, con_base & ~no_sujeto & (base_ret < base_baja) & con_ret,
                           'Retención de ' + retencion.map(_pesos) + ' con base ' + texto_base + ' bajo mínimo de '
                           + criterio_baja + ' [concepto por ' + origen + ']')

            esperada = col_ret.map(lambda c: _tarifa_cuenta(c) if isinstance(c, str) else np.nan).astype(float)
            desvio = (retencion - base_ret * esperada).abs() > np.maximum(UMBRAL_TOLERANCIA, base_ret * esperada * TOLERANCIA_TARIFA_RETENCION)
            con_tarifa = con_base & con_ret & unica & esperada.notna()
            texto_tarifa = ('Tarifa aplicada distinta a la de la cuenta (aplicada '
                            + (retencion / base_ret.replace(0, np.nan)).map(lambda v: _porcentaje(v) if pd.notna(v) else '')
                            + ' vs ' + esperada.map(lambda v: _porcentaje(v) if pd.notna(v) else '')
                            + ' sobre base ' + texto_base + ')')
            agregar_alerta(alertas, con_tarifa & desvio, texto_tarifa)
            agregar_alerta(alertas, con_base & con_ret & ~(con_tarifa) & (retencion > base_ret * TARIFA_MAXIMA_RETENCION + UMBRAL_TOLERANCIA),
                           'Tarifa de retención atípica (' + (retencion / base_ret.replace(0, np.nan)).map(
                               lambda v: _porcentaje(v) if pd.notna(v) else '') + ' sobre base ' + texto_base + ')')

            cruce = con_ret & con_concepto & (concepto_gasto != 'mixto') & (concepto_ret != concepto_gasto) & (
                concepto_ret.isin(CONCEPTOS_ESTRICTOS) | concepto_gasto.isin(CONCEPTOS_ESTRICTOS))
            agregar_alerta(alertas, cruce, 'Cuenta de retención no corresponde al gasto (retención por '
                           + concepto_ret + ', gasto de ' + concepto_gasto + ')')
        else:
            agregar_alerta(alertas, con_ret, 'Retención practicada sin ser agente de retención')

    if seriales_reteiva:
        con_reteiva = reteiva > UMBRAL_TOLERANCIA
        if agente_iva:
            exento = simple | comentario.str.contains('GRAN CONTRIBUYENTE', na=False)
            agregar_alerta(alertas, con_base & supera & (iva > UMBRAL_TOLERANCIA) & ~con_reteiva & ~exento,
                           'Sin reteIVA: agente de IVA y proveedor con IVA')
            difiere = (reteiva - iva * TARIFA_RETEIVA).abs() > np.maximum(UMBRAL_TOLERANCIA, iva * 0.01)
            agregar_alerta(alertas, con_reteiva & difiere, 'ReteIVA distinta al 15% del IVA')
        else:
            agregar_alerta(alertas, con_reteiva, 'ReteIVA practicada sin ser agente de retención de IVA')

    resultado['ALERTA'] = alertas
    return resultado


def _indice_periodo(fechas, meses):
    return fechas.dt.year * (12 // meses) + (fechas.dt.month - 1) // meses


def _nit_empresa(df, es_dian, perfil):
    nit = normalizar_nit((perfil or {}).get('nit', ''))
    if nit:
        return nit
    recibidos = es_dian & contiene(df['Grupo'], 'recib')
    receptores = df.loc[recibidos, 'NIT Receptor'].map(normalizar_nit)
    receptores = receptores[receptores != '']
    return receptores.mode().iat[0] if not receptores.empty else ''


def alertas_conciliacion(df, prioridad_dian, prioridad_cont, mask_pairs, perfil=None):
    """Alertas sobre DIAN VS CONT. Requiere Box_Color calculado."""
    alertas = df[COL_ALERTAS].fillna('').astype(str) if COL_ALERTAS in df.columns else pd.Series('', index=df.index)
    es_dian = df['Prioridad_Fila'] == prioridad_dian
    es_cont = df['Prioridad_Fila'] == prioridad_cont
    contable = ~es_dian
    claves = df['Num.Ext_Clean']

    def por_grupo(serie, func):
        return serie[mask_pairs].groupby(claves[mask_pairs]).transform(func).reindex(df.index)

    # Estado y documento DIAN
    estado = df['Estado'].fillna('').astype(str).str.upper()
    agregar_alerta(alertas, es_dian & estado.str.contains('RECHAZ', na=False), 'Factura rechazada en DIAN')
    es_nc = es_dian & contiene(df['Tipo de documento'], PATRON_NOTA_CREDITO)
    agregar_alerta(alertas, mask_pairs & (por_grupo(es_nc, 'any') == True) & (por_grupo(es_cont, 'any') == True),  # noqa: E712
                   'Nota crédito DIAN cruzada contra compras')

    nit_empresa = _nit_empresa(df, es_dian, perfil)
    if nit_empresa:
        receptor = df['NIT Receptor'].map(normalizar_nit)
        agregar_alerta(alertas, es_dian & contiene(df['Grupo'], 'recib') & (receptor != '') & (receptor != nit_empresa),
                       'Factura a nombre de otro adquirente (art. 771-2 E.T.)')

    # Valores
    base, iva = a_numero(df['BASE']).abs(), a_numero(df['IVA']).abs()
    agregar_alerta(alertas, (base > 0) & (iva > base * TARIFA_MAXIMA_IVA + UMBRAL_TOLERANCIA),
                   'IVA supera el 19% de la base')
    total_dian = por_grupo(a_numero(df['Total']).abs().where(es_dian, 0), 'sum')
    otros_dian = por_grupo(a_numero(df['OTROS IMPUESTOS']).abs().where(es_dian, 0), 'sum')
    total_cont = por_grupo((base + a_numero(df['BASE_2']).abs() + iva).where(contable, 0), 'sum')
    # La BASE DIAN (Total - IVA) ya incluye los otros impuestos de la factura: el total cuadra si coincide con lo
    # contabilizado, o con lo contabilizado más los otros impuestos (cuando estos se registraron aparte y sin marcar)
    tol_total = tolerancia(total_dian)
    descuadre = ((total_dian - total_cont).abs() > tol_total) & ((total_dian - total_cont - otros_dian).abs() > tol_total)
    agregar_alerta(alertas, mask_pairs & (df['Box_Color'] == 'BLACK') & (total_dian > 0) & descuadre,
                   'Total DIAN no cuadra (BASE_2 u otros impuestos)')

    # Periodo de contabilización
    emision = por_grupo(a_fecha(df['Fecha Emisión']).where(es_dian), 'max')
    f_cont = a_fecha(df['Fecha'])
    validas = mask_pairs & contable & emision.notna() & f_cont.notna()
    agregar_alerta(alertas, validas & (f_cont < emision.dt.normalize()), 'Contabilizada antes de la emisión')
    agregar_alerta(alertas, validas & (f_cont.dt.year != emision.dt.year), 'Contabilizada en año distinto a la emisión')
    periodo = PERIODOS_IVA.get((perfil or {}).get('periodicidad'))
    if periodo:
        meses, permitidos = periodo
        fuera = (_indice_periodo(f_cont, meses) - _indice_periodo(emision, meses)) > permitidos
        agregar_alerta(alertas, validas & (iva > UMBRAL_TOLERANCIA) & fuera,
                       'IVA descontable fuera del término (art. 496 E.T.)')

    # Soporte electrónico
    novedad = df['Novedad'].fillna('').astype(str)
    sin_soporte = es_cont & (novedad.str.contains('CONT - Sin Pareja en DIAN', regex=False)
                             | novedad.str.contains('CONT - Sin Num.Ext', regex=False))
    agregar_alerta(alertas, sin_soporte, 'Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)')
    return alertas

# Texto con que empieza (o que contiene) cada alerta -> nombre en GUIA_ALERTAS
_FIRMAS_ALERTA = (
    (r'^Sin retención:', 'Sin retención'),
    (r'practicada a no sujeto', 'Retención practicada a no sujeto'),
    (r'^Tercero con observación', 'Tercero con observación en listado de autorretenedores'),
    (r'bajo mínimo', 'Retención con base bajo el mínimo'),
    (r'^Tarifa aplicada distinta', 'Tarifa aplicada distinta a la de la cuenta'),
    (r'^Tarifa de retención atípica', 'Tarifa de retención atípica'),
    (r'^Cuenta de retención no corresponde', 'Cuenta de retención no corresponde al gasto'),
    (r'^IVA sin base gravada', 'IVA sin base gravada'),
    (r'^Factura rechazada', 'Factura rechazada en DIAN'),
    (r'^Nota crédito DIAN cruzada', 'Nota crédito DIAN cruzada contra compras'),
    (r'^Factura a nombre de otro adquirente', 'Factura a nombre de otro adquirente (art. 771-2 E.T.)'),
    (r'^IVA supera el 19', 'IVA supera el 19 % de la base'),
    (r'^Total DIAN no cuadra', 'Total DIAN no cuadra (BASE_2 u otros impuestos)'),
    (r'^Contabilizada antes', 'Contabilizada antes de la emisión'),
    (r'^Contabilizada en año distinto', 'Contabilizada en año distinto a la emisión'),
    (r'^Sin factura/documento soporte', 'Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)'),
    (r'^IVA descontable en empresa no responsable', 'IVA descontable en empresa no responsable de IVA'),
    (r'^IVA descontable fuera del término', 'IVA descontable fuera del término (art. 496 E.T.)'),
    (r'^Retención practicada sin ser agente', 'Retención practicada sin ser agente de retención'),
    (r'^Sin reteIVA', 'Sin reteIVA: agente de IVA y proveedor con IVA'),
    (r'^ReteIVA distinta', 'ReteIVA distinta al 15 % del IVA'),
    (r'^ReteIVA practicada sin ser agente', 'ReteIVA practicada sin ser agente de retención de IVA'),
)


def tipo_alerta(texto):
    """Nombre de la alerta en GUIA_ALERTAS para un texto de alerta (con valores); '' si no se reconoce."""
    return next((nombre for patron, nombre in _FIRMAS_ALERTA if re.search(patron, str(texto))), '')


def explicacion_alerta(nombre):
    """(sección, cuándo aparece, qué significa) de la guía para el nombre de alerta."""
    from core.config import GUIA_ALERTAS
    for seccion, _, alertas in GUIA_ALERTAS:
        for alerta, cuando, significa in alertas:
            if alerta == nombre:
                return seccion, cuando, significa
    return '', '', ''