"""Ubica en el libro las hojas Token, Contabilidad y Terceros: primero la configurada, luego por nombre y, si hace
falta, por contenido (encabezados). Cada decisión se informa en la consola para que el usuario sepa qué hoja se usó."""
import pandas as pd

from shared.excel_io import leer_excel
from shared.texto import normalizar_ascii

COLUMNAS_TOKEN = ('NIT Emisor', 'NIT Receptor', 'Grupo', 'Prefijo', 'Folio', 'Total', 'IVA')
COLUMNAS_CONTABILIDAD = ('NIT', 'TIPO', 'DETALLE')
HOJAS_GENERADAS = {'dianvscont', 'resultados', 'auditoria', 'resultadosauditoria', 'informeiva'}
ETIQUETAS = {'principal': 'Token', 'contabilidad': 'Contabilidad', 'terceros': 'Terceros'}
FILAS_MUESTRA = 15


def clave(texto):
    return "".join(ch for ch in normalizar_ascii(texto) if ch.isalnum())


def _muestras(ruta):
    try:
        return leer_excel(ruta, rapido=True, sheet_name=None, header=None, nrows=FILAS_MUESTRA)
    except Exception:
        return {}


def _filas_encabezado(df, solo_primera=False):
    filas = df.head(1) if solo_primera else df
    return [{clave(v) for v in fila if pd.notna(v) and str(v).strip()} for fila in filas.itertuples(index=False)]


def tiene_columnas(df, columnas, solo_primera=False):
    requeridas = {clave(c) for c in columnas}
    return df is not None and any(requeridas <= fila for fila in _filas_encabezado(df, solo_primera))


def _por_nombre(hojas, palabra, excluir=()):
    objetivo = clave(palabra)
    exactas = [h for h in hojas if clave(h) == objetivo and h not in excluir]
    return exactas[0] if exactas else next((h for h in hojas if objetivo in clave(h) and h not in excluir), None)


def resolver_hojas(ruta, sheet_names):
    """(sheet_names con principal/contabilidad/terceros resueltos, [mensajes para la consola])."""
    muestras = _muestras(ruta)
    hojas = list(muestras)
    resuelto = dict(sheet_names)
    mensajes = []
    reservadas = {sheet_names.get('aud_comp'), sheet_names.get('Aud_dc')} - {None, ''}
    candidatas = [h for h in hojas if h not in reservadas and clave(h) not in HOJAS_GENERADAS]

    def decidir(clave_hoja, encontrada, motivo, sin_resultado):
        configurada = (sheet_names.get(clave_hoja) or '').strip()
        resuelto[clave_hoja] = encontrada or ''
        etiqueta = ETIQUETAS[clave_hoja]
        if not encontrada:
            mensajes.append(f"  Hoja {etiqueta}: no se encontró en el libro. {sin_resultado}")
        elif encontrada == configurada:
            mensajes.append(f"  Hoja {etiqueta}: '{encontrada}' (la configurada).")
        else:
            previa = f"la configurada '{configurada}' no existe o no tiene las columnas necesarias" if configurada \
                else "no se configuró ninguna"
            mensajes.append(f"  Hoja {etiqueta}: se usa '{encontrada}' ({motivo}; {previa}).")

    # Token: encabezado en la fila 1 con las columnas de la factura electrónica
    configurada = (sheet_names.get('principal') or '').strip()
    if configurada in muestras and tiene_columnas(muestras[configurada], COLUMNAS_TOKEN, solo_primera=True):
        encontrada, motivo = configurada, ''
    else:
        encontrada = next((h for h in candidatas if tiene_columnas(muestras[h], COLUMNAS_TOKEN, solo_primera=True)), None)
        motivo = "tiene las columnas de la factura electrónica (NIT Emisor, Prefijo, Folio, Total, IVA…)"
        if not encontrada and configurada in muestras:
            encontrada, motivo = configurada, ''
    decidir('principal', encontrada, motivo, "Selecciónala en Configurar auditoría.")
    if resuelto['principal']:
        candidatas = [h for h in candidatas if h != resuelto['principal']]

    # Contabilidad: la configurada, por nombre o por contenido (NIT + TIPO + DETALLE)
    configurada = (sheet_names.get('contabilidad') or '').strip()
    if configurada in candidatas:
        encontrada, motivo = configurada, ''
    else:
        encontrada = _por_nombre(candidatas, 'contabilidad')
        motivo = "encontrada por nombre"
        if not encontrada:
            encontrada = next((h for h in candidatas if tiene_columnas(muestras[h], COLUMNAS_CONTABILIDAD)), None)
            motivo = "tiene las columnas NIT, TIPO y DETALLE"
    decidir('contabilidad', encontrada, motivo,
            "Todos los conceptos quedarán 'CONCEPTO NO CREADO' y no se podrán identificar los gastos personales.")
    if resuelto['contabilidad']:
        candidatas = [h for h in candidatas if h != resuelto['contabilidad']]

    # Terceros: la configurada o por nombre (por contenido no: muchas hojas tienen columna NIT)
    configurada = (sheet_names.get('terceros') or '').strip()
    encontrada = configurada if configurada in candidatas else _por_nombre(candidatas, 'terceros') \
        or _por_nombre(candidatas, 'tercero')
    decidir('terceros', encontrada, "encontrada por nombre",
            "Todos los terceros quedarán 'NO CREADO'.")
    return resuelto, mensajes