"""Fila de totales al final de las hojas (Token, MOVS_AUD / AUD-COMP).

Es informativa: usa SUBTOTAL (respeta filtros) y toda lectura de esas hojas la descarta con `quitar_fila_totales`,
así que no afecta la conciliación ni los informes.
"""
import pandas as pd
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter

MARCA_TOTALES = "TOTALES"
FORMATO_TOTAL = '#,##0.00'
_BORDE_TOTAL = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))


def es_marca(valor):
    return str(valor).strip().upper() == MARCA_TOTALES


def quitar_fila_totales(df, columnas=3):
    """Descarta las filas cuya marca TOTALES está en alguna de las primeras columnas."""
    if df is None or df.empty:
        return df
    marca = df.iloc[:, :columnas].apply(lambda c: c.map(es_marca)).any(axis=1)
    return df[~marca] if marca.any() else df


def quitar_filas_totales_ws(ws, fila_encabezado=1, columnas=3):
    """Elimina de la hoja (openpyxl) las filas de totales de ejecuciones anteriores."""
    for r in range(ws.max_row, fila_encabezado, -1):
        if any(es_marca(ws.cell(row=r, column=c).value) for c in range(1, columnas + 1)):
            ws.delete_rows(r)


def agregar_fila_totales_ws(ws, columnas, fila_inicio, fila_fin, columna_etiqueta=1):
    """Escribe TOTALES y =SUBTOTAL(9, ...) en `columnas` (índices 1-based) una fila debajo de los datos."""
    if fila_fin < fila_inicio or not columnas:
        return None
    from shared.excel_io import registrar_valor_formula

    fila = fila_fin + 1
    etiqueta = ws.cell(row=fila, column=columna_etiqueta, value=MARCA_TOTALES)
    etiqueta.font = Font(bold=True)
    for c in columnas:
        letra = get_column_letter(c)
        celda = ws.cell(row=fila, column=c, value=f"=SUBTOTAL(9,{letra}{fila_inicio}:{letra}{fila_fin})")
        suma = sum(v for v in (ws.cell(row=r, column=c).value for r in range(fila_inicio, fila_fin + 1))
                   if isinstance(v, (int, float)) and not isinstance(v, bool))
        registrar_valor_formula(ws, f"{letra}{fila}", suma)
        celda.font = Font(bold=True)
        celda.number_format = FORMATO_TOTAL
        celda.border = _BORDE_TOTAL
    return fila


def columnas_numericas(df, excluir=()):
    """Índices (1-based) de las columnas con valores numéricos (montos) del DataFrame."""
    indices = []
    for i, col in enumerate(df.columns, start=1):
        if col in excluir:
            continue
        serie = pd.to_numeric(df[col], errors="coerce")
        if serie.notna().any() and serie.notna().sum() >= df[col].notna().sum() * 0.9:
            indices.append(i)
    return indices