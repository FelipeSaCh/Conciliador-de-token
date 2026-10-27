from xml.sax.saxutils import escape as escapar_xml

import pandas as pd


def fmt_moneda(valor):
    try:
        return f"$ {float(valor):,.2f}"
    except (TypeError, ValueError):
        return escapar_xml(str(valor))


def fmt_entero(valor):
    if pd.isna(valor):
        return ""
    try:
        val_float = float(valor)
        return str(int(val_float)) if val_float.is_integer() else str(valor)
    except (TypeError, ValueError):
        return escapar_xml(str(valor))


def texto_celda(valor):
    if pd.isna(valor):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return escapar_xml(str(valor))


def nit_texto(valor):
    return str(valor).replace('.0', '')
