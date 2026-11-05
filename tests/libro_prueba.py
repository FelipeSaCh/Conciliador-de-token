"""Libro de Excel mínimo (Token, Contabilidad, Terceros, AUD-COMP) para pruebas de conciliación completas."""
from openpyxl import Workbook

COLUMNAS_TOKEN = ['Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Divisa', 'Forma de Pago', 'Medio de Pago',
                  'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor', 'Nombre Emisor', 'NIT Receptor', 'Nombre Receptor',
                  'IVA', 'INC', 'Total', 'Estado', 'Grupo']
FACTURAS = [  # (folio, prefijo, nit, nombre, iva, total)
    ("100", "FE", "800111222", "PROVEEDOR UNO", 19000, 119000),
    ("200", "FV", "800333444", "PROVEEDOR PERSONAL", 1900, 11900),
    ("300", "AB", "860001942", "BAYER S.A.", 3800, 23800),
    ("400", "XX", "900555666", "SIN CONTABILIZAR", 0, 50000),
]
CONTABILIDAD = [("800111222", "GASTO", "PAPELERIA"), ("800333444", "GASTO", "RESTAURANTE"),
                ("860001942", "COSTO", "MATERIA PRIMA")]
AUD_COMP = [  # (num, num.ext, nit, tercero, base, iva, retención)
    ("1", "FE100", 800111222, "PROVEEDOR UNO", 100000, 19000, -2500),
    ("2", "FV200", 800333444, "PROVEEDOR PERSONAL", 10000, 1900, 0),
    ("3", "AB300", 860001942.0, "BAYER S.A.", 20000, 3800, 0),
]
SERIALES = {"iva": "240802019 - IVA DESCONTABLE", "base": "143505 - MERCANCIAS", "ret": "236540 - Compras (2.5%)"}
HOJAS = {"principal": "Token", "contabilidad": "contabilidad", "terceros": "TERCEROS", "aud_comp": "AUD-COMP", "Aud_dc": ""}


def crear_libro(ruta, nombre_conta="contabilidad", nombre_terceros="TERCEROS", tipos=None):
    tipos = tipos or {}
    wb = Workbook()
    ws = wb.active
    ws.title = "Token"
    ws.append(COLUMNAS_TOKEN)
    for i, (folio, pref, nit, nombre, iva, total) in enumerate(FACTURAS, start=1):
        ws.append(["Factura electrónica", f"c{i}", folio, pref, "COP", "Contado", "Efectivo", "2026-09-01",
                   "2026-09-02", nit, nombre, "901999888", "MI EMPRESA", iva, 0, total, "Aprobado", "Recibido"])
    c = wb.create_sheet(nombre_conta)
    c.append(["Conceptos por tercero"])
    c.append([])
    c.append(["NIT", "NOMBRE", "TIPO", "DETALLE", "TIPO-DETALLE"])
    for fila, (nit, tipo, detalle) in enumerate(CONTABILIDAD, start=4):
        c.append([nit, "x", tipos.get(nit, tipo), detalle, f'=CONCATENATE(C{fila}," ",D{fila})'])
    t = wb.create_sheet(nombre_terceros)
    t.append(["NIT", "NOMBRE"])
    for nit in ("800111222", "860001942"):
        t.append([nit, "x"])
    a = wb.create_sheet("AUD-COMP")
    a.append(["Tipo", "Número", "Num.Ext", "Fecha", "Nit/C.C.", "Tercero", "Detalle",
              SERIALES["base"], SERIALES["iva"], SERIALES["ret"], "220505 - PROVEEDORES"])
    for num, num_ext, nit, tercero, base, iva, ret in AUD_COMP:
        a.append(["FC1", num, num_ext, "2026-09-03", nit, tercero, "COMPRA", base, iva, ret, -(base + iva + ret)])
    wb.save(ruta)
    return ruta