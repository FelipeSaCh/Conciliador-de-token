"""Render en PDF del informe de IVA."""
from datetime import datetime
from xml.sax.saxutils import escape as escapar_xml

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from core.config import COLUMNAS_OTROS_IMPUESTOS
from core.errors import ErrorSistema
from shared.formato import fmt_entero, fmt_moneda, texto_celda
from services.reportes.informe_iva import COLUMNA_IVA, NOMBRES_CONCEPTOS, etiqueta_saldo

COLUMNAS_EXCLUIDAS_PDF = {
    "CUFE/CUDE", "CONCEPTO", "TERCERO", "TIPO", "TIPO-DETALLE", "BASE", "Num.Ext", "Signo", "IVA Ajustado",
    *COLUMNAS_OTROS_IMPUESTOS,
}
COLUMNAS_MONEDA = {COLUMNA_IVA, "Total"}
COLUMNAS_IDENTIFICADORES = {"NIT Emisor", "NIT Receptor", "Forma de Pago", "Medio de Pago"}
COLUMNAS_ANCHAS = {COLUMNA_IVA, "Total", "NIT Emisor", "NIT Receptor"}
MARGEN = 0.8 * cm


def _estilos():
    s = getSampleStyleSheet()
    celda = ParagraphStyle("Celda", parent=s["Normal"], fontName="Helvetica", fontSize=5.5, leading=6.5, textColor=colors.black)
    resumen = ParagraphStyle("CeldaResumen", parent=s["Normal"], fontName="Helvetica", fontSize=8, leading=10, textColor=colors.black)
    return {
        "titulo": ParagraphStyle("TituloInforme", parent=s["Title"], fontName="Helvetica-Bold", fontSize=14,
                                 textColor=colors.black, alignment=TA_CENTER, spaceAfter=2),
        "subtitulo": ParagraphStyle("Subtitulo", parent=s["Normal"], fontName="Helvetica", fontSize=8,
                                    textColor=colors.HexColor("#333333"), alignment=TA_CENTER, spaceAfter=0),
        "seccion": ParagraphStyle("Seccion", parent=s["Heading2"], fontName="Helvetica-Bold", fontSize=10, textColor=colors.black),
        "celda": celda,
        "celda_num": ParagraphStyle("CeldaNum", parent=celda, alignment=TA_RIGHT),
        "enc_tabla": ParagraphStyle("EncTabla", parent=s["Normal"], fontName="Helvetica-Bold", fontSize=6, leading=7,
                                    textColor=colors.white, alignment=TA_CENTER),
        "resumen": resumen,
        "resumen_num": ParagraphStyle("CeldaResumenNum", parent=resumen, alignment=TA_RIGHT),
        "enc_resumen": ParagraphStyle("EncResumen", parent=s["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10,
                                      textColor=colors.white, alignment=TA_CENTER),
    }


def _seccion(texto, estilo, ancho):
    tabla = Table([[Paragraph(texto, estilo)]], colWidths=[ancho])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return tabla


def _anchos_columnas(df, columnas, ancho_disponible):
    longitudes = []
    for col in columnas:
        peso = 1.5 if col in COLUMNAS_ANCHAS else 1.0
        max_len = max([len(str(col))] + [len("" if pd.isna(v) else str(v)) for v in df[col]])
        longitudes.append(max(max_len * peso, 4))
    total = sum(longitudes)
    return [max(ancho_disponible * (l / total), 15) for l in longitudes]


def _tabla_resumen(resumen, e, ancho):
    filas = [[Paragraph(h, e["enc_resumen"]) for h in
              ("Concepto", "Facturas y ND<br/>(cant. / IVA)", "Notas Crédito<br/>(cant. / IVA)", "IVA Neto")]]
    for grupo in ("Emitido", "Recibido"):
        d = resumen["grupos"][grupo]
        filas.append([
            Paragraph(NOMBRES_CONCEPTOS[grupo], e["resumen"]),
            Paragraph(f"{d['facturas_cantidad']} / {fmt_moneda(d['facturas_iva'])}", e["resumen_num"]),
            Paragraph(f"{d['notas_cantidad']} / {fmt_moneda(d['notas_iva'])}", e["resumen_num"]),
            Paragraph(fmt_moneda(d["iva_neto"]), e["resumen_num"]),
        ])
    total = resumen["totales"]["iva_a_pagar"]
    filas.append([
        Paragraph(f"<b>{etiqueta_saldo(total)}</b>", e["resumen"]), "", "",
        Paragraph(f"<b>{fmt_moneda(total)}</b>", e["resumen_num"]),
    ])
    tabla = Table(filas, colWidths=[ancho * 0.34, ancho * 0.24, ancho * 0.24, ancho * 0.18], hAlign="LEFT")
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F2F2F2")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DDDDDD")),
        ("SPAN", (1, -1), (2, -1)),
        ("LINEABOVE", (0, -1), (-1, -1), 1, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]))
    return tabla


def _tabla_detalle(df, e, ancho):
    columnas = [c for c in df.columns if str(c).strip() not in COLUMNAS_EXCLUIDAS_PDF]

    def celda(col, valor):
        if col in COLUMNAS_MONEDA:
            return Paragraph(fmt_moneda(valor), e["celda_num"])
        if col in COLUMNAS_IDENTIFICADORES:
            return Paragraph(fmt_entero(valor), e["celda_num"] if "NIT" in col else e["celda"])
        return Paragraph(texto_celda(valor), e["celda"])

    filas = [[Paragraph(escapar_xml(str(c)), e["enc_tabla"]) for c in columnas]]
    filas += [[celda(c, v) for c, v in zip(columnas, fila)] for fila in df[columnas].itertuples(index=False, name=None)]

    tabla = Table(filas, colWidths=_anchos_columnas(df, columnas, ancho), repeatRows=1, hAlign="LEFT")
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#333333")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B0B0B0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    return tabla


def construir_pdf_iva(ruta_pdf, df, resumen, periodo_texto, hoja_token):
    ancho_pagina, _ = landscape(letter)
    ancho_util = ancho_pagina - 2 * MARGEN
    e = _estilos()
    ahora = datetime.now().strftime('%d/%m/%Y %H:%M')

    historia = [Paragraph("INFORME DE IVA", e["titulo"])]
    if periodo_texto:
        historia += [Paragraph(escapar_xml(periodo_texto), e["subtitulo"]), Spacer(1, 2)]
    historia += [
        Paragraph(f"Hoja origen: {escapar_xml(hoja_token)}   |   Generado: {ahora}", e["subtitulo"]),
        Spacer(1, 4),
        HRFlowable(width="100%", thickness=1, color=colors.black, spaceBefore=2, spaceAfter=8),
        _seccion("RESUMEN GENERAL", e["seccion"], ancho_util), Spacer(1, 4),
        _tabla_resumen(resumen, e, ancho_util), Spacer(1, 10),
        _seccion("DETALLE DE REGISTROS", e["seccion"], ancho_util), Spacer(1, 4),
        _tabla_detalle(df, e, ancho_util),
    ]

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#999999"))
        canvas.line(MARGEN, 0.9 * cm, ancho_pagina - MARGEN, 0.9 * cm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#555555"))
        canvas.drawString(MARGEN, 0.6 * cm, f"Informe generado el {ahora}")
        canvas.drawRightString(ancho_pagina - MARGEN, 0.6 * cm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        ruta_pdf, pagesize=landscape(letter), leftMargin=MARGEN, rightMargin=MARGEN,
        topMargin=1.0 * cm, bottomMargin=1.2 * cm, title="Informe de IVA",
    )
    try:
        doc.build(historia, onFirstPage=pie, onLaterPages=pie)
    except Exception as ex:
        raise ErrorSistema(f"No se pudo generar el PDF del informe: {ex}")
