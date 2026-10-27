"""Acta (memorando) en PDF con las inconsistencias de la conciliación DIAN vs Contabilidad."""
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from shared.formato import nit_texto

ORIGENES = {1: "DIAN", 2: "CONTAB.", 3: "DEVOL."}
NOVEDADES_SIN_PAREJA = ['DIAN - Sin Pareja en Contabilidad', 'CONT - Sin Pareja en DIAN', 'DEVOL - Sin Pareja en DIAN', '']

ESTILO_BASE_TABLA = [
    ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
]


def _moneda(valor):
    return f"$ {pd.to_numeric(valor, errors='coerce'):,.2f}"


def _origen(prioridad, con_devol=False):
    if con_devol:
        return ORIGENES.get(prioridad, "DEVOL.")
    return "DIAN" if prioridad == 1 else "CONTAB."


class GeneradorActaErrores:
    def __init__(self, df_unificado, ruta_original):
        self.df = df_unificado.copy()
        self.ruta_original = Path(ruta_original)
        base = getSampleStyleSheet()
        self.estilo_titulo = ParagraphStyle('TituloActa', parent=base['Heading1'], fontName='Helvetica-Bold',
                                            fontSize=16, textColor=colors.black, alignment=TA_CENTER, spaceAfter=20)
        self.estilo_subtitulo = ParagraphStyle('Subtitulo', parent=base['Heading2'], fontName='Helvetica-Bold',
                                               fontSize=12, textColor=colors.darkgrey, spaceBefore=15, spaceAfter=10)
        self.estilo_normal = ParagraphStyle('TextoNormal', parent=base['Normal'], fontName='Helvetica',
                                            fontSize=10, alignment=TA_JUSTIFY, spaceAfter=10)
        self.estilo_celda = ParagraphStyle('Celda', parent=base['Normal'], fontName='Helvetica',
                                           fontSize=8, alignment=TA_LEFT)

    def nombre_sugerido(self):
        return f"Acta_Auditoria_{self.ruta_original.stem}_{datetime.now().strftime('%Y%m%d')}.pdf"

    def _p(self, texto):
        return Paragraph(str(texto), self.estilo_celda)

    def generar(self, ruta_pdf):
        elementos = []
        self._encabezado(elementos)
        hay_errores = False
        hay_errores |= self._seccion_parejas_discrepantes(elementos)
        hay_errores |= self._seccion_huerfanos(elementos)
        hay_errores |= self._seccion_otros_errores(elementos)
        if not hay_errores:
            elementos.append(Paragraph("<b>RESULTADO EXCELENTE:</b> No se encontraron inconsistencias en la auditoría.", self.estilo_normal))

        SimpleDocTemplate(ruta_pdf, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40).build(elementos)
        return ruta_pdf

    def _encabezado(self, elementos):
        elementos.append(Paragraph("MEMORANDO DE AUDITORÍA CONTABLE", self.estilo_titulo))
        elementos.append(Paragraph(f"""
        <b>Fecha de emisión:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}<br/>
        <b>Archivo auditado:</b> {self.ruta_original.name}<br/><br/>
        <b>Asunto:</b> Acta de inconsistencias y correcciones requeridas.<br/><br/>
        El presente documento detalla las discrepancias halladas durante el cruce automatizado
        entre el reporte de la DIAN y los movimientos contables registrados. Se requiere la
        revisión y ajuste correspondiente en el software contable de los siguientes hallazgos:
        """, self.estilo_normal))
        elementos.append(Spacer(1, 10))

    def _seccion_parejas_discrepantes(self, elementos):
        df = self.df[self.df['Conteo_Pareja'] >= 2]
        df = df[(df['Diff_NIT'] == True) | (df['Diff_BASE'] == True) | (df['Diff_IVA'] == True) | (df['Box_Color'] == 'YELLOW')]  # noqa: E712
        if df.empty:
            return False

        elementos.append(Paragraph("1. DISCREPANCIAS EN REGISTROS EMPAREJADOS", self.estilo_subtitulo))
        elementos.append(Paragraph(
            "Las siguientes facturas fueron encontradas en ambos reportes, pero presentan diferencias en el "
            "<b>NIT del tercero</b>, los <b>valores contabilizados (Base/IVA)</b>, o <b>no generan efecto cero</b> "
            "tras sumar comprobantes y devoluciones:", self.estilo_normal))

        for num_ext, grupo in df.groupby('Num.Ext_Clean'):
            grupo = grupo.sort_values(by='Prioridad_Fila')
            primera = grupo.iloc[0]
            errores = [txt for cond, txt in (
                (primera['Box_Color'] == 'YELLOW', "No hay Efecto Cero"),
                (primera['Diff_BASE'] == True, "Diferencia en Base"),  # noqa: E712
                (primera['Diff_IVA'] == True, "Diferencia en IVA"),  # noqa: E712
                (primera['Diff_NIT'] == True, "NIT no coincide"),  # noqa: E712
            ) if cond]
            elementos.append(self._p(f"<b>Num.Ext: {num_ext}</b> - <i>Errores: {', '.join(errores)}</i>"))

            datos = [["Origen", "NIT Emisor", "Tercero", "Base", "IVA", "Tipo Doc"]]
            for _, f in grupo.iterrows():
                datos.append([
                    self._p(_origen(f['Prioridad_Fila'], con_devol=True)),
                    self._p(nit_texto(f['NIT Emisor'])),
                    self._p(str(f['Nombre Emisor'])[:25]),
                    _moneda(f['BASE']), _moneda(f['IVA']),
                    self._p(f['Tipo de documento']),
                ])
            tabla = Table(datos, colWidths=[50, 70, 160, 80, 80, 70])
            tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (3, 1), (4, -1), 'RIGHT'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            elementos += [tabla, Spacer(1, 10)]
        return True

    def _seccion_huerfanos(self, elementos):
        df = self.df[(self.df['Conteo_Pareja'] == 1) & self.df['Novedad'].str.contains('Sin Pareja', na=False)]
        if df.empty:
            return False

        elementos.append(Paragraph("2. DOCUMENTOS FALTANTES (HUÉRFANOS)", self.estilo_subtitulo))
        elementos.append(Paragraph("Los siguientes registros existen en una plataforma pero no en la otra:", self.estilo_normal))

        datos = [["Origen", "Factura", "NIT", "Tercero", "Base", "Observación"]]
        for _, f in df.sort_values(by=['Prioridad_Fila', 'Num.Ext']).iterrows():
            datos.append([
                _origen(f['Prioridad_Fila']), str(f['Num.Ext']), nit_texto(f['NIT Emisor']),
                self._p(str(f['Nombre Emisor'])[:30]), _moneda(f['BASE']), self._p(f['Novedad']),
            ])
        tabla = Table(datos, colWidths=[50, 70, 70, 140, 70, 110], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('ALIGN', (4, 1), (4, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos += [tabla, Spacer(1, 10)]
        return True

    def _seccion_otros_errores(self, elementos):
        nov = self.df['Novedad']
        df = self.df[
            ~nov.isin(NOVEDADES_SIN_PAREJA) & nov.notna()
            & ~nov.str.contains('Sin Pareja', na=False)
            & ~nov.str.contains('Gasto Personal', na=False, case=False)
            & ~nov.str.contains('Application Response', na=False, case=False)
            & ~nov.str.startswith('DIAN:')
        ]
        if df.empty:
            return False

        elementos.append(Paragraph("3. DUPLICADOS Y ERRORES DE ESTRUCTURA", self.estilo_subtitulo))
        elementos.append(Paragraph(
            "Se detectaron problemas en la digitación, documentos registrados más de una vez, o falta de número de factura:",
            self.estilo_normal))

        datos = [["Origen", "Factura Digitada", "Tercero", "Error Detectado"]]
        for _, f in df.sort_values(by=['Novedad', 'Prioridad_Fila']).iterrows():
            datos.append([
                _origen(f['Prioridad_Fila']), str(f['Num.Ext_Original']),
                self._p(str(f['Nombre Emisor'])[:35]), self._p(f['Novedad']),
            ])
        tabla = Table(datos, colWidths=[50, 100, 170, 190], repeatRows=1)
        tabla.setStyle(TableStyle(ESTILO_BASE_TABLA + [('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        elementos.append(tabla)
        return True
