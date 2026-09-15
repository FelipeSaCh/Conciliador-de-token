# acta_errores.py
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT

class GeneradorActaErrores:
    """
    Genera un acta formal en PDF detallando las inconsistencias encontradas 
    en la conciliación DIAN vs Contabilidad.
    """
    
    def __init__(self, df_unificado, ruta_original):
        self.df = df_unificado.copy()
        self.ruta_original = Path(ruta_original)
        self.estilos = getSampleStyleSheet()
        self._configurar_estilos()

    def _configurar_estilos(self):
        """Define estilos minimalistas, elegantes y en escala de grises."""
        self.estilo_titulo = ParagraphStyle(
            'TituloActa',
            parent=self.estilos['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=16,
            textColor=colors.black,
            alignment=TA_CENTER,
            spaceAfter=20
        )
        self.estilo_subtitulo = ParagraphStyle(
            'Subtitulo',
            parent=self.estilos['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=12,
            textColor=colors.darkgrey,
            spaceBefore=15,
            spaceAfter=10,
            borders={'bottom': 1, 'color': colors.grey}
        )
        self.estilo_normal = ParagraphStyle(
            'TextoNormal',
            parent=self.estilos['Normal'],
            fontName='Helvetica',
            fontSize=10,
            alignment=TA_JUSTIFY,
            spaceAfter=10
        )
        self.estilo_celda = ParagraphStyle(
            'Celda',
            parent=self.estilos['Normal'],
            fontName='Helvetica',
            fontSize=8,
            alignment=TA_LEFT
        )

    def pedir_ruta_guardado(self):
        """Abre un cuadro de diálogo para que el usuario elija dónde guardar el PDF."""
        fecha_str = datetime.now().strftime("%Y%m%d")
        nombre_por_defecto = f"Acta_Auditoria_{self.ruta_original.stem}_{fecha_str}.pdf"
        
        # Ocultar ventana principal de Tkinter si se abre sola
        root = tk.Tk()
        root.withdraw()
        # Asegurar que la ventana de diálogo quede al frente
        root.attributes('-topmost', True) 

        ruta_pdf = filedialog.asksaveasfilename(
            title="Guardar Acta de Errores (PDF)",
            initialdir=self.ruta_original.parent,
            initialfile=nombre_por_defecto,
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        
        root.destroy()
        return ruta_pdf

    def generar(self):
        ruta_pdf = self.pedir_ruta_guardado()
        if not ruta_pdf:
            return None # El usuario canceló

        doc = SimpleDocTemplate(
            ruta_pdf,
            pagesize=letter,
            rightMargin=40,
            leftMargin=40,
            topMargin=40,
            bottomMargin=40
        )

        elementos = []
        self._construir_encabezado(elementos)
        
        # Hay errores si el DataFrame tiene novedades, diferencias numéricas o de NIT
        hay_errores = False

        # 1. Analizar Parejas con Diferencias (Efecto Cero, Bases, IVA, NIT)
        hay_errores |= self._seccion_parejas_discrepantes(elementos)
        
        # 2. Analizar Registros Huérfanos (Faltan en Contabilidad o en DIAN)
        hay_errores |= self._seccion_huerfanos(elementos)
        
        # 3. Analizar Duplicados y Formatos
        hay_errores |= self._seccion_otros_errores(elementos)

        if not hay_errores:
            elementos.append(Paragraph("<b>RESULTADO EXCELENTE:</b> No se encontraron inconsistencias en la auditoría.", self.estilo_normal))

       

        doc.build(elementos)
        return ruta_pdf

    def _construir_encabezado(self, elementos):
        elementos.append(Paragraph("MEMORANDO DE AUDITORÍA CONTABLE", self.estilo_titulo))
        
        texto_intro = f"""
        <b>Fecha de emisión:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}<br/>
        <b>Archivo auditado:</b> {self.ruta_original.name}<br/><br/>
        
        <b>Asunto:</b> Acta de inconsistencias y correcciones requeridas.<br/><br/>
        El presente documento detalla las discrepancias halladas durante el cruce automatizado 
        entre el reporte de la DIAN y los movimientos contables registrados. Se requiere la 
        revisión y ajuste correspondiente en el software contable de los siguientes hallazgos:
        """
        elementos.append(Paragraph(texto_intro, self.estilo_normal))
        elementos.append(Spacer(1, 10))



    def _seccion_parejas_discrepantes(self, elementos):
        """Extrae parejas (DIAN y CONT) donde los valores no cruzan (Efecto Cero), NIT es distinto, etc."""
        df_parejas = self.df[self.df['Conteo_Pareja'] >= 2].copy()
        
        # Filtrar aquellas que tengan alerta en NIT, BASE, IVA o Box_Color (Efecto no cero)
        mask_error = (df_parejas['Diff_NIT'] == True) | \
                     (df_parejas['Diff_BASE'] == True) | \
                     (df_parejas['Diff_IVA'] == True) | \
                     (df_parejas['Box_Color'] == 'YELLOW')
                     
        df_errores = df_parejas[mask_error]
        
        if df_errores.empty:
            return False

        elementos.append(Paragraph("1. DISCREPANCIAS EN REGISTROS EMPAREJADOS", self.estilo_subtitulo))
        desc = "Las siguientes facturas fueron encontradas en ambos reportes, pero presentan diferencias en el <b>NIT del tercero</b>, los <b>valores contabilizados (Base/IVA)</b>, o <b>no generan efecto cero</b> tras sumar comprobantes y devoluciones:"
        elementos.append(Paragraph(desc, self.estilo_normal))

        # Agrupar por la llave única
        agrupado = df_errores.groupby('Num.Ext_Clean')
        
        for num_ext, grupo in agrupado:
            # Ordenar para que DIAN (1) salga primero, luego CONT (2), luego DEVOL (3)
            grupo = grupo.sort_values(by='Prioridad_Fila')
            
            # Determinar el tipo de error para el título del bloque
            errores_txt = []
            if grupo['Box_Color'].iloc[0] == 'YELLOW': errores_txt.append("No hay Efecto Cero")
            if grupo['Diff_BASE'].iloc[0] == True: errores_txt.append("Diferencia en Base")
            if grupo['Diff_IVA'].iloc[0] == True: errores_txt.append("Diferencia en IVA")
            if grupo['Diff_NIT'].iloc[0] == True: errores_txt.append("NIT no coincide")
            
            titulo_error = f"<b>Num.Ext: {num_ext}</b> - <i>Errores: {', '.join(errores_txt)}</i>"
            elementos.append(Paragraph(titulo_error, self.estilo_celda))
            
            # Construir tabla para este grupo
            datos_tabla = [["Origen", "NIT Emisor", "Tercero", "Base", "IVA", "Tipo Doc"]]
            
            for _, fila in grupo.iterrows():
                origen = "DIAN" if fila['Prioridad_Fila'] == 1 else ("CONTAB." if fila['Prioridad_Fila'] == 2 else "DEVOL.")
                datos_tabla.append([
                    Paragraph(origen, self.estilo_celda),
                    Paragraph(str(fila['NIT Emisor']).replace('.0',''), self.estilo_celda),
                    Paragraph(str(fila['Nombre Emisor'])[:25], self.estilo_celda),
                    f"$ {pd.to_numeric(fila['BASE'], errors='coerce'):,.2f}",
                    f"$ {pd.to_numeric(fila['IVA'], errors='coerce'):,.2f}",
                    Paragraph(str(fila['Tipo de documento']), self.estilo_celda)
                ])
                
            tabla = Table(datos_tabla, colWidths=[50, 70, 160, 80, 80, 70])
            tabla.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
                ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('ALIGN', (3,1), (4,-1), 'RIGHT'), # Números a la derecha
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ]))
            elementos.append(tabla)
            elementos.append(Spacer(1, 10))
            
        return True

    def _seccion_huerfanos(self, elementos):
        """Registros que están en DIAN pero no en Contabilidad, y viceversa."""
        mask_huerfanos = (self.df['Conteo_Pareja'] == 1) & (self.df['Novedad'].str.contains('Sin Pareja', na=False))
        df_huerfanos = self.df[mask_huerfanos]
        
        if df_huerfanos.empty:
            return False

        elementos.append(Paragraph("2. DOCUMENTOS FALTANTES (HUÉRFANOS)", self.estilo_subtitulo))
        elementos.append(Paragraph("Los siguientes registros existen en una plataforma pero no en la otra:", self.estilo_normal))

        datos_tabla = [["Origen", "Factura", "NIT", "Tercero", "Base", "Observación"]]
        
        for _, fila in df_huerfanos.sort_values(by=['Prioridad_Fila', 'Num.Ext']).iterrows():
            origen = "DIAN" if fila['Prioridad_Fila'] == 1 else "CONTAB."
            datos_tabla.append([
                origen,
                str(fila['Num.Ext']),
                str(fila['NIT Emisor']).replace('.0',''),
                Paragraph(str(fila['Nombre Emisor'])[:30], self.estilo_celda),
                f"$ {pd.to_numeric(fila['BASE'], errors='coerce'):,.2f}",
                Paragraph(str(fila['Novedad']), self.estilo_celda)
            ])

        tabla = Table(datos_tabla, colWidths=[50, 70, 70, 140, 70, 110], repeatRows=1)
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (4,1), (4,-1), 'RIGHT'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elementos.append(tabla)
        elementos.append(Spacer(1, 10))
        return True

    def _seccion_otros_errores(self, elementos):
        """Duplicados, Sin Num.Ext, Diferencias de Formato."""
        novedades_excluir = ['DIAN - Sin Pareja en Contabilidad', 'CONT - Sin Pareja en DIAN', 'DEVOL - Sin Pareja en DIAN', '']
        
        mask_otros = (~self.df['Novedad'].isin(novedades_excluir)) & \
                     (self.df['Novedad'].notna()) & \
                     (~self.df['Novedad'].str.contains('Sin Pareja', na=False)) & \
                     (~self.df['Novedad'].str.contains('Gasto Personal', na=False, case=False)) & \
                     (~self.df['Novedad'].str.contains('Application Response', na=False, case=False)) & \
                     (~self.df['Novedad'].str.startswith('DIAN:')) # Excluir personales de token

        df_otros = self.df[mask_otros]
        
        if df_otros.empty:
            return False

        elementos.append(Paragraph("3. DUPLICADOS Y ERRORES DE ESTRUCTURA", self.estilo_subtitulo))
        elementos.append(Paragraph("Se detectaron problemas en la digitación, documentos registrados más de una vez, o falta de número de factura:", self.estilo_normal))

        datos_tabla = [["Origen", "Factura Digitada", "Tercero", "Error Detectado"]]
        
        for _, fila in df_otros.sort_values(by=['Novedad', 'Prioridad_Fila']).iterrows():
            origen = "DIAN" if fila['Prioridad_Fila'] == 1 else "CONTAB."
            datos_tabla.append([
                origen,
                str(fila['Num.Ext_Original']),
                Paragraph(str(fila['Nombre Emisor'])[:35], self.estilo_celda),
                Paragraph(str(fila['Novedad']), self.estilo_celda)
            ])

        tabla = Table(datos_tabla, colWidths=[50, 100, 170, 190], repeatRows=1)
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elementos.append(tabla)
        return True