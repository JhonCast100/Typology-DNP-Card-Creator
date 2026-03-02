from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
import os

# Colores exactos del DNP
DNP_BLUE = HexColor('#003366')  # Azul oscuro del DNP
DNP_LIGHT_BLUE = HexColor('#e6f0ff')  # Azul muy claro para fondos
DNP_GREY = HexColor('#808080')  # Gris para textos secundarios
TABLE_HEADER_BLUE = HexColor('#4a6fa5')  # Azul para encabezados de tabla

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._page_number = 0
        self._total_pages = None

    def showPage(self):
        self._page_number += 1
        canvas.Canvas.showPage(self)

    def save(self):
        self._total_pages = self._page_number
        self._page_number = 0
        canvas.Canvas.save(self)

    def draw_header_footer(self):
        page_num = self._page_number
        total_pages = self._total_pages or 1
        
        # Header with logo (top left)
        try:
            logo_path = "dnp_logo.png"
            if os.path.exists(logo_path):
                self.drawImage(logo_path, 40, 740, width=120, height=35, preserveAspectRatio=True, mask='auto')
        except:
            pass
        
        # Header line (blue)
        self.setStrokeColor(DNP_BLUE)
        self.setLineWidth(2)
        self.line(40, 730, 570, 730)
        
        # Footer line (blue)
        self.setStrokeColor(DNP_BLUE)
        self.setLineWidth(1)
        self.line(40, 60, 570, 60)
        
        # Footer text - exactly as in original
        self.setFont("Helvetica", 7)
        self.setFillColor(DNP_GREY)
        
        # First line of footer
        self.drawString(40, 45, "Dirección: Calle 26 # 13 – 19 Bogotá, D.C., Colombia")
        
        # Second line with two parts
        self.drawString(40, 35, "Conmutador: 601 3815000")
        self.drawString(180, 35, "Línea gratuita: PBX 381 5000")
        
        # Page number (right side)
        self.setFont("Helvetica", 7)
        self.setFillColor(DNP_GREY)
        self.drawRightString(570, 35, f"Página {page_num} de {total_pages}")

    def drawPage(self):
        self.draw_header_footer()
        canvas.Canvas.drawPage(self)


class PdfGenerator:
    def __init__(self, filename):
        self.filename = filename
        self.styles = getSampleStyleSheet()
        self._setup_styles()
    
    def _setup_styles(self):
        """Configure custom styles exactly like the original PDF"""
        
        # Main title - centered, bold
        self.styles.add(ParagraphStyle(
            name='MainTitle',
            parent=self.styles['Heading1'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=8,
            leading=18,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        ))
        
        # Department title - centered, bold, blue
        self.styles.add(ParagraphStyle(
            name='DeptTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=15,
            leading=20,
            textColor=DNP_BLUE,
            fontName='Helvetica-Bold'
        ))
        
        # Section title - left aligned, bold, blue
        self.styles.add(ParagraphStyle(
            name='SectionTitle',
            parent=self.styles['Heading2'],
            fontSize=12,
            alignment=TA_LEFT,
            spaceAfter=8,
            leading=14,
            textColor=DNP_BLUE,
            fontName='Helvetica-Bold',
            leftIndent=0
        ))
        
        # Subtitle - centered, bold
        self.styles.add(ParagraphStyle(
            name='SubTitle',
            parent=self.styles['Heading2'],
            fontSize=11,
            alignment=TA_CENTER,
            spaceAfter=10,
            leading=13,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        ))
        
        # Normal text - justified
        self.styles.add(ParagraphStyle(
            name='NormalText',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
            leading=11,
            textColor=colors.black,
            fontName='Helvetica'
        ))
        
        # List text - left aligned with bullet
        self.styles.add(ParagraphStyle(
            name='ListText',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_LEFT,
            spaceAfter=4,
            leading=12,
            textColor=colors.black,
            fontName='Helvetica',
            leftIndent=10,
            bulletIndent=5
        ))
        
        # Table header style
        self.styles.add(ParagraphStyle(
            name='TableHeader',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_CENTER,
            textColor=colors.white,
            leading=12,
            fontName='Helvetica-Bold'
        ))
        
        # Table cell style
        self.styles.add(ParagraphStyle(
            name='TableCell',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_CENTER,
            textColor=colors.black,
            leading=11,
            fontName='Helvetica'
        ))
        
        # Table cell left aligned
        self.styles.add(ParagraphStyle(
            name='TableCellLeft',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_LEFT,
            textColor=colors.black,
            leading=11,
            fontName='Helvetica'
        ))
        
        # Source text - italic
        self.styles.add(ParagraphStyle(
            name='SourceText',
            parent=self.styles['Italic'],
            fontSize=8,
            alignment=TA_LEFT,
            textColor=DNP_GREY,
            leading=10,
            fontName='Helvetica-Oblique'
        ))
        
        # Footnote text
        self.styles.add(ParagraphStyle(
            name='Footnote',
            parent=self.styles['Italic'],
            fontSize=6,
            alignment=TA_LEFT,
            textColor=DNP_GREY,
            leading=8,
            fontName='Helvetica-Oblique'
        ))
    
    def generatePdf(self, departmentName, data):
        """
        Generates the PDF exactly matching the original document
        """
        doc = SimpleDocTemplate(
            self.filename,
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=80,
            bottomMargin=80
        )
        
        elements = []
        
        # ==================== PAGE 1 ====================
        
        # Main title
        title_text = """<font size=14><b>Tipologías de las Entidades Territoriales para el Reconocimiento de Capacidades. Resultados para la Vigencia 2026</b></font>"""
        title = Paragraph(title_text, self.styles["MainTitle"])
        elements.append(title)
        elements.append(Spacer(1, 0.3 * inch))
        
        # Department subtitle
        dept_text = f"""<font size=16><b><font color='#003366'>Departamento: {departmentName}</font></b></font>"""
        dept_title = Paragraph(dept_text, self.styles["DeptTitle"])
        elements.append(dept_title)
        elements.append(Spacer(1, 0.5 * inch))
        
        # Map title
        map_title = Paragraph("<b>Mapa de tipologías municipales</b>", self.styles["SubTitle"])
        elements.append(map_title)
        elements.append(Spacer(1, 0.2 * inch))
        
        # Map image
        try:
            map_path = "mapa_casanare.jpg"
            if os.path.exists(map_path):
                mapa = Image(map_path, width=6*inch, height=7*inch)
                elements.append(mapa)
            else:
                elements.append(Paragraph("<i>Mapa de tipologías municipales</i>", self.styles["NormalText"]))
        except:
            elements.append(Paragraph("<i>Mapa de tipologías municipales</i>", self.styles["NormalText"]))
        
        elements.append(PageBreak())
        
        # ==================== PAGE 2 ====================
        
        # Municipal typologies
        elements.append(Paragraph("Tipologías municipales 2026", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Typologies description with bullet points
        typologies_items = [
            f"• El departamento de <b>{departmentName}</b> está conformado por <b>{data['total_municipalities']}</b> municipios.",
            "• Ningún municipio de este departamento pertenece a las Tipologías de Ciudades Grandes o Centros de Aglomeración - Sistema de Ciudades.",
            "• En la <b>Tipología 1</b> no se encuentra ningún municipio de este departamento. Los municipios de esta tipología se caracterizan por tener en promedio los más altos niveles de capacidades fiscales y administrativas, y al mismo tiempo son los municipios mejores conectados y más densos.",
            f"• En la <b>Tipología 2</b> se encuentra <b>{data['typology2']['quantity']}</b> municipio{'' if data['typology2']['quantity'] == 1 else 's'}, <b>{data['typology2']['municipalities']}</b>. Los municipios de esta tipología se caracterizan porque tienen niveles intermedios-altos de capacidad fiscal y administrativa, conectividad y densidad poblacional.",
            f"• En la <b>Tipología 3</b> se encuentran <b>{data['typology3']['quantity']}</b> municipios. Los municipios de esta tipología se caracterizan por tener niveles intermedios de capacidad fiscal y administrativa, conectividad y densidad poblacional.",
            f"• En la <b>Tipología 4</b> se encuentran <b>{data['typology4']['quantity']}</b> municipios. Los municipios de esta tipología se caracterizan por tener niveles intermedios-bajos de capacidad administrativa y fiscal conectividad y densidad poblacional.",
            f"• En la <b>Tipología 5</b> se encuentran <b>{data['typology5']['quantity']}</b> municipios. Los municipios de esta tipología se caracterizan por tener bajos niveles de capacidad administrativa y fiscal, al mismo tiempo son los más desconectados y menos densos (mayor ruralidad)."
        ]
        
        for item in typologies_items:
            elements.append(Paragraph(item, self.styles["ListText"]))
            elements.append(Spacer(1, 0.05 * inch))
        
        elements.append(Spacer(1, 0.2 * inch))
        
        # Results table title
        elements.append(Paragraph("Resultados de tipologías municipales y distritales - Departamento de Casanare", 
                                  self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Typologies results table
        typology_table_data = [
            ["Tipología", "Núm. municipios y distritos", "Índice de capacidades\nponderado (Promedio)"],
            ["Bogotá", "0", "-"],
            ["Ciudades grandes", "0", "-"],
            ["SC- Centro aglomeración", "0", "-"],
            ["1", "0", "-"],
            ["2", str(data['typology2']['quantity']), "0,53"],
            ["3", str(data['typology3']['quantity']), "0,43"],
            ["4", str(data['typology4']['quantity']), "0,29"],
            ["5", str(data['typology5']['quantity']), "0,10"],
            ["Total", str(data['total_municipalities']), "0,25"]
        ]
        
        typology_table = Table(typology_table_data, colWidths=[4*cm, 5*cm, 6*cm])
        typology_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(typology_table)
        elements.append(Spacer(1, 0.2 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 3 ====================
        
        # Socioeconomic characteristics
        elements.append(Paragraph("Características socioeconómicas", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        socioeconomic_text = """
        Para hacer un mayor énfasis en las brechas sociales y económicas de los municipios del departamento, 
        se realiza un contraste con algunos indicadores territoriales de interés. Se observa el comportamiento 
        por Tipología del promedio del índice de pobreza multidimensional (IPM), del índice de necesidades 
        básicas insatisfechas (NBI), del índice de riesgo de la calidad del agua para consumo humano (IRCA) 
        y del índice de incidencia del conflicto armado (IICA), que en términos generales muestran una fuerte 
        relación con las capacidades fiscales y administrativas de los municipios y su situación geográfica.
        """
        elements.append(Paragraph(socioeconomic_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        socioeconomic_text2 = """
        Adicionalmente, se contrastan las Tipologías con la medición del desempeño municipal (MDM), haciendo 
        énfasis en el componente de resultados y en sus dimensiones: educación, salud, servicios públicos y 
        seguridad y convivencia.
        """
        elements.append(Paragraph(socioeconomic_text2, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # Complementary variables table
        elements.append(Paragraph("Tipologías vs. variables complementarias", self.styles["NormalText"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        complementary_table_data = [
            ["Tipología", "IPM", "NBI", "IRCA", "IICA"],
            ["Bogotá", "-", "-", "-", "-"],
            ["Ciudades grandes", "-", "-", "-", "-"],
            ["SC- Centro aglomeración", "-", "-", "-", "-"],
            ["1", "-", "-", "-", "-"],
            ["2", "22,90", "11,53", "0,60", "1,36"],
            ["3", "28,10", "12,28", "2,56", "7,57"],
            ["4", "44,23", "26,67", "6,94", "1,06"],
            ["5", "42,37", "25,93", "8,11", "4,49"],
            ["Total", "37,98", "21,73", "6,01", "4,42"]
        ]
        
        complementary_table = Table(complementary_table_data, colWidths=[3.5*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.8*cm])
        complementary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(complementary_table)
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en la información del DANE y DNP", self.styles["SourceText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # MDM table
        elements.append(Paragraph("Tipologías vs. MDM", self.styles["NormalText"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        mdm_table_data = [
            ["Tipología", "MDM", "MDM\nResultados", "Educación", "Salud", "Servicios", "Seguridad y\nconvivencia"],
            ["Bogotá", "-", "-", "-", "-", "-", "-"],
            ["Ciudades grandes", "-", "-", "-", "-", "-", "-"],
            ["SC- Centro aglomeración", "-", "-", "-", "-", "-", "-"],
            ["1", "-", "-", "-", "-", "-", "-"],
            ["2", "72,42", "75,97", "64,00", "96,14", "70,19", "73,54"],
            ["3", "65,14", "68,89", "55,43", "90,60", "38,89", "90,62"],
            ["4", "52,28", "61,52", "49,84", "86,88", "28,32", "69,34"],
            ["5", "57,60", "64,21", "49,29", "88,05", "31,55", "82,70"],
            ["Total", "59,24", "65,49", "51,80", "88,90", "34,84", "81,49"]
        ]
        
        mdm_table = Table(mdm_table_data, colWidths=[2.8*cm, 2.2*cm, 2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 3*cm])
        mdm_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(mdm_table)
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en información del DNP", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 4 ====================
        
        # Income section
        elements.append(Paragraph("Ingresos", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        income_text = "A medida que se avanza en las tipologías los ingresos propios en promedio tienden a reducirse."
        elements.append(Paragraph(income_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        elements.append(Paragraph("Ingresos propios 2023. Cifras en miles de millones de pesos", self.styles["NormalText"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        income_table_data = [
            ["Tipología", "Ingresos tributarios"],
            ["Bogotá", "-"],
            ["Ciudades grandes", "-"],
            ["SC- Centro aglomeración", "-"],
            ["1", "-"],
            ["2", "168.236"],
            ["3", "34.324"],
            ["4", "4.583"],
            ["5", "8.037"],
            ["Total", "22.659"]
        ]
        
        income_table = Table(income_table_data, colWidths=[5*cm, 6*cm])
        income_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(income_table)
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en las operaciones efectivas de caja calculadas por el DNP", 
                                  self.styles["SourceText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # Environmental areas section
        elements.append(Paragraph("Áreas ambientales y territorios étnicos", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        environmental_text = """
        La variable porcentaje de "Áreas protegidas y ecosistemas estratégicos" resulta del cruce de información 
        del Registro Único Nacional de Áreas Protegidas (RUNAP)<super>1</super> y del Registro de Ecosistemas y Áreas 
        Ambientales (REAA)<super>2</super>, en relación con el total del área municipal.
        """
        elements.append(Paragraph(environmental_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # Protected areas table
        elements.append(Paragraph("Tipologías vs. Rango de % de áreas protegidas y ecosistemas estratégicos", self.styles["NormalText"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        protected_areas_data = [
            ["Tipología", "", "% Área de protección ambiental y ecosistemas estratégicos", "", "", "Total"],
            ["", "0% - 30%", "31% - 50%", "51% - 70%", "71% - 100%", ""],
            ["Bogotá", "0", "0", "0", "0", "0"],
            ["Ciudades grandes", "0", "0", "0", "0", "0"],
            ["SC- Centro aglomeración", "0", "0", "0", "0", "0"],
            ["1", "0", "0", "0", "0", "0"],
            ["2", "1", "0", "0", "0", "1"],
            ["3", "5", "0", "0", "0", "5"],
            ["4", "4", "0", "0", "0", "4"],
            ["5", "8", "1", "0", "0", "9"],
            ["Total", "18", "1", "0", "0", "19"]
        ]
        
        protected_areas_table = Table(protected_areas_data, colWidths=[2.5*cm, 2.2*cm, 2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm])
        protected_areas_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 1), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 1), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('SPAN', (0, 0), (0, 1)),
            ('SPAN', (1, 0), (4, 0)),
            ('SPAN', (5, 0), (5, 1)),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (1, 0), (4, 0), 2, colors.black),
        ]))
        
        elements.append(protected_areas_table)
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        elements.append(PageBreak())
        
        # ==================== PAGE 5 ====================
        
        # Ethnic territories text
        ethnic_text = """
        La variable de porcentaje de "Área de territorios étnicos" abarca las áreas de los resguardos indígenas 
        legalmente constituidos y las tierras colectivas tituladas a las comunidades negras, afro, raizales y 
        palenqueras de acuerdo con la información oficial de la Agencia Nacional de Tierras, en relación con el 
        total del área municipal.
        """
        elements.append(Paragraph(ethnic_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # Ethnic territories table
        elements.append(Paragraph("Tipologías vs. Rango de % de área de territorios étnicos", self.styles["NormalText"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        ethnic_territories_data = [
            ["Tipología", "", "% Área de Territorios Étnicos", "", "", "", "Total"],
            ["", "0%", "0,01% - 30%", "31% - 50%", "51% - 70%", "71% - 100%", ""],
            ["Bogotá", "0", "0", "0", "0", "0", "0"],
            ["Ciudades grandes", "0", "0", "0", "0", "0", "0"],
            ["SC- Centro aglomeración", "0", "0", "0", "0", "0", "0"],
            ["1", "0", "0", "0", "0", "0", "0"],
            ["2", "0", "1", "0", "0", "0", "1"],
            ["3", "4", "1", "0", "0", "0", "5"],
            ["4", "2", "2", "0", "0", "0", "4"],
            ["5", "6", "3", "0", "0", "0", "9"],
            ["Total", "12", "7", "0", "0", "0", "19"]
        ]
        
        ethnic_territories_table = Table(ethnic_territories_data, colWidths=[2.5*cm, 1.8*cm, 2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm])
        ethnic_territories_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 1), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 1), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('SPAN', (0, 0), (0, 1)),
            ('SPAN', (1, 0), (5, 0)),
            ('SPAN', (6, 0), (6, 1)),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (1, 0), (5, 0), 2, colors.black),
        ]))
        
        elements.append(ethnic_territories_table)
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 7 ====================
        
        # Annex
        elements.append(Paragraph("Anexo", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.2 * inch))
        elements.append(Paragraph("Tipologías a nivel municipal- Departamento de Casanare", self.styles["SubTitle"]))
        elements.append(Spacer(1, 0.2 * inch))
        
        # Municipality data for annex
        annex_data = [
            ["Código DANE", "Departamento", "Municipio", "Tipología 2026"]
        ]
        
        for municipality in data.get('municipalities', []):
            annex_data.append([
                municipality['code'],
                municipality['department'],
                municipality['name'],
                str(municipality['typology'])
            ])
        
        annex_table = Table(annex_data, colWidths=[3*cm, 3.5*cm, 4.5*cm, 3*cm])
        annex_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(annex_table)
        elements.append(Spacer(1, 0.3 * inch))
        
        # Footnotes
        footnote1 = """
        <font size=6>
        <super>1</super> Dentro del registro se incluyen los Parques Naturales Regionales y las categorías del Sistema de Áreas de Parques Nacionales con base en el Decreto 2811 de 1974 (Artículo 329)
        </font>
        """
        elements.append(Paragraph(footnote1, self.styles["Footnote"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        footnote2 = """
        <font size=6>
        <super>2</super> Se incluyen todas las categorías de esta fuente de información, a saber: Páramos, Humedales RAMSAR, Bosque Seco Tropical, Manglares, Pastos Marinos, Arrecifes coralinos, Reservas Forestales de Ley 2 de 1959 (Zona Tipo A), Áreas Susceptibles a Procesos de Restauración Ecológica, Áreas de proyectos Bosques de Paz orientados a la restauración ambiental y reconciliación de víctimas
        </font>
        """
        elements.append(Paragraph(footnote2, self.styles["Footnote"]))
        
        # Build PDF with custom canvas for header and footer
        doc.build(elements, onFirstPage=self._header_footer, onLaterPages=self._header_footer)
    
    def _header_footer(self, canvas, doc):
        """Add header and footer to each page"""
        canvas.saveState()
        
        # Header with logo (top left)
        try:
            logo_path = "dnp_logo.png"
            if os.path.exists(logo_path):
                canvas.drawImage(logo_path, 40, 740, width=120, height=35, preserveAspectRatio=True, mask='auto')
        except:
            pass
        
        # Header line (blue)
        canvas.setStrokeColor(DNP_BLUE)
        canvas.setLineWidth(2)
        canvas.line(40, 730, 570, 730)
        
        # Footer line (blue)
        canvas.setStrokeColor(DNP_BLUE)
        canvas.setLineWidth(1)
        canvas.line(40, 60, 570, 60)
        
        # Footer text - exactly as in original
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(DNP_GREY)
        
        # First line of footer
        canvas.drawString(40, 45, "Dirección: Calle 26 # 13 – 19 Bogotá, D.C., Colombia")
        
        # Second line with two parts
        canvas.drawString(40, 35, "Conmutador: 601 3815000")
        canvas.drawString(180, 35, "Línea gratuita: PBX 381 5000")
        
        # Page number (right side)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(DNP_GREY)
        canvas.drawRightString(570, 35, f"Página {doc.page} de {len(doc.pages) + 1 if hasattr(doc, 'pages') else 1}")
        
        canvas.restoreState()


# Example usage
if __name__ == "__main__":
    # Data based on the Casanare PDF
    data_casanare = {
        'total_municipalities': 19,
        'typology2': {
            'quantity': 1,
            'municipalities': 'Yopal'
        },
        'typology3': {
            'quantity': 5,
            'municipalities': 'Villanueva, Aguazul, Monterrey, Tauramena, Sabanalarga'
        },
        'typology4': {
            'quantity': 4,
            'municipalities': 'Pore, Recetor, Nunchía, Chámeza'
        },
        'typology5': {
            'quantity': 9,
            'municipalities': 'Sácama, La Salina, Támara, Maní, Trinidad, Paz de Ariporo, San Luis de Palenque, Hato Corozal, Orocué'
        },
        'municipalities': [
            {'code': '85001', 'department': 'CASANARE', 'name': 'YOPAL', 'typology': 2},
            {'code': '85440', 'department': 'CASANARE', 'name': 'VILLANUEVA', 'typology': 3},
            {'code': '85010', 'department': 'CASANARE', 'name': 'AGUAZUL', 'typology': 3},
            {'code': '85162', 'department': 'CASANARE', 'name': 'MONTERREY', 'typology': 3},
            {'code': '85410', 'department': 'CASANARE', 'name': 'TAURAMENA', 'typology': 3},
            {'code': '85300', 'department': 'CASANARE', 'name': 'SABANALARGA', 'typology': 3},
            {'code': '85263', 'department': 'CASANARE', 'name': 'PORE', 'typology': 4},
            {'code': '85279', 'department': 'CASANARE', 'name': 'RECETOR', 'typology': 4},
            {'code': '85225', 'department': 'CASANARE', 'name': 'NUNCHÍA', 'typology': 4},
            {'code': '85015', 'department': 'CASANARE', 'name': 'CHÁMEZA', 'typology': 4},
            {'code': '85315', 'department': 'CASANARE', 'name': 'SÁCAMA', 'typology': 5},
            {'code': '85136', 'department': 'CASANARE', 'name': 'LA SALINA', 'typology': 5},
            {'code': '85400', 'department': 'CASANARE', 'name': 'TÁMARA', 'typology': 5},
            {'code': '85139', 'department': 'CASANARE', 'name': 'MANÍ', 'typology': 5},
            {'code': '85430', 'department': 'CASANARE', 'name': 'TRINIDAD', 'typology': 5},
            {'code': '85250', 'department': 'CASANARE', 'name': 'PAZ DE ARIPORO', 'typology': 5},
            {'code': '85325', 'department': 'CASANARE', 'name': 'SAN LUIS DE PALENQUE', 'typology': 5},
            {'code': '85125', 'department': 'CASANARE', 'name': 'HATO COROZAL', 'typology': 5},
            {'code': '85230', 'department': 'CASANARE', 'name': 'OROCUÉ', 'typology': 5}
        ]
    }
    
    # Create generator instance and produce PDF
    generator = PdfGenerator("reporte_casanare_final.pdf")
    generator.generatePdf("Casanare", data_casanare)
    print("✅ PDF generado exitosamente: reporte_casanare_final.pdf")
