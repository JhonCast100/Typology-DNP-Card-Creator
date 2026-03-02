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
DNP_BLUE = HexColor('#003366') 
DNP_LIGHT_BLUE = HexColor('#e6f0ff')  
DNP_GREY = HexColor('#808080')  
TABLE_HEADER_BLUE = HexColor('#4a6fa5')  

class PdfGenerator:
    def __init__(self, filename):
        self.filename = filename
        self.styles = getSampleStyleSheet()
        self._setup_styles()
        
        # Rutas corregidas
        self.base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.logo_path = os.path.join(self.base_path, 'src', 'Views', 'Imgs', 'DNPLogo.png')
        self.map_path = os.path.join(self.base_path, 'Data', 'Maps', 'Casanare_1.jpg')
    
    def _setup_styles(self):
        """Configure custom styles exactly like the original PDF"""
        
        # Main title - centered, bold (más compacto)
        self.styles.add(ParagraphStyle(
            name='MainTitle',
            parent=self.styles['Heading1'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=4,
            leading=16,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        ))
        
        # Department title - centered, bold, blue (más compacto)
        self.styles.add(ParagraphStyle(
            name='DeptTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=8,
            leading=18,
            textColor=DNP_BLUE,
            fontName='Helvetica-Bold'
        ))
        
        # Section title - left aligned, bold, blue
        self.styles.add(ParagraphStyle(
            name='SectionTitle',
            parent=self.styles['Heading2'],
            fontSize=12,
            alignment=TA_LEFT,
            spaceAfter=6,
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
            spaceAfter=6,
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
            spaceAfter=4,
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
            spaceAfter=2,
            leading=12,
            textColor=colors.black,
            fontName='Helvetica',
            leftIndent=10,
            bulletIndent=5
        ))
        
        # Table title - bold, centered (cambiado a centrado)
        self.styles.add(ParagraphStyle(
            name='TableTitle',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_CENTER,  # Cambiado a centrado
            spaceAfter=2,
            leading=11,
            textColor=colors.black,
            fontName='Helvetica-Bold'
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
        
        # Source text - italic, centrado debajo de la tabla
        self.styles.add(ParagraphStyle(
            name='SourceText',
            parent=self.styles['Italic'],
            fontSize=8,
            alignment=TA_CENTER,
            textColor=DNP_GREY,
            leading=10,
            spaceAfter=2,
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
            topMargin=50,
            bottomMargin=90  # Aumentado para dar más espacio al footer de 3 líneas
        )
        
        elements = []
        
        # ==================== PAGE 1 ====================
        
        # Main title - más pegado
        title_text = """<font size=14><b>Tipologías de las Entidades Territoriales para el Reconocimiento de Capacidades. Resultados para la Vigencia 2025</b></font>"""
        title = Paragraph(title_text, self.styles["MainTitle"])
        elements.append(title)
        elements.append(Spacer(1, 0.1 * inch))
        
        # Department subtitle
        dept_text = f"""<font size=16><b><font color='#003366'>Departamento: {departmentName}</font></b></font>"""
        dept_title = Paragraph(dept_text, self.styles["DeptTitle"])
        elements.append(dept_title)
        elements.append(Spacer(1, 0.15 * inch))
        
        # Map title
        map_title = Paragraph("<b>Mapa de tipologías municipales</b>", self.styles["SubTitle"])
        elements.append(map_title)
        elements.append(Spacer(1, 0.1 * inch))
        
        # Map image
        try:
            if os.path.exists(self.map_path):
                mapa = Image(self.map_path, width=6*inch, height=7*inch)
                elements.append(mapa)
            else:
                print(f"⚠️ Mapa no encontrado en: {self.map_path}")
                elements.append(Paragraph("<i>Mapa de tipologías municipales</i>", self.styles["NormalText"]))
        except Exception as e:
            print(f"Error al cargar el mapa: {e}")
            elements.append(Paragraph("<i>Mapa de tipologías municipales</i>", self.styles["NormalText"]))
        
        elements.append(PageBreak())
        
        # ==================== PAGE 2 ====================
        
        # Municipal typologies 2025
        elements.append(Paragraph("Tipologías municipales 2025", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
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
            elements.append(Spacer(1, 0.02 * inch))
        
        elements.append(Spacer(1, 0.1 * inch))
        
        # Results table title (centrado)
        elements.append(Paragraph("Resultados de tipologías municipales y distritales - Departamento de Casanare", 
                                  self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
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
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(typology_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 3 ====================
        
        # Socioeconomic characteristics
        elements.append(Paragraph("Características socioeconómicas", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        socioeconomic_text = """
        Para hacer un mayor énfasis en las brechas sociales y económicas de los municipios del departamento, 
        se realiza un contraste con algunos indicadores territoriales de interés. Se observa el comportamiento 
        por Tipología del promedio del índice de pobreza multidimensional (IPM), del índice de necesidades 
        básicas insatisfechas (NBI), del índice de riesgo de la calidad del agua para consumo humano (IRCA) 
        y del índice de incidencia del conflicto armado (IICA), que en términos generales muestran una fuerte 
        relación con las capacidades fiscales y administrativas de los municipios y su situación geográfica.
        """
        elements.append(Paragraph(socioeconomic_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        socioeconomic_text2 = """
        Adicionalmente, se contrastan las Tipologías con la medición del desempeño municipal (MDM), haciendo 
        énfasis en el componente de resultados y en sus dimensiones: educación, salud, servicios públicos y 
        seguridad y convivencia.
        """
        elements.append(Paragraph(socioeconomic_text2, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Complementary variables table
        elements.append(Paragraph("Tipologías vs. variables complementarias", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
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
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en la información del DANE y DNP", self.styles["SourceText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # MDM table
        elements.append(Paragraph("Tipologías vs. MDM", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
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
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en información del DNP", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 4 ====================
        
        # Income section
        elements.append(Paragraph("Ingresos", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        income_text = "A medida que se avanza en las tipologías los ingresos propios en promedio tienden a reducirse."
        elements.append(Paragraph(income_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        elements.append(Paragraph("Ingresos propios 2023. Cifras en miles de millones de pesos", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
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
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en las operaciones efectivas de caja calculadas por el DNP", 
                                  self.styles["SourceText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Environmental areas section
        elements.append(Paragraph("Áreas ambientales y territorios étnicos", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        environmental_text = """
        La variable porcentaje de "Áreas protegidas y ecosistemas estratégicos" resulta del cruce de información 
        del Registro Único Nacional de Áreas Protegidas (RUNAP)<super>1</super> y del Registro de Ecosistemas y Áreas 
        Ambientales (REAA)<super>2</super>, en relación con el total del área municipal.
        """
        elements.append(Paragraph(environmental_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Protected areas table
        elements.append(Paragraph("Tipologías vs. Rango de % de áreas protegidas y ecosistemas estratégicos", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        protected_areas_data = [
            ["Tipología", "% Área de protección ambiental y ecosistemas estratégicos", "", "", "", "Total"],
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
        
        protected_areas_table = Table(protected_areas_data, colWidths=[2.5*cm, 2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm])
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
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
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
        elements.append(Spacer(1, 0.1 * inch))
        
        # Ethnic territories table
        elements.append(Paragraph("Tipologías vs. Rango de % de área de territorios étnicos", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de Casanare", self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        ethnic_territories_data = [
            ["Tipología", "% Área de Territorios Étnicos", "", "", "", "", "Total"],
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
        
        ethnic_territories_table = Table(ethnic_territories_data, colWidths=[2.5*cm, 2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm])
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
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia", self.styles["SourceText"]))
        elements.append(PageBreak())
        
        # ==================== PAGE 7 ====================
        
        # Annex
        elements.append(Paragraph("Anexo", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Tipologías a nivel municipal- Departamento de Casanare", self.styles["SubTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Municipality data for annex
        annex_data = [
            ["Código DANE", "Departamento", "Municipio", "Tipología 2025"]
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
        elements.append(Spacer(1, 0.2 * inch))
        
        # Footnotes
        footnote1 = """
        <font size=6>
        <super>1</super> Dentro del registro se incluyen los Parques Naturales Regionales y las categorías del Sistema de Áreas de Parques Nacionales con base en el Decreto 2811 de 1974 (Artículo 329)
        </font>
        """
        elements.append(Paragraph(footnote1, self.styles["Footnote"]))
        elements.append(Spacer(1, 0.02 * inch))
        
        footnote2 = """
        <font size=6>
        <super>2</super> Se incluyen todas las categorías de esta fuente de información, a saber: Páramos, Humedales RAMSAR, Bosque Seco Tropical, Manglares, Pastos Marinos, Arrecifes coralinos, Reservas Forestales de Ley 2 de 1959 (Zona Tipo A), Áreas Susceptibles a Procesos de Restauración Ecológica, Áreas de proyectos Bosques de Paz orientados a la restauración ambiental y reconciliación de víctimas
        </font>
        """
        elements.append(Paragraph(footnote2, self.styles["Footnote"]))
        
        # Build PDF with custom header and footer on each page
        doc.build(elements, onFirstPage=self._header_footer, onLaterPages=self._header_footer)
    
    def _header_footer(self, canvas, doc):
        """Add header and footer to each page"""
        canvas.saveState()
        
        # Calcular posición para el logo (centrado)
        page_width = 612  # letter width in points (8.5 * 72)
        logo_width = 120
        logo_x = (page_width - logo_width) / 2
        
        # Header with logo (centrado) - BAJADO UN POCO PARA QUE NO SE CORTE
        try:
            if os.path.exists(self.logo_path):
                # Bajé la posición Y de 760 a 750 para dar más espacio arriba
                canvas.drawImage(self.logo_path, logo_x, 750, width=logo_width, height=35, preserveAspectRatio=True, mask='auto')
            else:
                print(f"⚠️ Logo no encontrado en: {self.logo_path}")
        except Exception as e:
            print(f"Error al cargar el logo: {e}")
        
        # Header line (blue) - centrada
        line_width = 530
        line_start = (page_width - line_width) / 2
        canvas.setStrokeColor(DNP_BLUE)
        canvas.setLineWidth(2)
        canvas.line(line_start, 740, line_start + line_width, 740)  # Línea también bajada
        
        # Footer line (blue) - centrada
        canvas.setStrokeColor(DNP_BLUE)
        canvas.setLineWidth(1)
        canvas.line(line_start, 80, line_start + line_width, 80)  # Línea del footer más arriba para dar espacio
        
        # Footer text - ALINEADO A LA IZQUIERDA
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(DNP_GREY)
        
        # Primera línea del footer - alineada a la izquierda en lugar de centrada
        canvas.drawString(line_start, 65, "Dirección: Calle 26 # 13 – 19 Bogotá, D.C., Colombia")
        
        # Segunda línea: Conmutador - alineada a la izquierda
        canvas.drawString(line_start, 55, "Conmutador: 601 3815000")
        
        # Tercera línea: Línea gratuita - alineada a la izquierda
        canvas.drawString(line_start, 45, "Línea gratuita: PBX 381 5000")
        
        # Page number (right side) - alineado a la derecha pero dentro del margen
        page_num = getattr(doc, 'page', 1)
        canvas.drawRightString(line_start + line_width, 45, f"Página {page_num}")
        
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
    print(f"\n📋 Rutas configuradas:")
    print(f"  - Logo: {generator.logo_path}")
    print(f"  - Mapa: {generator.map_path}")