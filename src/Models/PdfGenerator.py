from turtle import color

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor


import pandas as pd
import os

# Definition of DNP colors
DNP_BLUE = HexColor('#003366') 
DNP_BLACK = HexColor("#000000") 
DNP_LIGHT_BLUE = HexColor('#e6f0ff')  
DNP_GREY = HexColor('#808080')  
TABLE_HEADER_BLUE = HexColor('#4a6fa5')  

class PdfGenerator:
    
    
    def __init__(self, filename):

        self.filename = filename
        self.styles = getSampleStyleSheet()
        self._setup_styles()

        #Base path
        self.base_path = os.path.dirname(
            os.path.dirname(
                os.path.dirname(os.path.abspath(__file__))
            )
        )
        self.logo_path = os.path.join(
            self.base_path, 'src', 'Views', 'Imgs', 'DNPLogo.png'
        )

        self.output_path = os.path.join(self.base_path, "Output")

        #Output directory
        os.makedirs(self.output_path, exist_ok=True)
    
    def _setup_styles(self):
        """Configure custom styles exactly like the original PDF"""
        
        # Main title
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
        
        # Department title
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
        
        # Section title
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
        
        # Subtitle
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
        
        # Normal text
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
        
        # List text
        self.styles.add(ParagraphStyle(
            name='ListText',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_JUSTIFY,   
            spaceAfter=4,
            leading=12,
            textColor=colors.black,
            fontName='Helvetica',

            leftIndent=18,        
            firstLineIndent=-12
        ))
        
        # Table footnote style
        self.styles.add(ParagraphStyle(
            name='TableFootnote',
            parent=self.styles['Normal'],
            alignment=TA_JUSTIFY, 
            fontSize=7,          # más pequeño
            leading=9,
            textColor=colors.black,
            leftIndent=0,
            spaceBefore=4,
            spaceAfter=2,
        ))
        
        # Table title
        self.styles.add(ParagraphStyle(
            name='TableTitle',
            parent=self.styles['Normal'],
            fontSize=9,
            alignment=TA_CENTER,  
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
        
        # Source text
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
        
        self.ranking_colors = [
            colors.HexColor("#C6E8C6"),  
            colors.HexColor("#E2F3D5"),  
            colors.HexColor("#FFF4C2"), 
            colors.HexColor("#FFD9A8"),  
            colors.HexColor("#FFBFA8"),  
            colors.HexColor("#F5A6A6")  
        ]
    
    def generatePdf(self, departmentName, data):
        
        # Format department name: first letter uppercase, rest lowercase
        departmentName = departmentName.title()
        
        elements = []
        
        # Total municipalities
        total_municipalities = len(data)

        # Count of municipalities by typology
        typology_counts = data["Tipologia_2026R"].value_counts()

        # Mean ICPond by typology
        typology_means = (
            data
            .groupby("Tipologia_2026R")["ICPond_2026"]
            .mean()
        )
        
        def get_map_info(department_name):
            formatted_name = department_name.lower().capitalize()

            maps_folder = os.path.join(self.base_path, 'Data', 'Maps')

            path_with_names = os.path.join(maps_folder, f"{formatted_name}_1.jpg")
            path_without_names = os.path.join(maps_folder, f"{formatted_name}_0.jpg")

            if os.path.exists(path_with_names):
                return path_with_names, True
            elif os.path.exists(path_without_names):
                return path_without_names, False
            else:
                return None, None
                
        
        """
        Generates the PDF exactly matching the original document
        """
        pdf_file_path = os.path.join(self.output_path, f"{departmentName}.pdf")

        
        doc = SimpleDocTemplate(
            pdf_file_path,
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=50,
            bottomMargin=90
        )
        

        
        # ==================== PAGE 1 ====================
        
        # Main title
        title_text = """<font size=14><b>Tipologías de las Entidades Territoriales para el Reconocimiento de Capacidades.</b></font>"""
        title_text2 = """<font size=14><b>Resultados para la Vigencia 2026</b></font>"""
        
        title = Paragraph(title_text, self.styles["MainTitle"])
        elements.append(title)
        title = Paragraph(title_text2, self.styles["MainTitle"])
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
        map_path, has_names = get_map_info(departmentName)

        if map_path:
            mapa = Image(map_path, width=5*inch, height=6*inch)
            elements.append(mapa)
            elements.append(Spacer(1, 0.05 * inch))
            elements.append(Paragraph("Fuente: Elaboración propia con base en información del DNP", self.styles["SourceText"]))
            elements.append(PageBreak())

            if not has_names:
                elements.append(Spacer(1, 0.2 * inch))

                data = data.copy()
                data["Cod_Municipio"] = data["CodDANE_txt"].astype(str).str[-3:]

                data = data.sort_values("Municipio").reset_index(drop=True)

                table_data = [["Cod", "Municipio", "Cod", "Municipio", "Cod", "Municipio"]]

                total = len(data)

                block_size = total // 3 + (1 if total % 3 > 0 else 0)

                left_data = data.iloc[:block_size]
                middle_data = data.iloc[block_size:block_size*2].reset_index(drop=True)
                right_data = data.iloc[block_size*2:].reset_index(drop=True)

                # Build table rows
                for i in range(block_size):

                    # LEFT
                    if i < len(left_data):
                        left_cod = left_data.iloc[i]["Cod_Municipio"]
                        left_mun = left_data.iloc[i]["Municipio"]
                    else:
                        left_cod, left_mun = "", ""

                    # MIDDLE
                    if i < len(middle_data):
                        mid_cod = middle_data.iloc[i]["Cod_Municipio"]
                        mid_mun = middle_data.iloc[i]["Municipio"]
                    else:
                        mid_cod, mid_mun = "", ""

                    # RIGHT
                    if i < len(right_data):
                        right_cod = right_data.iloc[i]["Cod_Municipio"]
                        right_mun = right_data.iloc[i]["Municipio"]
                    else:
                        right_cod, right_mun = "", ""

                    table_data.append([
                        left_cod, left_mun,
                        mid_cod, mid_mun,
                        right_cod, right_mun
                    ])

                # Create Table
                cod_table = Table(
                    table_data,
                    colWidths=[1.5*cm, 3.5*cm,
                            1.5*cm, 3.5*cm,
                            1.5*cm, 3.5*cm],
                    repeatRows=1
                )

                cod_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('GRID', (0, 0), (-1, -1), 0.4, colors.grey),
                    ('FONTSIZE', (0, 0), (-1, -1), 6),
                ]))

                elements.append(cod_table)
                
                
        else:
            elements.append(Paragraph("<i>Mapa no disponible</i>", self.styles["NormalText"]))
        
        
        # ==================== PAGE 2 ====================

        elements.append(Paragraph("Tipologías municipales 2026", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))

        typology_counts = data["Tipología_2026_CortesArcMap"].value_counts()

        analisis_tipologias = {
            1: "Los municipios de esta tipología se caracterizan por tener en promedio los más altos niveles de capacidades fiscales y administrativas, y al mismo tiempo son los municipios mejor conectados y más densos.",
            2: "Los municipios de esta tipología se caracterizan porque tienen niveles intermedios-altos de capacidad fiscal y administrativa, conectividad y densidad poblacional.",
            3: "Los municipios de esta tipología se caracterizan por tener niveles intermedios de capacidad fiscal y administrativa, conectividad y densidad poblacional.",
            4: "Los municipios de esta tipología se caracterizan por tener niveles intermedios-bajos de capacidad administrativa y fiscal, conectividad y densidad poblacional.",
            5: "Los municipios de esta tipología se caracterizan por tener bajos niveles de capacidad administrativa y fiscal; al mismo tiempo son los más desconectados y menos densos (mayor ruralidad)."
        }

        #First bullet point with total municipalities
        elements.append(Paragraph(
            f"• El departamento de <b>{departmentName}</b> está conformado por <b>{total_municipalities}</b> municipios.",
            self.styles["ListText"]
        ))
        elements.append(Spacer(1, 0.05 * inch))

        for t in [1, 2, 3, 4, 5]:

            cantidad = typology_counts.get(t, 0)

            if cantidad == 0:
                texto = (
                    f"• En la <b>Tipología {t}</b> no se encuentra ningún municipio de este departamento. "
                    f"{analisis_tipologias[t]}"
                )

            elif cantidad == 1:
                municipio = data.loc[data["Tipologia_2026R"] == t, "Municipio"].iloc[0]
                texto = (
                    f"• En la <b>Tipología {t}</b> se encuentra <b>1 municipio</b>, <b>{municipio}</b>. "
                    f"{analisis_tipologias[t]}"
                )

            else:
                texto = (
                    f"• En la <b>Tipología {t}</b> se encuentran <b>{cantidad} municipios</b>. "
                    f"{analisis_tipologias[t]}"
                )

            elements.append(Paragraph(texto, self.styles["ListText"]))
            elements.append(Spacer(1, 0.05 * inch))
            

        #Table 1
        # Typologies results table
        elements.append(Paragraph("Resultados de tipologías municipales y distritales", self.styles["TableTitle"]))
        elements.append(Paragraph("Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        
        total_municipalities = len(data)

        typology_table_data = [
            ["Tipología",
            "Núm. municipios y\n distritos",
            "% municipios y\n distritos",
            "Índice de capacidades\nponderado (Promedio)"]
        ]

        # Specify the order of typologies to display
        tipologias = [
            ("Bogotá", "Bogotá"),
            ("Ciudades grandes", "Ciudades grandes"),
            ("1", 1),
            ("2", 2),
            ("3", 3),
            ("4", 4),
            ("5", 5),
        ]

        promedios_tipologias = []
        for label, t in tipologias:

            subset = data[data["Tipología_2026_CortesArcMap"] == t]

            cantidad = len(subset)
            
            #Color Adjustment
            if cantidad == 0:
                promedio_display = "-"
                promedio = None
            else:
                promedio = subset["ICPond_2026"].mean()
                promedio_display = f"{promedio:.2f}".replace(".", ",")

            promedios_tipologias.append(promedio)

            if total_municipalities > 0:
                porcentaje = cantidad / total_municipalities
            else:
                porcentaje = 0

            porcentaje_display = f"{porcentaje*100:.0f}%"

            typology_table_data.append([
                str(label),
                str(cantidad),
                porcentaje_display,
                promedio_display
            ])

        total_mean = data["ICPond_2026"].mean()

        typology_table_data.append([
            "Total",
            str(total_municipalities),
            "100%",
            f"{total_mean:.2f}".replace(".", ",")
        ])

        typology_table = Table(
            typology_table_data,
            colWidths=[4*cm, 4*cm, 4*cm, 5*cm],
            repeatRows=1
        )
        
        # obtener valores válidos
        valid_values = [v for v in promedios_tipologias if v is not None]

        # ordenar de mayor a menor
        sorted_vals = sorted(valid_values, reverse=True)

        row_colors = {}

        for i, val in enumerate(promedios_tipologias):
            if val is None:
                continue

            rank = sorted_vals.index(val)

            if rank < len(self.ranking_colors):
                row_colors[i+1] =   self.ranking_colors[rank]
        

        table_style = TableStyle([
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

            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ])
        table_style.add('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
        table_style.add('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
        
        for row, color in row_colors.items():
            table_style.add('BACKGROUND', (3, row), (3, row), color)

        typology_table.setStyle(table_style)
        
        elements.append(typology_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(
            Paragraph(
                "Fuente: Elaboración propia con base en información del DNP",
                self.styles["SourceText"]
            )
        )
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
        
        #Table 2
        text_cols = data.select_dtypes(include='object').columns
        data[text_cols] = data[text_cols].replace(["Sin Información", "-"], pd.NA)
        pd.set_option('future.no_silent_downcasting', True)

        #Convert relevant columns to numeric, coercing errors to NaN
        data["IPM_2018"] = pd.to_numeric(data["IPM_2018"], errors="coerce")
        data["NBI_2018"] = pd.to_numeric(data["NBI_2018"], errors="coerce")
        data["IRCA_2024"] = pd.to_numeric(data["IRCA_2024"], errors="coerce")
        data["IICA_2023"] = pd.to_numeric(data["IICA_2023"], errors="coerce")
        
        ipm_vals = []
        nbi_vals = []
        irca_vals = []
        iica_vals = []
        
        def aplicar_colores(valores, col_index, table_style):

            valid = [v for v in valores if v is not None]

            if not valid:
                return

            # menor = mejor
            sorted_vals = sorted(valid)

            for i, val in enumerate(valores):

                if val is None:
                    continue

                rank = sorted_vals.index(val)

                if rank < len(self.ranking_colors):
                    table_style.add(
                        'BACKGROUND',
                        (col_index, i+1),
                        (col_index, i+1),
                        self.ranking_colors[rank]
                    )
        
        

        def promedio_tipologia(columna, tipologia):

            subset = data[data["Tipología_2026_CortesArcMap"] == tipologia]

            if subset.empty:
                return "-"

            valor = subset[columna].mean()

            if pd.isna(valor):
                return "-"

            return f"{valor:.2f}".replace(".", ",")

        
        # Complementary variables table
        elements.append(Paragraph("Tipologías vs. variables complementarias", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))

        tipologias = ["Bogotá", "Ciudades grandes", 1, 2, 3, 4, 5]

        def promedio_tipologia(columna, tipologia):
            subset = data[data["Tipología_2026_CortesArcMap"] == tipologia]
            if subset.empty:
                return "-"
            return f"{subset[columna].mean():.2f}".replace(".", ",")

        complementary_table_data = [
            ["Tipología", "IPM", "NBI", "IRCA", "IICA"]
        ]

        for t in tipologias:
            subset = data[data["Tipología_2026_CortesArcMap"] == t]

            if subset.empty:
                ipm = "-"
                nbi = "-"
                irca = "-"
                iica = "-"

                ipm_vals.append(None)
                nbi_vals.append(None)
                irca_vals.append(None)
                iica_vals.append(None)

            else:
                ipm_val = subset["IPM_2018"].mean()
                nbi_val = subset["NBI_2018"].mean()
                irca_val = subset["IRCA_2024"].mean()
                iica_val = subset["IICA_2023"].mean()

                ipm_vals.append(ipm_val)
                nbi_vals.append(nbi_val)
                irca_vals.append(irca_val)
                iica_vals.append(iica_val)

                ipm = f"{ipm_val:.2f}".replace(".", ",")
                nbi = f"{nbi_val:.2f}".replace(".", ",")
                irca = f"{irca_val:.2f}".replace(".", ",")
                iica = f"{iica_val:.2f}".replace(".", ",")

            complementary_table_data.append([
                str(t),
                ipm,
                nbi,
                irca,
                iica
            ])

        # Total row
        total_ipm = f"{data['IPM_2018'].mean():.2f}".replace(".", ",")
        total_nbi = f"{data['NBI_2018'].mean():.2f}".replace(".", ",")
        total_irca = f"{data['IRCA_2024'].mean():.2f}".replace(".", ",")
        total_iica = f"{data['IICA_2023'].mean():.2f}".replace(".", ",")

        complementary_table_data.append([
            "Total",
            total_ipm,
            total_nbi,
            total_irca,
            total_iica
        ])
        
        
        complementary_table = Table(complementary_table_data, colWidths=[3.5*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.8*cm])
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ])
        
        aplicar_colores(ipm_vals, 1, table_style)
        aplicar_colores(nbi_vals, 2, table_style)
        aplicar_colores(irca_vals, 3, table_style)
        aplicar_colores(iica_vals, 4, table_style)
        complementary_table.setStyle(table_style)
        
        table_style.add(
            'BACKGROUND',
            (0, len(complementary_table_data)-1),
            (-1, len(complementary_table_data)-1),
            colors.lightgrey
        )
        complementary_table.setStyle(table_style)
        

            

        elements.append(complementary_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en la información del DANE y DNP", self.styles["SourceText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        #Table 2.1 - MDM table
        elements.append(Paragraph("Tipologías vs. MDM 2024", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        mdm_data = [
            ["Tipología", "MDM", "MDM Resultados", "Educación", "Salud", "Servicios", "Seguridad y\n convivencia"]
        ]

        tipologias = [
            ("Bogotá", "Bogotá"),
            ("Ciudades grandes", "Ciudades grandes"),
            ("1", 1),
            ("2", 2),
            ("3", 3),
            ("4", 4),
            ("5", 5),
        ]

        columnas = [
            "MDM_2024",
            "MDM2024_resultados",
            "MDM2024_educacion",
            "MDM2024_salud",
            "MDM2024_servicios",
            "MDM2024_seguridad"
        ]

        for label, t in tipologias:

            subset = data[data["Tipología_2026_CortesArcMap"] == t]

            fila = [label]

            if len(subset) == 0:
                fila += ["-"] * len(columnas)

            else:

                for col in columnas:

                    promedio = subset[col].mean()

                    if pd.isna(promedio):
                        fila.append("-")
                    else:
                        valor = promedio
                        texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        fila.append(texto)

            mdm_data.append(fila)
            
        fila_total = ["Total"]

        for col in columnas:

            promedio = data[col].mean()

            if pd.isna(promedio):
                fila_total.append("-")
            else:
                valor = promedio
                texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                fila_total.append(texto)

        mdm_data.append(fila_total)
        
        mdm_table = Table(mdm_data, colWidths=[3*cm, 2.2*cm, 2.6*cm, 2.2*cm, 2.2*cm, 2.2*cm, 3*cm])

        mdm_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),

            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),

            ('GRID', (0, 0), (-1, -1), 1, colors.black),

            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))
        
        elements.append(mdm_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(
            Paragraph(
                "Fuente: Elaboración propia con base en información del DNP",
                self.styles["SourceText"]
            )
        )
        elements.append(Spacer(1, 0.1 * inch))
        
        
        # ==================== PAGE 4 ====================
        
        # Income section
        elements.append(Paragraph("Ingresos", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        income_text = "A medida que se avanza en las tipologías los ingresos propios en promedio tienden a reducirse."
        elements.append(Paragraph(income_text, self.styles["NormalText"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        elements.append(Paragraph("Ingresos propios 2024. Cifras en pesos", self.styles["TableTitle"]))
        elements.append(Paragraph("Nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        #Table 3
        income_table_data = [
            ["Tipología", "Ingresos tributarios pér capita"]
        ]

        tipologias = [
            ("Bogotá", "Bogotá"),
            ("Ciudades grandes", "Ciudades grandes"),
            ("1", 1),
            ("2", 2),
            ("3", 3),
            ("4", 4),
            ("5", 5),
        ]

        for label, t in tipologias:

            subset = data[data["Tipología_2026_CortesArcMap"] == t]

            cantidad = len(subset)

            if cantidad == 0:
                ingreso_display = "-"
            else:
                promedio = subset["Ingresos_propios_pc_2024"].mean() * 1000000
                ingreso_display = f"{promedio:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

            income_table_data.append([
                str(label),
                ingreso_display
            ])

        # Total row
        if len(data) == 0:
            total_display = "-"
        else:
            total_mean = data["Ingresos_propios_pc_2024"].mean() * 1000000
            total_display = f"{total_mean:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

        income_table_data.append([
            "Total",
            total_display
        ])
        
        income_table = Table(income_table_data, colWidths=[5*cm, 6*cm])

        income_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), TABLE_HEADER_BLUE),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),

            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),

            ('GRID', (0, 0), (-1, -1), 1, colors.black),

            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.black),
        ]))

        elements.append(income_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(
            Paragraph(
                "Fuente: Elaboración propia con base en las operaciones efectivas de caja calculadas por el DNP",
                self.styles["SourceText"]
            )
        )
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
        elements.append(Paragraph("Nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        
        #Table 4
        protected_areas_data = [
            ["Tipología", "% Área de protección ambiental y ecosistemas estratégicos", "", "", "", "", "Total"],
            ["", "0", "0,1% - 30%", "31% - 50%", "51% - 70%", "71% - 100%", ""],
        ]

        tipologias = [
            ("Bogotá", "Bogotá"),
            ("Ciudades grandes", "Ciudades grandes"),
            ("1", 1),
            ("2", 2),
            ("3", 3),
            ("4", 4),
            ("5", 5),
        ]

        rangos = [
            "0",
            "0,1% - 30%",
            "31% - 50%",
            "51% - 70%",
            "71% - 100%"
        ]

        totales_columnas = [0]*len(rangos)
        total_general = 0

        for label, t in tipologias:

            subset_tipologia = data[data["Tipología_2026_CortesArcMap"] == t]

            fila = [label]
            total_fila = 0

            for i, r in enumerate(rangos):

                col = subset_tipologia["Rangos_AA_V2026"].astype(str).str.strip()

                if r == "0":
                    count = (col == "0").sum()
                else:
                    count = (col == r).sum()

                fila.append(str(count))

                totales_columnas[i] += count
                total_fila += count

            fila.append(str(total_fila))
            total_general += total_fila

            protected_areas_data.append(fila)

        fila_total = ["Total"]

        for val in totales_columnas:
            fila_total.append(str(val))

        fila_total.append(str(total_general))

        protected_areas_data.append(fila_total)
        
        protected_areas_table = Table(protected_areas_data, colWidths=[2.5*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm])
        protected_areas_table.setStyle(TableStyle([
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
        
        elements.append(protected_areas_table)
        elements.append(Spacer(1, 0.05 * inch))
        elements.append(Paragraph("Fuente: Elaboración propia con base en información del DNP", self.styles["SourceText"]))
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
        elements.append(Paragraph("Nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["TableTitle"]))
        elements.append(Spacer(1, 0.05 * inch))
        
        #Table 5
        ethnic_territories_data = [
            ["Tipología", "% Área de Territorios Étnicos", "", "", "", "", "Total"],
            ["", "0%", "0,1% - 30%", "31% - 50%", "51% - 70%", "71% - 100%", ""],
        ]

        tipologias = [
            ("Bogotá", "Bogotá"),
            ("Ciudades grandes", "Ciudades grandes"),
            ("1", 1),
            ("2", 2),
            ("3", 3),
            ("4", 4),
            ("5", 5),
        ]

        rangos = [
            "0",
            "0,1% - 30%",
            "31% - 50%",
            "51% - 70%",
            "71% - 100%"
        ]

        totales_columnas = [0]*len(rangos)
        total_general = 0

        for label, t in tipologias:

            subset_tipologia = data[data["Tipología_2026_CortesArcMap"] == t]

            fila = [label]
            total_fila = 0

            col = subset_tipologia["Rangos_ATE_V2026"].astype(str).str.strip()

            for i, r in enumerate(rangos):

                if r == "0":
                    count = (col == "0").sum()
                else:
                    count = (col == r).sum()

                fila.append(str(count))

                totales_columnas[i] += count
                total_fila += count

            fila.append(str(total_fila))
            total_general += total_fila

            ethnic_territories_data.append(fila)


        fila_total = ["Total"]

        for val in totales_columnas:
            fila_total.append(str(val))

        fila_total.append(str(total_general))

        ethnic_territories_data.append(fila_total)
        
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
        elements.append(Paragraph("Fuente: Elaboración propia con base en información del DNP", self.styles["SourceText"]))
        
        
        
        elements.append(Spacer(1, 0.1 * inch))
       
        
        elements.append(PageBreak())
        # ==================== PAGE 7 ====================
        
        # Annex
        elements.append(Paragraph("Anexo", self.styles["SectionTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("Tipologías a nivel municipal- Departamento de {departmentName}".format(departmentName=departmentName), self.styles["SubTitle"]))
        elements.append(Spacer(1, 0.1 * inch))
        
        # Municipality data for annex
        annex_data = [
            ["Código DANE", "Departamento", "Municipio", "Tipología 2026"]
        ]
        
        # Personalized sorting to ensure Bogotá and "Ciudades grandes" are at the top
        orden_tipologia = ["Ciudades grandes", 1, 2, 3, 4, 5]

        data["Tipologia_orden"] = data["Tipologia_2026R"].astype(str)

        data["Tipologia_orden"] = pd.Categorical(
            data["Tipologia_orden"],
            categories=[str(x) for x in orden_tipologia],
            ordered=True
        )

        data_ordenado = data.sort_values(["Tipologia_orden", "Municipio"])

        for _, row in data_ordenado.iterrows():
            annex_data.append([
                str(row["CodDANE_txt"]),
                row["Departamento"],
                row["Municipio"],
                str(row["Tipologia_2026R"])
            ])
        
        annex_table = Table(
            annex_data,
            colWidths=[3*cm, 3.5*cm, 4.5*cm, 3*cm],
            repeatRows=1
        )
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
        
        # Build PDF with custom header and footer on each page
        doc.build(elements, onFirstPage=self._header_footer, onLaterPages=self._header_footer)
        
    
    def _header_footer(self, canvas, doc):
        """Add header and footer to each page"""
        canvas.saveState()
        
        # Calculate centered position for the logo
        page_width = 612  
        logo_width = 120
        logo_x = (page_width - logo_width) / 2
        
        try:
            if os.path.exists(self.logo_path):
                canvas.drawImage(self.logo_path, logo_x, 750, width=logo_width, height=35, preserveAspectRatio=True, mask='auto')
            else:
                print(f"Logo no encontrado en: {self.logo_path}")
        except Exception as e:
            print(f"Error al cargar el logo: {e}")
        
        # Header line
        line_width = 530
        line_start = (page_width - line_width) / 2
        canvas.setStrokeColor(DNP_BLACK)
        canvas.setLineWidth(2)
          
        # Footer line
        canvas.setStrokeColor(DNP_BLACK)
        canvas.setLineWidth(1)
        canvas.line(line_start, 80, line_start + line_width, 80)  
        
        # Footer text 
        canvas.setFont("Helvetica", 7)
        # Footer text
        canvas.setFillColor(DNP_BLACK)

        # Dirección
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(line_start, 65, "Dirección:")
        canvas.setFont("Helvetica", 7)
        canvas.drawString(line_start + 35, 65, " Calle 26 # 13 – 19 Bogotá, D.C., Colombia")

        # Conmutador
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(line_start, 55, "Conmutador:")
        canvas.setFont("Helvetica", 7)
        canvas.drawString(line_start + 45, 55, " 601 3815000")

        # Línea gratuita
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(line_start, 45, "Línea gratuita:")
        canvas.setFont("Helvetica", 7)
        canvas.drawString(line_start + 50, 45, " PBX 381 5000")
        
        # Page number 
        page_num = getattr(doc, 'page', 1)
        canvas.drawRightString(line_start + line_width, 45, f"Página {page_num}")
        
        canvas.restoreState()

