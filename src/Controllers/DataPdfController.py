
from Models.PdfGenerator import PdfGenerator


class  DataPdfController:
    def __init__(self , df):
        self.df = df
       
    #Method to generate the PDF report with the data of the selected department 
    def dataToPdf(self, departmentName):
        # Filter by department
        departmentData = self.df[self.df["Departamento"] == departmentName]
        
        # Create generator instance and produce PDF
        generator = PdfGenerator(f"reporte_{departmentName}_final.pdf")
        generator.generatePdf(departmentName, departmentData)
        
        