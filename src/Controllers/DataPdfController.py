
from Models.PdfGenerator import PdfGenerator


class  DataPdfController:
    def __init__(self , df):
        self.df = df
       
    #Method to generate the PDF report with the data of the selected department 
    def dataToPdf(self, departmentName):

        departmentData = self.df[
            self.df["Departamento"] == departmentName
        ].copy()  # importante el copy()

        # 🔥 AQUÍ VA LA TRANSFORMACIÓN
        departmentData.loc[:, "Municipio"] = departmentData["Municipio"].str.title()
        departmentData.loc[:, "Departamento"] = departmentData["Departamento"].str.title()

        generator = PdfGenerator(f"reporte_{departmentName}_final.pdf")
        generator.generatePdf(departmentName, departmentData)