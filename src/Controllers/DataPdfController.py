
class  DataPdfController:
    def __init__(self , df):
        self.df = df
       
    #Method to generate the PDF report with the data of the selected department 
    def dataToPdf(self, departmentName):
        # Filter by department
        departmentData = self.df[self.df["Departamento"] == departmentName]
        print(departmentData["Municipio"].tolist())
        