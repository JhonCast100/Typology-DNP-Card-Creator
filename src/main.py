from Controllers.DataPdfController import DataPdfController
from Models import PdfGenerator
from Views.MainView import run
from Models.DataTable import DataTable

def main():
    #Import data from Excel file
    data = DataTable("../Data/ExcelFiles/MatrizTipologias.xlsx")  
    df = data.getDataFrame() 
    
    #Create an instance of the DataPdfController with the imported data
    controller = DataPdfController(df)
    controller.dataToPdf("CUNDINAMARCA")
    controller.dataToPdf("ANTIOQUIA") 
    controller.dataToPdf("NARIÑO") 
    controller.dataToPdf("BOLÍVAR") 
     
    #Call to the main view of the program
    #run() 

if __name__ == "__main__":
    main()

