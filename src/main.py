from Controllers.DataPdfController import DataPdfController
from Models import PdfGenerator
from Views.MainView import run
from Models.DataTable import DataTable

def main():
    #Import data from Excel file
    data = DataTable("../Data/ExcelFiles/MatrizTipologias.xlsx")  
    df = data.getDataFrame() 
    #print(df.tail()) 
    
    #Create an instance of the DataPdfController with the imported data
    controller = DataPdfController(df)
    controller.dataToPdf("Casanare")  
     
    #Call to the main view of the program
    run() 

if __name__ == "__main__":
    main()

