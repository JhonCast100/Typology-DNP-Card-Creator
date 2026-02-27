from Views.MainView import run
from Models.DataTable import DataTable

def main():
    #Import data from Excel file
    data = DataTable("../Data/ExcelFiles/MatrizTipologias.xlsx")  
    df = data.getDataFrame()  
    print(df.head()) 
    
    #Call to the main view of the program
    run() 

if __name__ == "__main__":
    main()

