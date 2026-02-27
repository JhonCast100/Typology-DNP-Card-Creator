import pandas as pd


class DataTable:
    
    #Builder Method
    def __init__(self,route):
        self.route = route
        self.df = None
        
        #Se deben agregar las columnas requeridas para el programa
    def getDataFrame(self):
        requiredColumns = [
            "CodDANE_txt",
            "Departamento",
            "Municipio"
        ]
        
        df = pd.read_excel(self.route, sheet_name="Municipios")
        print(df.columns.tolist())
        
        # Verify that all required columns are present
        missing = [c for c in requiredColumns if c not in df.columns]
        if missing:
            raise ValueError(f"Faltan columnas en el Excel: {missing}")

        # Filter the DataFrame to keep only the necessary columns
        self.df = df[requiredColumns]

        return self.df
        