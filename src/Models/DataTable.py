import pandas as pd
from pathlib import Path


class DataTable:
    
    #Builder Method
    def __init__(self,route):
        # Path of the Excel file containing the data
        base_path = Path(__file__).resolve().parent.parent.parent
        self.route = base_path / "Data" / "ExcelFiles" / "MatrizTipologias.xlsx"
        self.df = None
        
        #Se deben agregar las columnas requeridas para el programa
    def getDataFrame(self):
        requiredColumns = [
            "CodDANE_txt",
            "Departamento",
            "Municipio",
            "Tipología_2026_CortesArcMap",
            "CodDANE_dpto",
            "Tipologia_2026R",
            "IC_2026",
            "ICPond_2026",
            "IPM_2018",
            "NBI_2018",
            "IRCA_2024",
            "IICA_2023",
            "Ingresos_propios_pc_2024",
            "Rangos_AA_V2026",
            "Rangos_ATE_V2026",
            "MDM_2024",
            "MDM2024_resultados",
            "MDM2024_educacion",
            "MDM2024_salud",
            "MDM2024_servicios",
            "MDM2024_seguridad"
            
        ]
        
        df = pd.read_excel(self.route, 
                           sheet_name="Municipios",
                           header=1,
                           nrows=1104,
                           dtype={"CodDANE_txt": str, 
                                  "CodDANE_dpto": str
                                  }
                           )

        
        # Verify that all required columns are present
        missing = [c for c in requiredColumns if c not in df.columns]
        if missing:
            raise ValueError(f"Faltan columnas en el Excel: {missing}")

        # Filter the DataFrame to keep only the necessary columns
        self.df = df[requiredColumns]
        

        return self.df
        