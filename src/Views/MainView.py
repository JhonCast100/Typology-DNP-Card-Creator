import tkinter as tk
from PIL import Image, ImageTk

from Controllers.DataPdfController import DataPdfController
from Models import PdfGenerator
from Models.DataTable import DataTable

def run():
    # Create main window
    window = tk.Tk()
    window.title("Creador de Fichas Tipología DNP")
    window.geometry("900x600")
    window.configure(bg="#f4f6f8")  # color de fondo suave

    # Main container
    container = tk.Frame(window, bg="#f4f6f8")
    container.pack(expand=True)

    # Title
    title = tk.Label(
        container,
        text="Bienvenido al Creador de Fichas de Tipologías del DNP",
        font=("Arial", 22, "bold"),
        bg="#f4f6f8",
        fg="#1f2933",
        wraplength=700,
        justify="center"
    )
    title.pack(pady=(30, 20))

    # Instructions
    instructions_text = (
        "Para utilizar el programa:\n\n"
        "1. Coloca el Excel de Tipologías en la carpeta \\Data\\ExcelFiles.\n"
        "   El archivo debe llamarse: \"MatrizTipologias.xlsx\"\n\n"
        "2. El Excel debe contener la hoja llamada \"Municipios\",\n"
        "   dentro de la cual deben estar los datos a cargar.\n\n"
        "3. Coloca las imágenes de los mapas en la carpeta \\Data\\Maps.\n"
        "   Cada imagen debe tener el nombre del departamento."
    )

    instructions = tk.Label(
        container,
        text=instructions_text,
        font=("Arial", 12),
        bg="#f4f6f8",
        fg="#374151",
        justify="left",
        wraplength=700
    )
    instructions.pack(pady=20)

    # Button to generate cards
    def generate():
        #Import data from Excel file
        data = DataTable("../Data/ExcelFiles/MatrizTipologias.xlsx")  
        df = data.getDataFrame() 
        
        #Create an instance of the DataPdfController with the imported data
        controller = DataPdfController(df)
        """
        for department in df["Departamento"].dropna().unique():
            controller.dataToPdf(department)
        """
        controller.dataToPdf("CUNDINAMARCA")
        controller.dataToPdf("ANTIOQUIA") 
        controller.dataToPdf("NARIÑO") 
        controller.dataToPdf("BOLÍVAR") 

    generate_button = tk.Button(
        container,
        text="Generar",
        font=("Arial", 14, "bold"),
        bg="#2563eb",
        fg="white",
        activebackground="#1d4ed8",
        activeforeground="white",
        width=20,
        height=2,
        command=generate
    )
    generate_button.pack(pady=30)


    # Watermark (footer)
    watermark = tk.Label(
        window,
        text="Diseñado por Jhon Javier Castañeda Alvarado",
        font=("Arial", 9),
        bg="#f4f6f8",
        fg="#6b7280"
    )
    watermark.pack(side="bottom", pady=10)

    # Execute the main loop
    window.mainloop()
    
if __name__ == "__main__":
    run()