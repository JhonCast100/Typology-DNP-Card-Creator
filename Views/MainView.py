import tkinter as tk
from PIL import Image, ImageTk

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
    "1. Coloca el Excel de Tipologías en la carpeta ExcelFiles.\n"
    "   El archivo debe llamarse: \"MatrizTipologias.xlsx\"\n\n"
    "2. El Excel debe contener la hoja llamada \"Municipios\",\n"
    "   dentro de la cual deben estar los datos a cargar.\n\n"
    "3. Coloca las imágenes de los mapas en la carpeta Maps.\n"
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
def generar():
    print("Generar fichas...")  # luego aquí pondrás la lógica real

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
    command=generar
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

# Ejecutar app
window.mainloop()