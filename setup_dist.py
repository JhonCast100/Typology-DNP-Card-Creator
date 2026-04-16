import os
import shutil
import sys

def setup_distribution():
    """Configura la carpeta dist con las carpetas necesarias para el ejecutable"""
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, 'dist', 'DNPCardCreator')
    
    # Crear carpetas de Data y Output en dist si no existen
    data_dir = os.path.join(dist_dir, 'Data')
    output_dir = os.path.join(dist_dir, 'Output')
    
    # Crear carpeta Data desde el raíz si es necesario
    original_data = os.path.join(base_dir, 'Data')
    if not os.path.exists(data_dir) and os.path.exists(original_data):
        print(f"Copiando Data a {data_dir}...")
        shutil.copytree(original_data, data_dir)
    
    # Crear carpeta Output
    if not os.path.exists(output_dir):
        print(f"Creando {output_dir}...")
        os.makedirs(output_dir, exist_ok=True)
    
    print(f"✓ Distribución configurada correctamente en: {dist_dir}")

if __name__ == "__main__":
    setup_distribution()
