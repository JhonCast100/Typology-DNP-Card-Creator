# Generación de Ejecutable - DNPCardCreator

## ¿Cómo generar el .exe?

### Opción 1: Automático (Recomendado)

1. Abre PowerShell o CMD en la carpeta raíz del proyecto
2. Ejecuta:
   ```
   .\build.bat
   ```
3. ¡Listo! El ejecutable se generará automáticamente en `dist\DNPCardCreator\DNPCardCreator.exe`

### Opción 2: Manual

1. Instala PyInstaller (si no lo tienes):
   ```
   pip install pyinstaller
   ```

2. Desde la carpeta raíz del proyecto, ejecuta:
   ```
   pyinstaller build.spec
   ```

3. Configura las carpetas de distribución:
   ```
   python setup_dist.py
   ```

## Estructura del ejecutable

El .exe se generará con la siguiente estructura:
```
dist/
└── DNPCardCreator/
    ├── DNPCardCreator.exe          ← El ejecutable
    ├── Data/                        ← Carpeta de datos (Excel)
    │   └── ExcelFiles/
    │       └── MatrizTipologias.xlsx
    ├── Output/                      ← Carpeta de salida (PDFs)
    └── [otras librerías necesarias]
```

## Instrucciones de uso del ejecutable

1. Coloca el archivo **MatrizTipologias.xlsx** en la carpeta `Data\ExcelFiles\`
2. Coloca tus mapas en la carpeta `Data\Maps\`
3. Ejecuta **DNPCardCreator.exe**
4. Los PDFs generados aparecerán en la carpeta `Output\`

## Requisitos

- Python 3.7 o superior
- Las siguientes librerías (instaladas automáticamente por build.bat):
  - pyinstaller
  - openpyxl
  - pillow
  - reportlab
  - pandas

## Distribución del ejecutable

Para compartir el ejecutable con otros usuarios:
1. Comparte la carpeta completa `dist\DNPCardCreator\`
2. Asegúrate de que incluya:
   - DNPCardCreator.exe
   - Las carpetas Data\ y Output\
   - Todas las librerías necesarias (generadas automáticamente)

**No necesita cambiar rutas ni código. Funciona tal como está.**
