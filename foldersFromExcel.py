import os
import openpyxl

# Ruta donde se crearán las carpetas
foldersCreated = r"C:\Users\Sistemas\Desktop\Permisos"

# Ruta del archivo Excel
foldersNames = r"C:\Users\Sistemas\Downloads\Usuario (res.users).xlsx"

# Cargar el libro y seleccionar la hoja
workbook = openpyxl.load_workbook(foldersNames)
sheet = workbook['Sheet1']

# Leer los valores de la columna A desde la fila 2
columnValues = [cell.value for cell in sheet['A'][1:] if cell.value]

# Crear carpetas
for folderName in columnValues:
    folderPath = os.path.join(foldersCreated, str(folderName))
    if not os.path.exists(folderPath):
        os.makedirs(folderPath)
        print(f"Carpeta creada: {folderPath}")
    else:
        print(f"La carpeta ya existe: {folderPath}")

