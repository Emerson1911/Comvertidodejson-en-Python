import pandas as pd

# Cargar el archivo Excel
file_path = "C:\\Users\\emers\\Desktop\\Galones de gasolina.xlsx"
excel_data = pd.ExcelFile(file_path)

# Seleccionar la hoja que quieres convertir (por ejemplo, la primera)
sheet_name = excel_data.sheet_names[0]  # Puedes cambiar esto al nombre de la hoja si la conoces
df = excel_data.parse(sheet_name)

# Limpiar el DataFrame si es necesario
# df = df.dropna(how='all')  # Elimina filas completamente vacías

# Convertir el DataFrame a un diccionario y luego a JSON, con indentación para mayor legibilidad
json_data = df.to_json(orient='records', indent=4)

# Guardar el JSON en un archivo si lo deseas
with open('La vaz a matar perro.json', 'w') as json_file:
    json_file.write(json_data)

# Mostrar el JSON generado de manera legible
print(json_data)
