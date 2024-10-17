import pandas as pd
import json

# Función para cargar cualquier archivo Excel y ajustar el contenido correctamente
def excel_to_json(file_path, output_file='salida.txt'):
    try:
        # Cargar el archivo Excel
        excel_data = pd.ExcelFile(file_path)
        
        # Iterar sobre todas las hojas para detectar el contenido
        for sheet_name in excel_data.sheet_names:
            # Leer cada hoja en un DataFrame
            df = excel_data.parse(sheet_name)

            # Limpiar datos (opcional, dependiendo de lo que esperes en otros archivos)
            df = df.dropna(how='all')  # Elimina filas completamente vacías
            df = df.fillna('')  # Rellena NaN con cadenas vacías si es necesario

            # Convertir el DataFrame a una lista de diccionarios (JSON-like structure)
            json_data = df.to_dict(orient='records')

            # Guardar el JSON en un archivo .txt con el formato deseado
            with open(output_file, 'w') as json_file:
                json.dump(json_data, json_file, indent=4)

            # Imprimir el JSON generado para revisión (opcional)
            print(json.dumps(json_data, indent=4))
        
        print(f"El archivo ha sido procesado y guardado en: {output_file}")

    except Exception as e:
        print(f"Ocurrió un error al procesar el archivo Excel: {e}")

# Llamada a la función con el archivo Excel
file_path = r"C:\Users\emers\Desktop\Galones de gasolina.xlsx"
excel_to_json(file_path)
