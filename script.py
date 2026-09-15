import pandas as pd

# 1. Cargar el archivo Excel
archivo_excel = r"C:\Users\USUARIO\Desktop\Listado-actualizado-Renta-04062026 -ACTUA.xlsx"
# Si deseas leer una hoja específica, usa: sheet_name='NombreDeHoja'
df = pd.read_excel(archivo_excel)

# 2. Convertir y guardar a un archivo JSON
archivo_json = 'resultado.json'
df.to_json(archivo_json, orient='records', indent=4, force_ascii=False)

print(f"Conversión completada. Archivo guardado como {archivo_json}")