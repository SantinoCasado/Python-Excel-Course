import pandas as pd

simp = pd.read_html('https://es.wikipedia.org/wiki/Anexo:Pa%C3%ADses_por_poblaci%C3%B3n', header=0)  # Leemos todas las tablas HTML de la página web y almacenamos en una lista de DataFrames
df = simp[0]  # Seleccionamos la primera tabla de la lista de Data

# Guardamos la primera tabla en un archivo CSV sin el índice
simp[0].to_csv('Seccion_3/poblacion.csv', index=False)  # Guardamos la primera tabla en un archivo CSV sin el índice

# Guardamos la primera tabla en un archivo Excel sin el índice
simp[0].to_excel('Seccion_3/poblacion.xlsx', index=False)
