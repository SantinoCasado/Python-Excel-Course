import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, Color, PatternFill, Border, Side

df = pd.read_excel('Seccion_4/data/supermarket_sales.xlsx')  # Leemos el archivo Excel y lo almacenamos en un DataFrame

# ------------ Manipulacion y Guardado a traves de pivots ----------------------
# Creamos una tabla pivote con los datos del DataFrame, agrupando por 'Gender' y 'Product line', y sumando los valores de la columna 'Total'
table_pivote = df.pivot_table(index='Gender', columns='Product line', values='Total', aggfunc='sum').round(0)  # Un pivote es una tabla que permite reorganizar y resumir datos de un DataFrame, facilitando el análisis de la información.

# Guardamos la tabla pivote en un archivo Excel
table_pivote.to_excel('Seccion_4/data/supermarket_sales_pivot.xlsx', startrow=4, sheet_name='Report')  # Guardamos la tabla pivote en un archivo Excel, comenzando desde la fila 4 y en la pestaña 'Report'

# ------------ Grafico de barras y otros ----------------------
wb = load_workbook('Seccion_4/data/supermarket_sales_pivot.xlsx')  # Cargamos el archivo Excel con la tabla pivote
pestania = wb['Report'] # Seleccionamos la pestaña 'Report' del archivo Excel

# Seleccionamos el rango de celdas que contiene los datos de la tabla pivote para crear un gráfico
min_col = wb.active.min_column  # Obtenemos el número de la primera columna con datos en la pestaña activa
max_col = wb.active.max_column  # Obtenemos el número de la última columna con datos en la pestaña activa
min_row = wb.active.min_row  # Obtenemos el número de la primera fila con datos en la pestaña activa
max_row = wb.active.max_row  # Obtenemos el número de la última fila con datos en la pestaña activa
print(f"Rango de celdas: de {min_col} a {max_col} columnas y de {min_row} a {max_row} filas")  # Imprimimos el rango de celdas que contiene los datos de la tabla pivote

barChart = BarChart()  # Creamos un objeto de tipo BarChart para generar un gráfico de barras

data = Reference(pestania, min_col=min_col+1, max_col=max_col, min_row=min_row, max_row=max_row)  # Creamos una referencia a los datos de la tabla pivote para el gráfico
categories = Reference(pestania, min_col=min_col, max_col=min_col, min_row=min_row+1, max_row=max_row)  # Creamos una referencia a las categorías de la tabla pivote para el gráfico

barChart.add_data(data, titles_from_data=True)  # Agregamos los datos al gráfico
barChart.add_data(categories)

pestania.add_chart(barChart, "B12")  # Agregamos el gráfico a la pestaña 'Report' del archivo Excel, comenzando desde la celda B10
barChart.title = "Ventas"  # Establecemos el título del gráfico
barChart.style = 5  # Establecemos el estilo del gráfico

wb.save('Seccion_4/data/supermarket_sales_pivot.xlsx')  # Guardamos los cambios en el archivo Excel