from pathlib import Path
import calendar

# Creamos una carpeta para almacenar los archivos de calendario
Path('nueva_carpeta').mkdir(exist_ok=True)

# Creamos una carpeta dentro de la carpeta principal
Path('Seccion_3/carpeta1/carpeta2/carpeta3').mkdir(parents=True, exist_ok=True)

# Creamos una estructura de carpetas para cada mes del año y cada día de la semana
meses_ano = list(calendar.month_name)[1:]  # Obtenemos los nombres de los meses del año (excluyendo el primer elemento vacío)
dias_semana = ['Dia 1', 'Dia 10', 'Dia 20', 'Dia 30']  # Lista de días de la semana

for i, mes in enumerate(meses_ano):
    for dia in dias_semana:
        Path(f'Seccion_3/2026/{i+1}.{mes}/{dia}').mkdir(parents=True, exist_ok=True)  # Creamos la estructura de carpetas para cada mes y día
        print('')