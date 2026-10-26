from pathlib import Path

# Creamos una carpeta llamada 'test'
Path('Seccion_3/test').mkdir(exist_ok=True)

# Renombramos la carpeta 'test' a 'test_renamed'
path = Path('Seccion_3/test')
path.rename('Seccion_3/test_renamed')  # Renombramos la carpeta 'test' a 'test_renamed'

# Renombramos un archivo dentro de la carpeta 'test_renamed'
path = Path('Seccion_3/test_renamed/gastos.txt')
nuevoNombrePath = path.with_name('gastos_renombrado.txt')  # Creamos un nuevo Path con el nuevo nombre
path.rename(nuevoNombrePath)  # Renombramos el archivo 'gastos.txt'


carpeta = Path('Seccion_3/2026')  # Creamos un objeto Path para la carpeta '2026'
for path in list(carpeta.iterdir()):  # Iteramos sobre los elementos dentro de la carpeta '2026'
    print(path.name)  # Imprimimos el nombre de cada elemento dentro de la carpeta '2026'

# Obtener path de todos los subdirectorios
carpeta = Path('Seccion_3/2026')  # Creamos un objeto Path para la carpeta '2026'
paths = carpeta.glob('**/*')  # Obtenemos todos los subdirectorios dentro de la carpeta '2026'

for path in paths:  # Iteramos sobre los subdirectorios obtenidos
    if path.is_file():  # Verificamos si el path es un archivo
        print(path)  # Imprimimos el path de cada archivo
    elif path.is_dir():  # Verificamos si el path es un directorio
        print(path)  # Imprimimos el path de cada subdirectorio
    else:  # Si el path no es ni un archivo ni un directorio
        print('No hay archivos ni subdirectorios')  # Imprimimos un mensaje indicando que no hay archivos ni subdirectorios


# Renombrar archivos con extensión .txt a .csv
folder = Path('Seccion_3/extensiones')

# Pasar de txt a csv
for path in list(folder.iterdir()):  # Iteramos sobre los elementos dentro de la carpeta 'extensiones'
    print(path)
    if path.suffix == '.txt':  # Verificamos si el archivo tiene la extensión '.txt'
        nuevoNombrePath = path.with_suffix('.csv')  # Creamos un nuevo Path con la extensión '.csv'
        path.rename(nuevoNombrePath)  # Renombramos el archivo cambiando su extensión a '.csv'

# Pasar de csv a txt
for path in folder.glob('**/*.csv'):  # Iteramos sobre los archivos con extensión '.csv' dentro de la carpeta 'extensiones'
    print(path)
    nuevoNombrePath = path.with_suffix('.txt')  # Creamos un nuevo Path con la extensión '.txt'
    path.rename(nuevoNombrePath)  # Renombramos el archivo cambiando su extensión a '.txt'
