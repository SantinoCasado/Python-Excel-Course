from pathlib import Path
import zipfile

directorio_actual = Path('.')
directorio_objetivo = Path('Seccion_3/2026')

for path in directorio_actual.glob('*.zip'):  # Iteramos sobre todos los archivos con extensión .zip en el directorio actual
    print(path)  # Imprimimos el path del archivo zip que se va a descomprimir
    with zipfile.ZipFile(path, 'r') as zip_ref:  # Abrimos el archivo zip en modo lectura
        zip_ref.extractall(directorio_objetivo)  # Extraemos todos los archivos del zip al directorio objetivo