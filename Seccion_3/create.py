from pathlib import Path

numeros = [1,2,3,4,5,6,7,8,9]

# Creamos archivos de texto con nombres 'test1.txt', 'test2.txt', ..., 'test9.txt'
for i in numeros:
    with open(f'test{i}.txt', 'w') as f:
        f.write(f'Archivo {i}')

# Eliminar todos los archivos .txt creados
for path in Path('.').glob('*.txt'): # * Iteramos sobre todos los archivos con extensión .txt en el directorio actual
    print(path)  # Imprimimos el path del archivo que se va a eliminar
    path.unlink()  # Eliminamos el archivo

# Eliminar todos los archivos .txt menos el archivo 'test9.txt'
for path in Path('.').glob('test[1-8].txt'):  # Iteramos sobre todos los archivos con extensión .txt en el directorio actual
    print(path)  # Imprimimos el path del archivo que se va a eliminar
    path.unlink()  # Eliminamos el archivo