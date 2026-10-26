import pandas as pd
import sys

NOTAS_ALUMNOS_PATH = r'Seccion_1/data/Notas_Alumnos.xlsx'  # Constante donde se guarda el path

# Diccionario que contiene las asignaturas y sus nombres correctos
dict_asig = {
    'LENGUA CASTELLANA Y LITERATURA': 'Lengua Castellana y Literatura',
    'BIOLOGIA': 'Biología',
    'GEOGRAFIA E HISTORIA': 'Geografía e Historia',
    'MATEMATICAS': 'Matemáticas',
    'INGLES': 'Inglés',
    'EDUCACION FISICA': 'Educación Física',
    'ETICA': 'Ética',
    'CULTURA CLASICA': 'Cultura clásica',
    'MUSICA': 'Música',
    'TECNOLOGIA': 'Tecnología',
    'EDUCACION PLASTICA': 'Educación Plástica',
    'FRANCES': 'Francés',
}

# Función que detecta errores en el DataFrame
def deteccionErrores(df):
    err1, err2, err3 = False, False, False  # Se inicializan las variables de error como False

    alumnos_list = sorted(list(df['NOMBRE'].drop_duplicates()))  # Se toma la columna del dataFrame y se eliminan duplicados, luego se ordena la lista
    asignatura_list = sorted(list(df['ASIGNATURA'].drop_duplicates())) # Se toma la columna del dataFrame y se eliminan duplicados, luego se ordena la lista

    for al in alumnos_list: # Por cada alumno en la  columna alumnos
        for asig in asignatura_list:    # Toma su asignatura de la columna asignaturas
            filt_al_as_df = df[(df['NOMBRE'] == al) & (df['ASIGNATURA'] == asig)]  # Se filtra el dataframe por alumno y asignatura
            print('')

            if filt_al_as_df.empty:  # Si el dataframe filtrado esta vacio
                err1 = True
                print(f"El alumno {al} no tiene la asignatura {asig} asignada")  # Se imprime un mensaje indicando que el alumno no tiene la asignatura

            elif(len(filt_al_as_df) > 1):  # Si el dataframe filtrado tiene mas de una fila
                err2 = True
                print(f"El alumno {al} tiene la asignatura {asig} {len(filt_al_as_df)} veces duplicada")  # Se imprime un mensaje indicando que el alumno tiene la asignatura duplicada
        
        for index, row in df.iterrows():  # Se itera sobre cada fila del dataframe
            trimestre_list = ['NOTA T1' , 'NOTA T2', 'NOTA T3']
            for trim in trimestre_list:
                if not((row[trim] >= 0.0) and (row[trim] <= 10.0)): # Chequea que las notas esten en el rango correcto
                    err3 = True
                    print(f"El alumno {al} tiene el campo {trim} de la asigantura{asig} fuera de rango {str(row[trim])}")

        # Chequeo de errores y finalizacion del programa en caso de que asi sea
        if err1 or err2 or err3:
            print('')
            print('Debes corregir los errores para continuar con la ejecucion del programa')
            sys.exit(1)  # Se termina la ejecucion del programa con un codigo de error 1
        else:
            print('Ningun error detectado')

def main():
    # Aqui se muestra como leer y printear una columna de un archivo del tipo excel usando pandas 
    excel_df = pd.read_excel(NOTAS_ALUMNOS_PATH, sheet_name='Notas')  # Se lee el path y sheetName del archivo excel y se guarda en un dataframe
    for index, row in excel_df.iterrows():  # Se itera sobre cada fila del dataframe
        print(f"Alumno: {row['NOMBRE']}")  # Se imprime el nombre del alumno

    # Aqui se toma las asignaturas y se eliminan duplicados
    asig_list = list(excel_df['ASIGNATURA'].drop_duplicates()) # Tomo la columna del dataFrame y elimino duplicados
    print(asig_list)

    # Aqui se ejecuta el codigo necesario para reformar las asignaturas a su correcto nombre, usando el diccionario dict_asig
    filter_td_asing = []
    for item in asig_list:
        valorTd = dict_asig[item]   # Se busca el valor del diccionario a partir de la clave (item) y se guarda en la variable valorTd
        filter_td_asing.append(valorTd)  # Se agrega el valor del diccionario a la lista
    print('')

    deteccionErrores(excel_df)


# Función principal que se ejecuta al correr el script
if __name__ == '__main__':
    main()