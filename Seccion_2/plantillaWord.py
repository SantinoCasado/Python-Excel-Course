import pandas as pd
import sys
import os
import shutil
import copy
from docxtpl import DocxTemplate    # Libreria de word para generar documentos a partir de plantillas

NOTAS_ALUMNOS_PATH = r'Seccion_2/data/Notas_Alumnos.xlsx'  # Constante donde se guarda el path
PLANTILLA_CURSOS_PATH = r'Seccion_2/data/Plantilla_Notas.docx'  # Constante donde se guarda el path
PATH_OUTPUT = r'Seccion_2/outputs'
CURSO = '2025/2026'

# Colores
SUSPENSO_COLOR = 'ec7c7b'
APROBADO_COLOR = 'fbe083'
NOTABLE_COLOR = '4db4d7'
SOBRESALIENTE_COLOR = '48bf91'

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


def deteccionErrores(df):
    err = False

    # Se cuenta cuántas filas (asignaciones de clase) tiene cada alumno en el dataframe.
    conteo = df['NOMBRE'].value_counts()

    # Se recorre cada alumno junto con la cantidad de filas que tiene asociadas
    for alumno, cantidad in conteo.items():

        # Si un alumno aparece más de una vez, significa que está asignado a más de una clase
        if cantidad > 1:
            err = True  

            # Se obtienen todas las clases a las que quedó asignado ese alumno, para mostrarlas en el mensaje de error
            clases = df[df['NOMBRE'] == alumno]['CLASE'].tolist()
            print(f"El alumno {alumno} está asignado a más de una clase: {clases}")

    # Chequeo final: si se detectó al menos un error, se corta la ejecución del programa
    if err:
        print('')
        print('Debes corregir los errores para continuar con la ejecucion del programa')
        sys.exit(1)  # Código de salida 1 indica finalización por error
    else:
        # Si no hubo errores, se informa y el programa puede continuar normalmente
        print('Ningun error detectado')

def eliminarTildes(texto):
    tildes_dict = {
        "Á" : "A",
        "á" : "a",
        "É" : "E",
        "é" : "e",
        "Í" : "I",
        "í" : "i",
        "Ó" : "O",
        "ó" : "o",
        "Ú" : "U",
        "ú" : "u"
    }

    textoSinTilde = texto
    for key in tildes_dict:
        textoSinTilde = textoSinTilde.replace(key, tildes_dict[key])

    return textoSinTilde  # Se devuelve el texto ya sin tildes

def eliminarCrearCarpetas(path):
    if os.path.exists(path):
        shutil.rmtree(path)  # Elimina la carpeta y todo su contenido
    os.mkdir(path)

def reformarNombresAsig(asig_list):
    filter_td_asig = []
    for item in asig_list:
        valorTd = dict_asig[item]   # Se busca el valor del diccionario a partir de la clave (item) y se guarda en la variable valorTd
        filter_td_asig.append(valorTd)  # Se agrega el valor del diccionario a la lista
    return filter_td_asig  # Se devuelve la lista de asignaturas reformadas

def ObtenerNotaFinal(dict_asignatura):
    newAsignaturaDic = copy.deepcopy(dict_asignatura)  # Se hace una copia profunda del diccionario para no modificar el original
    TRIMESTRE_LIST = ['t1', 't2', 't3']

    if any(newAsignaturaDic[trim] is None for trim in TRIMESTRE_LIST):  # Si alguna de las notas de los trimestres es None
        newAsignaturaDic['nota_final'] = None  # La nota final también será None
    else:
        # Obtengo la nota final
        nota_media = 0
        for trim in TRIMESTRE_LIST:
            nota_media += newAsignaturaDic[trim]
        nota_media /= 3
        newAsignaturaDic['nota_final'] = round(nota_media, 1)  # Se calcula la nota final como el promedio de las notas de los tres trimestres y se redondea a un decimal

        # Obtengo la clasificacion
        if(nota_media < 5.0):
            calif = 'SUSPENSO'
            color_calif = SUSPENSO_COLOR
        elif(nota_media < 7.0):
            calif = 'APROBADO'
            color_calif = APROBADO_COLOR
        elif(nota_media < 9.0):
            calif = 'NOTABLE'
            color_calif = NOTABLE_COLOR
        else:
            calif = 'SOBRESALIENTE'
            color_calif = SOBRESALIENTE_COLOR

        newAsignaturaDic['calificacion'] = calif  # Se asigna la clasificación correspondiente a la nota final
        newAsignaturaDic['color'] = color_calif

    return newAsignaturaDic  # Se devuelve el diccionario con la nota final y la clasificación
    
def generarReportesNotas(datos_alumnos_df, excel_df):
    asig_list = sorted(list(excel_df['ASIGNATURA'].drop_duplicates()))  # Se toma y ordena la lista de asignaturas pertenecientes a la columna "ASIGNATURA"
    filter_td_asig = reformarNombresAsig(asig_list)

    nombre_alumno_list = sorted(list(datos_alumnos_df['NOMBRE'])) # Se toma y ordena la lista de alumnos pertenecientes a la columna "NOMBRE"

    # --- Aqui se genera el reporte de notas para cada alumno, usando la plantilla de word y los datos del excel ---
    for nombre_alumno in nombre_alumno_list:
        docs_tpl = DocxTemplate(PLANTILLA_CURSOS_PATH)  # se recarga la plantilla en cada iteración

        filt_datos_alumnos_df = datos_alumnos_df[(datos_alumnos_df['NOMBRE'] == nombre_alumno)] # Filtra los alumnos por su nombre
        clase = filt_datos_alumnos_df.iloc[0]['CLASE']  # Se toma la clase del primer registro del alumno filtrado (asumiendo que todos los registros del mismo alumno tienen la misma clase)

        # Crear tabla de Notas
        asignatura_list = []

        #Iterar sobre los indicies de asignaturas
        for asig_index in range(len(asig_list)):
            asig = asig_list[asig_index]  # Se toma la asignatura correspondiente al índice actual
            filt_al_as_excel_df = excel_df[(excel_df['NOMBRE'] == nombre_alumno) & (excel_df['ASIGNATURA'] == asig)]  # Se filtra el dataframe de notas por alumno y asignatura

            asignatura_dict = {
                'nombre_asignatura': filter_td_asig[asig_index],  # Se toma el nombre de la asignatura reformado
                't1' : round(filt_al_as_excel_df.iloc[0]['NOTA T1'], 1) if not filt_al_as_excel_df.empty else None,  # Se toma la nota del primer trimestre si existe, sino None
                't2' : round(filt_al_as_excel_df.iloc[0]['NOTA T2'], 1) if not filt_al_as_excel_df.empty else None,  # Se toma la nota del segundo trimestre si existe, sino None
                't3' : round(filt_al_as_excel_df.iloc[0]['NOTA T3'], 1) if not filt_al_as_excel_df.empty else None,  # Se toma la nota del tercer trimestre si existe, sino None
            }

            asignatura_dict = ObtenerNotaFinal(asignatura_dict)
            asignatura_list.append(asignatura_dict)  # Se agrega el diccionario de la asignatura a la lista de asignaturas


        context = {
            'curso': CURSO,
            'nombre_alumno': nombre_alumno,
            'clase': clase,
            'asignatura_list' : asignatura_list,

        }

        docs_tpl.render(context)
        titulo = eliminarTildes(('NOTAS_' + nombre_alumno.replace(' ', '_')).upper() + '.docx')
        docs_tpl.save(os.path.join(PATH_OUTPUT, titulo))


def main():
    datos_alumnos_df = pd.read_excel(NOTAS_ALUMNOS_PATH, sheet_name='Datos_Alumnos')  # Se lee el path y sheetName del archivo excel y se guarda en un dataframe
    excel_df = pd.read_excel(NOTAS_ALUMNOS_PATH, sheet_name='Notas')  # Se lee el path y sheetName del archivo excel y se guarda en un dataframe

    eliminarCrearCarpetas(PATH_OUTPUT) # Se llama a la función para eliminar y crear la carpeta de outputs
    deteccionErrores(datos_alumnos_df)  # Se llama a la función de detección de errores con el dataframe leído
    generarReportesNotas(datos_alumnos_df, excel_df)

if __name__ == '__main__':
    main()