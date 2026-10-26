# 📊 Python-Excel-Course

Repositorio donde documento mi recorrido por un curso orientado a automatizar la **generación de reportes** con Python: leer y analizar datos en archivos **Excel** y volcar los resultados en documentos **Word**.

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-.xlsx-217346?logo=microsoftexcel&logoColor=white)
![Word](https://img.shields.io/badge/Word-.docx-2B579A?logo=microsoftword&logoColor=white)

---

## 📑 Índice

1. [Sobre el curso](#-sobre-el-curso)
2. [Objetivos de aprendizaje](#-objetivos-de-aprendizaje)
3. [Sección 1 — Lectura y validación de datos con pandas](#-sección-1--lectura-y-validación-de-datos-con-pandas)
4. [Sección 2 — Reportes en Word con plantillas (docxtpl)](#-sección-2--reportes-en-word-con-plantillas-docxtpl)
5. [Sección 3 — Automatización de archivos y carpetas con pathlib](#-sección-3--automatización-de-archivos-y-carpetas-con-pathlib)
6. [Sección 4 — Tablas dinámicas y gráficos en Excel con openpyxl](#-sección-4--tablas-dinámicas-y-gráficos-en-excel-con-openpyxl)
7. [Tecnologías y librerías utilizadas](#-tecnologías-y-librerías-utilizadas)
8. [Estructura del repositorio](#-estructura-del-repositorio)
9. [Cómo ejecutar los ejemplos](#-cómo-ejecutar-los-ejemplos)
10. [Conclusiones y próximos pasos](#-conclusiones-y-próximos-pasos)

---

## 🎯 Sobre el curso

El objetivo de este curso fue aprender cómo Python puede interactuar con archivos de tipo **Excel** y **Word** para generar reportes a partir del análisis de datos.

Este repositorio reúne el código, los ejercicios y mis apuntes de cada sección, organizados para poder consultarlos de forma rápida.

## 🧠 Objetivos de aprendizaje

- Leer, limpiar y analizar datos almacenados en planillas de Excel.
- Automatizar la creación de documentos Word a partir de esos datos.
- Combinar ambos mundos en un flujo de trabajo repetible: **datos → análisis → reporte**.
- Escribir scripts claros, reutilizables y fáciles de mantener.

---

## 📘 Sección 1 — Lectura y validación de datos con pandas

En esta sección trabajé con una planilla de notas de alumnos (`Notas_Alumnos.xlsx`, hoja `Notas`). Armé un script que **lee el Excel, normaliza los nombres de las asignaturas y detecta errores en los datos** antes de generar cualquier reporte. Validar primero evita producir documentos con información incorrecta.

**Qué aprendí**

- Leer un archivo Excel con `pd.read_excel()`, indicando la hoja con `sheet_name`, y trabajarlo como un `DataFrame`.
- Recorrer las filas con `iterrows()` y acceder a cada columna por su nombre.
- Obtener valores únicos con `drop_duplicates()` y ordenarlos con `sorted()`.
- Filtrar un `DataFrame` combinando condiciones con `&` (por ejemplo, alumno **y** asignatura).
- Corregir nombres con un diccionario de equivalencias (por ejemplo, `'BIOLOGIA'` → `'Biología'`).
- Validar la calidad de los datos y frenar el programa con `sys.exit(1)` si hay errores.
- Buenas prácticas: constante para la ruta del archivo, funciones separadas, `if __name__ == '__main__'` y f-strings para los mensajes.

**Validaciones que implementé**

| Error                               | Cómo lo detecto                                                       |
| ----------------------------------- | --------------------------------------------------------------------- |
| Alumno sin una asignatura asignada  | El `DataFrame` filtrado por alumno y asignatura está vacío (`.empty`) |
| Asignatura duplicada para un alumno | El filtrado devuelve más de una fila (`len(...) > 1`)                 |
| Nota fuera de rango                 | `NOTA T1`, `NOTA T2` o `NOTA T3` fuera del intervalo 0–10             |

**Qué utilicé**

- `pandas` para leer y analizar la planilla (para archivos `.xlsx` requiere `openpyxl` instalado).
- `sys` para terminar la ejecución con un código de error.

**Fragmento destacado**

```python
filt_al_as_df = df[(df['NOMBRE'] == al) & (df['ASIGNATURA'] == asig)]

if filt_al_as_df.empty:
    print(f"El alumno {al} no tiene la asignatura {asig} asignada")
elif len(filt_al_as_df) > 1:
    print(f"El alumno {al} tiene la asignatura {asig} {len(filt_al_as_df)} veces duplicada")
```

**Resultado / ejercicio práctico**
Script [`lecturaExcel.py`](Seccion_1/lecturaExcel.py): lee la planilla, lista alumnos y asignaturas, corrige los nombres de las materias y ejecuta la detección de errores. Si encuentra alguno, muestra qué corregir y se detiene.

[⬆ Volver al índice](#-índice)

---

## 📗 Sección 2 — Reportes en Word con plantillas (docxtpl)

Acá di el paso de analizar datos a **generar documentos**. A partir del Excel de notas (hojas `Datos_Alumnos` y `Notas`) y de una plantilla de Word (`Plantilla_Notas.docx`), el script genera **un informe de notas en `.docx` por cada alumno**, con el promedio final y una calificación con color para cada asignatura.

**Qué aprendí**

- Generar documentos Word desde una plantilla con `DocxTemplate` de `docxtpl`: armar un diccionario de contexto, aplicar `render(context)` y guardar con `save()`.
- Pasar a la plantilla datos simples (curso, nombre, clase) y también una **lista de diccionarios** (una entrada por asignatura).
- Leer varias hojas del mismo Excel con `pd.read_excel(..., sheet_name=...)`.
- Detectar alumnos asignados a más de una clase con `value_counts()`.
- Calcular la nota final como promedio de los tres trimestres, redondearla y clasificarla.
- Usar `copy.deepcopy` para no modificar el diccionario original al calcular resultados.
- Recargar la plantilla en cada iteración del bucle, para que cada documento parta de una plantilla limpia.
- Preparar la carpeta de salida con `os` y `shutil` (borrarla y volver a crearla en cada ejecución).
- Generar nombres de archivo seguros: mayúsculas, guiones bajos y sin tildes.

**Criterio de calificación**

| Promedio       | Calificación  | Color    |
| -------------- | ------------- | -------- |
| menor a 5      | SUSPENSO      | `ec7c7b` |
| 5 a menos de 7 | APROBADO      | `fbe083` |
| 7 a menos de 9 | NOTABLE       | `4db4d7` |
| 9 o más        | SOBRESALIENTE | `48bf91` |

**Qué utilicé**

- `pandas` y `openpyxl` para leer el Excel.
- `docxtpl` para completar la plantilla de Word.
- `os`, `shutil`, `copy` y `sys` para carpetas, copias de datos y control de errores.
- Un archivo de dependencias (`requeriments.txt`) con `pandas`, `openpyxl` y `docxtpl`.

**Fragmento destacado**

```python
docs_tpl = DocxTemplate(PLANTILLA_CURSOS_PATH)  # se recarga la plantilla en cada iteración

context = {
    'curso': CURSO,
    'nombre_alumno': nombre_alumno,
    'clase': clase,
    'asignatura_list': asignatura_list,
}

docs_tpl.render(context)
docs_tpl.save(os.path.join(PATH_OUTPUT, titulo))
```

**Resultado / ejercicio práctico**
Script [`plantillaWord.py`](Seccion_2/plantillaWord.py): valida los datos, limpia la carpeta `outputs` y genera un informe `NOTAS_<ALUMNO>.docx` por cada alumno.

[⬆ Volver al índice](#-índice)

---

## 📙 Sección 3 — Automatización de archivos y carpetas con pathlib

En esta sección aprendí a **manejar el sistema de archivos desde Python**: crear estructuras de carpetas, buscar archivos por patrón, renombrarlos, eliminarlos y descomprimirlos. También incorporé datos desde la web hacia Excel y CSV. Son las tareas de "preparación" que rodean a cualquier reporte automatizado.

**Scripts de la sección**

| Script                | Qué hace                                                                                    |
| --------------------- | ------------------------------------------------------------------------------------------- |
| `create.py`           | Crea `test1.txt` a `test9.txt` y los elimina usando patrones de búsqueda                    |
| `main.py`             | Genera una estructura de carpetas `2026/<mes>/<día>` usando el módulo `calendar`            |
| `rename.py`           | Renombra carpetas y archivos, cambia extensiones (`.txt` ↔ `.csv`) y recorre subdirectorios |
| `zips.py`             | Descomprime todos los `.zip` del directorio actual en `Seccion_3/2026`                      |
| `extractHtmlTable.py` | Importa una tabla de Wikipedia (países por población) y la guarda en CSV y Excel            |

**Qué aprendí**

| Tarea                              | Herramienta                                              |
| ---------------------------------- | -------------------------------------------------------- |
| Crear carpetas anidadas            | `Path(...).mkdir(parents=True, exist_ok=True)`           |
| Buscar archivos por patrón         | `glob('*.txt')`, `glob('test[1-8].txt')`, `glob('**/*')` |
| Eliminar archivos                  | `path.unlink()`                                          |
| Renombrar archivos y carpetas      | `path.rename()` y `path.with_name()`                     |
| Cambiar la extensión               | `path.suffix` y `path.with_suffix('.csv')`               |
| Distinguir archivos de carpetas    | `is_file()` y `is_dir()`                                 |
| Listar el contenido de una carpeta | `iterdir()`                                              |
| Descomprimir ZIP                   | `zipfile.ZipFile(...).extractall()`                      |
| Importar tablas de una web         | `pd.read_html()`                                         |
| Exportar datos                     | `to_csv(index=False)` y `to_excel(index=False)`          |

**Qué utilicé**

- `pathlib` como herramienta principal para trabajar con rutas.
- `zipfile` para archivos comprimidos.
- `calendar` para los nombres de los meses.
- `pandas` para importar y exportar tablas (`read_html` necesita un parser como `lxml`).

**Fragmento destacado**

```python
meses_ano = list(calendar.month_name)[1:]
dias = ['Dia 1', 'Dia 10', 'Dia 20', 'Dia 30']

for i, mes in enumerate(meses_ano):
    for dia in dias:
        Path(f'Seccion_3/2026/{i+1}.{mes}/{dia}').mkdir(parents=True, exist_ok=True)
```

**Resultado / ejercicio práctico**
Un conjunto de scripts para organizar, renombrar y descomprimir archivos, y para convertir una tabla de una página web en un archivo CSV y otro Excel listos para analizar.

[⬆ Volver al índice](#-índice)

---

## 📕 Sección 4 — Tablas dinámicas y gráficos en Excel con openpyxl

En la última sección llevé el análisis de datos a un **reporte final en Excel**. Partí de un archivo de ventas de supermercado (`supermarket_sales.xlsx`), armé una tabla dinámica con pandas y luego, con openpyxl, abrí el libro generado para agregarle un gráfico de barras.

**Qué aprendí**

- Resumir datos con `pivot_table`: filas (`index`), columnas (`columns`), valores (`values`) y función de agregación (`aggfunc`).
- Redondear los resultados con `round()` y exportar la tabla con `to_excel()`, eligiendo la hoja (`sheet_name`) y la fila de inicio (`startrow`) para dejar espacio arriba.
- Reabrir un libro existente con `load_workbook` y seleccionar una hoja por su nombre.
- Obtener los límites de los datos con `min_row`, `max_row`, `min_column` y `max_column`, para no fijar rangos a mano.
- Definir rangos con `Reference` para los datos y las categorías del gráfico.
- Crear un `BarChart`, ponerle título y estilo, y ubicarlo en una celda concreta de la hoja.

**Qué utilicé**

- `pandas` para la tabla dinámica y la exportación.
- `openpyxl` (`load_workbook`, `BarChart`, `Reference`) para modificar el libro y crear el gráfico.

**Fragmento destacado**

```python
table_pivote = df.pivot_table(
    index='Gender', columns='Product line', values='Total', aggfunc='sum'
).round(0)

table_pivote.to_excel('Seccion_4/data/supermarket_sales_pivot.xlsx',
                      startrow=4, sheet_name='Report')
```

**Resultado / ejercicio práctico**
Script [`main.py`](Seccion_4/main.py): genera `supermarket_sales_pivot.xlsx` con una hoja `Report` que contiene las ventas totales por género y línea de producto, más un gráfico de barras titulado "Ventas".

[⬆ Volver al índice](#-índice)

---

## 🛠 Tecnologías y librerías utilizadas

| Herramienta | Uso                                                                                |
| ----------- | ---------------------------------------------------------------------------------- |
| Python      | Lenguaje principal                                                                 |
| pandas      | Lectura, filtrado y análisis de datos; tablas dinámicas; importación de tablas web |
| openpyxl    | Lectura de `.xlsx`, modificación de libros y gráficos de Excel                     |
| docxtpl     | Generación de documentos Word a partir de una plantilla                            |
| pathlib     | Rutas, carpetas, búsqueda por patrón y renombrado de archivos                      |
| zipfile     | Descompresión de archivos `.zip`                                                   |
| os / shutil | Creación y borrado de carpetas                                                     |
| calendar    | Nombres de los meses del año                                                       |
| copy        | Copias profundas de diccionarios                                                   |
| sys         | Finalizar el programa con código de error                                          |

## 🗂 Estructura del repositorio

```text
Python-Excel-Course/
├── Seccion_1/
│   ├── data/
│   │   └── Notas_Alumnos.xlsx
│   └── lecturaExcel.py
├── Seccion_2/
│   ├── data/
│   │   ├── Notas_Alumnos.xlsx
│   │   └── Plantilla_Notas.docx
│   ├── outputs/
│   ├── plantillaWord.py
│   └── requeriments.txt
├── Seccion_3/
│   ├── create.py
│   ├── extractHtmlTable.py
│   ├── main.py
│   ├── rename.py
│   └── zips.py
├── Seccion_4/
│   ├── data/
│   │   └── supermarket_sales.xlsx
│   └── main.py
└── README.md
```

## ▶️ Cómo ejecutar los ejemplos

```bash
# 1. Clonar el repositorio
git clone <url-del-repositorio>
cd Python-Excel-Course

# 2. (Opcional) Crear un entorno virtual
python -m venv venv
source venv/bin/activate      # En Windows: venv\Scripts\activate

# 3. Instalar dependencias (pandas, openpyxl y docxtpl)
pip install -r Seccion_2/requeriments.txt
pip install lxml    # necesario para pd.read_html (Sección 3)

# 4. Ejecutar los scripts desde la raíz del repositorio
#    (las rutas de los datos son relativas a la raíz)
python Seccion_1/lecturaExcel.py
python Seccion_2/plantillaWord.py
python Seccion_3/extractHtmlTable.py
python Seccion_4/main.py
```

## ✅ Conclusiones y próximos pasos

Este curso me permitió recorrer el ciclo completo de un reporte automatizado: **leer y validar los datos** (Sección 1), **generar documentos a partir de ellos** (Sección 2), **preparar y organizar los archivos** (Sección 3) y **resumir la información en tablas y gráficos de Excel** (Sección 4). Lo más valioso fue entender que validar los datos antes de generar cualquier salida ahorra errores, y que una plantilla bien definida permite producir decenas de documentos con el mismo código.

**Ideas para seguir**

- Unir la validación, la generación de documentos Word y los gráficos en un único flujo de trabajo.
- Aplicar lo aprendido a otros conjuntos de datos y otros tipos de reportes.

---

<p align="center">Hecho con 🐍 por <b>Santino</b></p>
