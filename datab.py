# Importación de las bibliotecas necesarias
import psycopg2  # Para interactuar con PostgreSQL
import easygui    # Para abrir una interfaz gráfica de usuario para seleccionar un archivo
import openpyxl   # Para manipular archivos de Excel

# Definición de la función experienciaEmpledo
def experienciaEmpledo():
    # Establece una conexión a la base de datos PostgreSQL
    conexion = psycopg2.connect("host=localhost dbname=DB user=postgres password=2003")
    
    # Crea un objeto cursor para ejecutar comandos SQL
    cur = conexion.cursor()
    
    # Abre un cuadro de diálogo para que el usuario seleccione un archivo de Excel
    var = str(easygui.fileopenbox())
    
    # Carga el libro de trabajo de Excel y selecciona la hoja activa
    book = openpyxl.load_workbook(var)
    hoja = book.active
    
    # Selecciona las celdas desde A2 hasta D4 en la hoja activa
    celdas = hoja['A2':'D4']
    
    # Inicializa una lista vacía para almacenar los datos de las herramientas
    lista = []
    
    # Itera sobre las filas y columnas seleccionadas y almacena los valores en la lista
    for fila in celdas:
        herramienta = [celda.value for celda in fila]
        lista.append(herramienta)
    
    # Imprime la lista de herramientas
    print(lista)
    
    # Itera sobre la lista de herramientas y ejecuta consultas SQL para insertar los datos en la tabla Herramientas
    i = 0
    for i in range(len(lista)):
        sql = 'INSERT INTO public.\"Herramientas\" (\"Nombre\",\"Porcentaje\",\"AniosExp\",\"IdEmpleado\") VALUES (%s,%s,%s,%s)'
        cur.execute(sql, lista[i])
        # Confirma los cambios en la base de datos
        conexion.commit()
    
    # Imprime el número de registros insertados
    print(f'Registros insertados: {i + 1}')
    
    # Cierra la conexión a la base de datos
    conexion.close()

# Llama a la función experienciaEmpledo para ejecutar el código
experienciaEmpledo()
