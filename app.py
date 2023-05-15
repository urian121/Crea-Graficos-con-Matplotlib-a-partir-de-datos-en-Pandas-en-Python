import pandas as pd
import mysql.connector
import openpyxl
import matplotlib.pyplot as plt


# Función para establecer la conexión a la base de datos
def connectionBD():
    try:
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            passwd="",
            database="demo",
            raise_on_warnings=True
        )
        if connection.is_connected():
            return connection

    except mysql.connector.Error as error:
        print(f"No se pudo conectar: {error}")

# Función para generar la gráfica de barras


def generar_grafica():
    # Se realiza la consulta a la base de datos
    with connectionBD() as conexion_MySQLdb:
        with conexion_MySQLdb.cursor(dictionary=True) as mycursor:
            querySQL = """SELECT id_persona, nombre, sexo, email, telefono, marca, company, saldo FROM personas"""
            mycursor.execute(querySQL,)
            dataBD = mycursor.fetchall()

    # Se convierte la consulta a un objeto DataFrame de pandas
    df = pd.DataFrame(dataBD)

    # Se genera la gráfica de barras con la librería matplotlib
    ax = df.plot(kind='bar', x='nombre', y='saldo', legend=False)

    # Se establece el título y las etiquetas de los ejes de la gráfica
    plt.title('Saldo por usuario')
    plt.xlabel('Usuarios')
    plt.ylabel('Saldo')

    # Se exporta la gráfica a un archivo PNG
    plt.savefig('grafica.png')

    # Se crea un objeto de libro de trabajo de Excel con la biblioteca openpyxl
    wb = openpyxl.Workbook()

    # Se crea una hoja de trabajo en el libro de Excel
    hoja = wb.active

    # O se cambia el nombre de la hoja de trabajo existente
    hoja.title = 'Hoja N° 1'

    # Crea la fila del encabezado con los títulos
    hoja.append(('N° Registro', 'Usuario', 'Nombre',
                 'Sexo', 'Email', 'Telefono', 'Compañia', 'Saldo'))

    # Se convierte el DataFrame de pandas a una lista de listas para insertarla en la hoja de trabajo
    data = df.values.tolist()

    # Se insertan los datos en la hoja de trabajo
    for r in data:
        hoja.append(r)

    # Se agrega la gráfica al libro de trabajo
    img = openpyxl.drawing.image.Image('grafica.png')
    hoja.add_image(img, 'J5')

    # Se guarda el archivo Excel
    wb.save('reporte_Excel_Python_Grafica.xlsx')


if __name__ == '__main__':
    generar_grafica()
