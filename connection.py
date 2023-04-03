from openpyxl import *
import mysql.connector
import openpyxl
from mysql.connector import Error
import os


try:
    conexion = mysql.connector.connect(
        host='localhost',
        port=3306,
        user='root',
        password='',
        db='testv1'
    )

    if conexion.is_connected():
        print("Conexión exitosa.")
        with conexion.cursor() as cursor:
            # En este caso no necesitamos limpiar ningún dato
            cursor.execute("SELECT Producto, Descrip, sum(Cantidad), cast(FADUA as date)   "    +
            "from ticketsdetalle where FADUA >= '2023-01-01' and FADUA <= '2023-03-31' " +
            "group by Cantidad, FADUA " +
            "order by FADUA ;")

            # Con fetchall traemos todas las filas
            data = cursor.fetchall()

            # Recorrer e imprimir
            producto = []
            desc= []
            cantidad = []
            fecha = []
            #input(print(data))
            ruta_actual = os.path.dirname(os.path.abspath(__file__))

            # Crear la ruta completa de la carpeta de destino y crearla si no existe
            ruta_destino = os.path.join(ruta_actual, "ArchivosExcel")
            if not os.path.exists(ruta_destino):
                os.makedirs(ruta_destino)

            # Crear el nombre completo del archivo a guardar
            nombre_archivo = os.path.join(ruta_destino, "productos.xlsx")


            wb = openpyxl.Workbook()
            hoja = wb.active
            hoja.append(('Producto', 'Descripción', 'Venta', 'Fecha'))
            # Crear una lista de tuplas con los datos
            datos_filas = [(datos[0], datos[1], datos[2], datos[3]) for datos in data]

            # Agregar la lista completa como una fila en la hoja de trabajo
            for fila in datos_filas:
                hoja.append(fila)
            try:
                wb.save(nombre_archivo)
                            #print(datos[0])
            except:
                print("Cierre alrchivo para continuar")

            
            
       
except Error as ex:
    print("Error durante la conexión: {}".format(ex))
finally:
    if conexion.is_connected():
        conexion.close()  # Se cerró la conexión a la BD.
        print("La conexión ha finalizado.")
