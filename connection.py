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
            "group by Producto, FADUA " +
            "order by FADUA ;")

            # Con fetchall traemos todas las filas
            data = cursor.fetchall()

            # Recorrer e imprimir
            producto = []
            desc = []
            cantidad = []
            fecha = []

            ruta_actual = os.path.dirname(os.path.abspath(__file__))

            # Crear la ruta completa de la carpeta de destino y crearla si no existe
            ruta_destino = os.path.join(ruta_actual, "ArchivosExcel")
            if not os.path.exists(ruta_destino):
                os.makedirs(ruta_destino)

            # Crear el nombre completo del archivo a guardar
            nombre_archivo = os.path.join(ruta_destino, "productos.xlsx")
            nombre_archivo_2 = os.path.join(ruta_destino, "almacen.xlsx")

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
            except:
                print("Cierre archivo para continuar")

        with conexion.cursor() as historico:
            # En este caso no necesitamos limpiar ningún dato
            historico.execute("select Producto, ExistenciaActual from historicoalmacen")

            # Con fetchall traemos todas las filas
            datahist = historico.fetchall()

            producto_2 = []
            cntidad_2 = []
            wb_2 = openpyxl.Workbook()
            hoja_2 = wb_2.active
            hoja_2.append(('Producto', 'Cantidad'))
            # Crear una lista de tuplas con los datos
            datos_filas_2 = [(datos[0], datos[1]) for datos in datahist]

            for fila_2 in datos_filas_2:
                hoja_2.append(fila_2)

            try:
                wb_2.save(nombre_archivo_2)
            except:
                print("Cierre archivo para continuar")
except Exception as e:
    print(f"Ocurrió un error: {str(e)}")
finally:
    if conexion.is_connected():
        cursor.close()
        conexion.close()
        print("Conexión cerrada.")

            
