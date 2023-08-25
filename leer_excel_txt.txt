import pandas as pd
import mysql.connector

ruta_excel = 'Control Siemens Q2-2023-global-Wk19 anonimizado.xlsx'     #ruta donde este el excel
datos_excel = pd.read_excel(ruta_excel)                                 #traduce el excel para meterlo en la variable
fila_final=[]

mydb = mysql.connector.connect(                                         #creo la conexion con la base de datos
  host="127.0.0.1",
  user="root",
  password="",
  database="LEER_EXEL"
)

cursor = mydb.cursor()                                                  #Creamos un objeto "cursor" para ejecutar consultas SQL en la base de datos

filas = datos_excel.shape[0]                                            # Obtener el número de filas de la matriz
columnas = datos_excel.shape[1]                                         # Obtener el número de columnas de la matriz
fila_inicio = 29                                                        #decido donde empezar
columna_inicio = 3                                                      #aqui igual
columna_fin = 6                                                         #es hasta donde quiero llegar
filas_fin=33                                                            #empiezo aqui porq he visto que los 4 usuarios son diferentes en las siguentes filas
contador_recorre_filas=fila_inicio                                      #esta variable es la que uso para sacar el valor de la celda que recorro y autoincremento en la funcion
contador_recorre_col=columna_inicio                                     #lo mismo pero con las columnas
conttt_f=0                                                              #esta variable solo la he creado para ver las iteraciones que hago, es mi modo de depurar
contador_celdas=0                                                       #cuento las celdas que he leido en el total de las iteraciones, esto estaria bien cuando se haga la funcion para un excle grande, pienso
print(datos_excel.iloc[29, 3])

def obtener_valores_excel(filas_fin, columna_inicio, columna_fin, contador_recorre_filas, contador_recorre_col, conttt_f, contador_celdas):
    if contador_recorre_filas < filas_fin:
        if contador_recorre_col <= columna_fin:
            valor=datos_excel.iloc[contador_recorre_filas, contador_recorre_col]#guardo la celda
            fila_final.append(valor)                                    #le meto la celda a la cadena que luego voy a insertar cuando termine de recorrer la fila
            contador_recorre_col +=1                                    #aumento la columna para la siguiente iteracion
            contador_celdas +=1                                         #bueno, cuento aqui las veces que entra aquí, por curiosidad simplemente
            obtener_valores_excel(filas_fin, columna_inicio, columna_fin, contador_recorre_filas, contador_recorre_col, conttt_f, contador_celdas)
                                                                        #llamo a la funcion otra vez
        else:
            print(fila_final)                                           #cuando ha terminado las columnas de esa fila significa que ya tengo la fila para insertar, la muestro por pantalla aqui para simplemente ver que esté bien
            sql = "INSERT INTO usuarios (apellidos, nombre, siglum, posicion) VALUES (%s, %s, %s ,%s)" #Escribe la instrucción SQL para realizar la inserción. Utiliza la sintaxis adecuada de acuerdo con la base de datos que estés utilizando. Aquí tienes un ejemplo de cómo insertar datos en una tabla llamada "usuarios"
            val = (fila_final[0],fila_final[1],fila_final[2],fila_final[3])#van a ser los valores de los %s
            cursor.execute(sql, val)                                    #Ejecuta la consulta SQL utilizando el cursor y pasandole los valores que he metido en la variable fila_final
            mydb.commit()                                               #hago el comit para que se realicen todas las operaciones (insert en este caso)
            contador_recorre_col=columna_inicio                         #como he terminado de recorrer las columnas, reinicio el recorre_colum
            contador_recorre_filas +=1                                  #como aun no he terminado de recorrer las filas, pues la aumento 1
            fila_final.clear()                                          #una vez hago el insert, voy a limpiar la la fila que inserto
            #print(fila_final)
            conttt_f +=1                                                #tambien cuento las filas que hay
            obtener_valores_excel(filas_fin, columna_inicio, columna_fin, contador_recorre_filas, contador_recorre_col, conttt_f, contador_celdas)
                                                                        #vuelvo a llamar a la funcion ya que aun no he terminado las filas
    else:
                                                                        #cuando entro aqui significa que ya ha terminado, deja de llamarse y simplemente muestro resultados
        print()
        print(f"has realizado {conttt_f} saltos de filas")
        print()
        print(f"el total de las celdas que hemos insertado en la base de datos es : {contador_celdas}")
        print()
        print("Operacion realizada correctamente")
        print()

obtener_valores_excel(filas_fin, columna_inicio, columna_fin, contador_recorre_filas, contador_recorre_col, conttt_f, contador_celdas)
mydb.close()                                                            #cierro conexion