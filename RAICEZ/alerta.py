from http import server
from multiprocessing import context
#Librerias para correos
import smtplib, ssl
import getpass
#Librerias para excel
import openpyxl

unsername = imput("Ingrese su nombre de usuario: ")
password = getpass.getpass("Ingrese su pasword: ")

#Crear la conexion
context = ssl.create_defaul_context()


#Leer el archivo
book = openpyxl.load_workbook('plantilla.xlsx', data_only=True)
#Fijar la hoja 
hoja = book.active

celdas = hoja ['AQ3': 'BF6640']
lista_raicez = []

for fila in celdas:
    raiz = [celda.value for celda in fila]
    lista_raicez.append(raiz)
    
with smtplib.SMTP_SSL("smtp.gmail.com," 465, context=context) as server:a
    server.login(unsername, password)
    print("Inicio sesion!")
    
    for raiz in lista_raicez:
        destinatario = raiz[3]
        mesaje = f'Hola {raiz[0]},la temperatura es {raiz[2]}'
        server.sendmail(username, destinatario, mesaje)
        print("Mesaje Enviado")    

