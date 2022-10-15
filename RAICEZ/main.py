#Como leer un archivo de excel
import openpyxl

#leer el archivo
book = openpyxl.load_workbook('Excel/Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx', data_only=True)
#fijar la hoja 
hoja = book.active

celdas = hoja['AQ3': 'BF6640']

for fila in celdas:
    print([celda.value for celda in fila])#Compresion de listas
    
    #print([celda.value for celda in fila])