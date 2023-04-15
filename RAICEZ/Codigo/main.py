#Como leer un archivo de excel
import openpyxl
import matplotlib.pyplot as plt

#leer el archivo
book = openpyxl.load_workbook(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', data_only=True)
import pandas as pd
#fijar la hoja 
hoja = book.active

#celdas = hoja['AQ3': 'BF6640']

workbook1 = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'
df = pd.read_excel(workbook1)
valores = df[["TtarRC_Avg(1)", "TtarHC_Avg(1)"]]


print(valores)
ax = valores.plot.bar(x="TtarRC_Avg(1)", y="TtarHC_Avg(1)", rot=0)
plt.show()

#for fila in celdas:
 #   print([celda.value for celda in fila])#Compresion de listas
    
    #print([celda.value for celda in fila])