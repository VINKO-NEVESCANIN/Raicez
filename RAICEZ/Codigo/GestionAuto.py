import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

#Poner la ruta del acrivo excel a consultar
archivo_excel = pd.read_excel(
    'C:\ProyectosGIT\Raicez\RAICEZ\Excel\Temperatura_control.xlsx')

#archivo_excel[['TtarRC_Avg(1)', 'TtarRC_Avg(2)','TtarRC_Avg(3)', 'TtarRC_Avg(4)', 'TtarRC_Avg(5)', 'TtarRC_Avg(6)', 'TtarRC_Avg(7)', 'TtarRC_Avg(8)', 'TtarHC_Avg(1)', 'TtarHC_Avg(2)', 'TtarHC_Avg(3)', 'TtarHC_Avg(4)', 'TtarHC_Avg(5)', 'TtarHC_Avg(6)', 'TtarHC_Avg(7)', 'TtarHC_Avg(8)']]

Tabla_pivote = archivo_excel.pivot_table(
    index='Date', columns='TIMESTAMP', values='TtarRC_Avg(1)', aggfunc='sum').round(0)

print(Tabla_pivote)

Tabla_pivote.to_excel('Temperaturas_2022.xlsx',startrow=3,sheet_name='Report')

wb = load_workbook('Temperaturas_2022.xlsx')
pestana = wb['Report']

min_col = wb.active.min_column
max_col = wb.active.max_column
min_fila = wb.active.min_row
max_fila = wb.active.max_row

print(min_col)
print(max_col)
print(min_fila)
print(max_fila)


#Grafico
barchart = BarChart()
data = Reference(pestana, min_col=min_col+1, max_col=max_col,min_row=min_fila,max_row=max_fila)
categorias = Reference(pestana, min_col=min_col+1, max_col=max_col,min_row=min_fila, max_row=max_fila)

barchart.add_data(data, tittles_from_data=True)
barchart.set_categories(categorias)

pestana.add_chart(barchart, 'B12')
barcahrt.title = 'Ventanas'
barchart.style = 2 #5

pestana['B8'] = '=SUM(B6:B7)'
pestana['B8'].style = 'currency'

pestana['C6'] = '=SUM(B6:B7)'
pestana['C7'].style = 'currency'

abecedario = list(string.ascii_uppercase)
print(abecedario[0:max_col])

wb.save('Temperaturas_2022.xlsx')





#TIMESTAMP
