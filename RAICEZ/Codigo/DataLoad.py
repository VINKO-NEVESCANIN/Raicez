## Explorando el archivo de Excel ##

import openpyxl
import pandas as pd
#path = 

wb = openpyxl.load_workbook(
    r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx')
##'C:\ProyectosGIT\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'

print(wb.sheetnames)

df_TP = pd.read_excel(io = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', sheet_name='Temp control 15feb 13Abr 2020',header=0, names=None, index_col=None, usecols='AQ:BF', engine='openpyxl')

df_TP.head(3)

print(df_TP)

df_Modelo = pd.ExcelFile(
    r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx').parse(sheet_name='Temp control 15feb 13Abr 2020',header= 0,names=None,encoding = 'latin-1')

df_Modelo.head()

print(df_Modelo)

df_Modelo.to.to_excel('Clima_Escala.xlsx'sheet_name='Riegos',encoding= 'latin-1')