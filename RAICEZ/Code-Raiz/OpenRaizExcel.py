import pandas as pd

# Especifica el motor para leer el archivo Excel

data = pd.read_excel(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', engine='xlrd')


# Lee el archivo Excel
data = pd.read_excel(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx')

# Muestra la lista de hojas disponibles en el archivo
print(data.sheet_names)

# Selecciona una hoja específica
hoja = data['Temp control 15feb 13Abr 2020']

# Muestra los datos en la hoja seleccionada
print(hoja)

# Selecciona columnas específicas para mostrar
columnas = hoja[['TtarRC_Avg(1)','TtarRC_Avg(2)','TtarRC_Avg(3)','TtarRC_Avg(4)','TtarRC_Avg(5)','TtarRC_Avg(6)','TtarRC_Avg(7)','TtarRC_Avg(8)'
                 ,'TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']]

# Muestra los datos de las columnas seleccionadas
print(columnas)