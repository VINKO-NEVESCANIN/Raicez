import pandas as pd

# Leer el archivo Excel
ruta_archivo = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'  # Reemplaza con la ubicación real de tu archivo
datos_excel = pd.read_excel(ruta_archivo)

# Seleccionar datos específicos
columnas_interes = ['TIMESTAMP','RECORD','DateTime(1)','DateTime(9)','DateTime(4)','DateTime(5)','Target','TtarRC_Avg(1)','TtarRC_Avg(2)','TtarRC_Avg(3)','TtarRC_Avg(4)','TtarRC_Avg(5)','TtarRC_Avg(6)','TtarRC_Avg(7)','TtarRC_Avg(8)','TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']  # Reemplaza con las columnas que te interesen
datos_seleccionados = datos_excel[columnas_interes]

# Mostrar los datos seleccionados
print(datos_seleccionados)


