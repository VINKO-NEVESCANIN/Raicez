import pandas as pd

# Lee el archivo Excel
data = pd.read_excel(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', engine='xlrd')

# Muestra los datos en el archivo original
print(data)

# Filtra los datos específicos que deseas guardar
datos_filtrados = data[data['TtarRC_Avg(1)'] == '18,1256']

# Crea un nuevo archivo Excel para guardar los datos filtrados
archivo_guardado = pd.ExcelWriter('Temperaturas_2022', engine='xlsxwriter')

# Escribe los datos filtrados en el archivo guardado
datos_filtrados.to_excel(archivo_guardado, index=False)

# Guarda el archivo Excel
archivo_guardado.save()