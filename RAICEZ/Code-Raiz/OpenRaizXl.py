import pandas as pd
import matplotlib.pyplot as plt

archivo_excel = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx'
columnas_deseadas = ['TtarRC_Avg(1)', 'TtarRC_Avg(2)', 'TtarRC_Avg(3)', 'TtarRC_Avg(4)', 'TtarRC_Avg(5)', 'TtarRC_Avg(6)', 'TtarRC_Avg(7)', 'TtarRC_Avg(8)',
                     'TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']  # Reemplaza con los nombres de tus columnas
archivo_salida = r'C:\Users\VINKO\Documents\GitHub\Raicez\Temperaturas_2022.xlsx'
#print(df.head())

try:
    df = pd.read_excel(archivo_excel, engine='openpyxl')
    df_seleccionado = df[columnas_deseadas]
    df_seleccionado.to_excel(archivo_salida, index=False)
    print('Datos guardados exitosamente en el archivo:', archivo_salida)
    print(df_seleccionado)
    # Realiza las operaciones necesarias con el DataFrame
    # ...
    
    # Crear gráfica de líneas
    #df_seleccionado.plot(x='TtarRC_Avg(1)', y=['TtarRC_Avg(2)', 'TtarRC_Avg(3)'], kind='line')
    #plt.xlabel('TtarRC_Avg(1)')
    #plt.ylabel('Valores')
    #plt.title('Gráfica de Líneas')
    #plt.legend()
    #plt.show()

except FileNotFoundError:
    print('El archivo no existe o la ruta es incorrecta.')
except Exception as e:
    print(f'Error al leer el archivo: {str(e)}')
    
    
    