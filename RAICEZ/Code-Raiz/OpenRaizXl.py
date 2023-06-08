import pandas as pd

archivo_excel = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx'
columnas_deseadas = ['TtarRC_Avg(1)', 'TtarRC_Avg(2)', 'TtarRC_Avg(3)', 'TtarRC_Avg(4)', 'TtarRC_Avg(5)', 'TtarRC_Avg(6)', 'TtarRC_Avg(7)', 'TtarRC_Avg(8)',
                     'TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']  # Reemplaza con los nombres de tus columnas


try:
    df = pd.read_excel(archivo_excel, engine='openpyxl')
    # Realiza las operaciones necesarias con el DataFrame
    # ...
    print(df.head())
    df_seleccionado = df[columnas_deseadas]
    print(df_seleccionado)
except FileNotFoundError:
    print('El archivo no existe o la ruta es incorrecta.')
except Exception as e:
    print(f'Error al leer el archivo: {str(e)}')
    
    
    