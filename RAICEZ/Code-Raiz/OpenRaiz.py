import pandas as pd

archivo_excel = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx'
columnas = ['TtarRC_Avg(1)', 'TtarRC_Avg(2)', 'TtarRC_Avg(3)', 'TtarRC_Avg(4)', 'TtarRC_Avg(5)', 'TtarRC_Avg(6)', 'TtarRC_Avg(7)', 'TtarRC_Avg(8)',
            'TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']  # Reemplaza con los nombres de las columnas deseadas
archivo_salida = 'Temperaturas_2022.xlsx'  # Nombre del archivo de salida

try:
    df = pd.read_excel(archivo_excel, engine='openpyxl')

    # Crear un nuevo DataFrame para almacenar los datos seleccionados
    df_seleccionados = pd.DataFrame()

    for columna in columnas:
        # Agregar la columna seleccionada al DataFrame de datos seleccionados
        df_seleccionados[columna] = df[columna]

        # Imprimir los datos de la columna seleccionada
        print(f'Datos de la columna {columna}:')
        print(df[columna])
        print()

    # Guardar los datos seleccionados en un nuevo archivo Excel
    df_seleccionados.to_excel(archivo_salida, index=False)

    print('Los datos seleccionados se han guardado correctamente en el archivo:', archivo_salida)

except FileNotFoundError:
    print('El archivo no existe o la ruta es incorrecta.')
except Exception as e:
    print(f'Error al leer el archivo: {str(e)}')