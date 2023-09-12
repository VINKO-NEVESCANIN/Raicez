import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import Font
import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
from tkinter import *


archivo_excel = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx'
df = pd.read_excel(archivo_excel, engine='openpyxl')
print(df.head)


# Definir cajas de entrada para los rangos de las columnas como variables globales
cajas_rango_columnas = []
df = None
rango_columna1_min = 0.0
rango_columna1_max = float('inf')
rango_columna2_min = 0.0
rango_columna2_max = float('inf')



def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    if ruta_archivo:
        entry_ruta_archivo.delete(0, tk.END)
        entry_ruta_archivo.insert(tk.END, ruta_archivo)
        

def procesar_datos():
    global cajas_rango_columnas  # Para poder acceder a las cajas desde otras funciones
    global df

    archivo = entry_ruta_archivo.get()
    columnas_seleccionadas = entry_columnas.get().split(',')
    rango_fechas = pd.date_range(start=entry_fecha_inicio.get(), end=entry_fecha_fin.get())

    # Leer el archivo de Excel original
    df = pd.read_excel(archivo)

    # Filtrar los datos por las columnas seleccionadas y el rango de fechas
    # df_seleccionado = df[df['Fecha'].isin(rango_fechas)][columnas_seleccionadas]

    # Crear la interfaz gráfica
    root = Tk()
    root.title("Aplicación")

    # Nombres de las columnas
    nombres_columnas = ['Columna1', 'Columna2', 'Columna3', 'Columna4']

    # Crear cajas de entrada y etiquetas para cada columna
    for nombre_columna in nombres_columnas:
        label = Label(root, text=f'Rango para {nombre_columna}:')
        label.pack()

        caja_min = Entry(root)
        caja_max = Entry(root)

        caja_min.pack()
        caja_max.pack()

        cajas_rango_columnas.append((caja_min, caja_max))

    # Botón para realizar el filtrado
    boton_filtrar = Button(root, text="Filtrar", command=filtrar_datos)
    boton_filtrar.pack()

def filtrar_datos():
    global cajas_rango_columnas  # Para acceder a las cajas desde esta función

    # Obtener los parámetros ingresados desde la interfaz gráfica
    rangos = []

    for caja_min, caja_max in cajas_rango_columnas:
        rango_min_text = caja_min.get()
        rango_max_text = caja_max.get()

        try:
            rango_min = float(rango_min_text) if rango_min_text else 0.0
        except ValueError:
            rango_min = 0.0

        try:
            rango_max = float(rango_max_text) if rango_max_text else float('inf')
        except ValueError:
            rango_max = float('inf')

        rangos.append((rango_min, rango_max))

    # Filtrar los datos según los rangos de valores
    for i, (rango_min, rango_max) in enumerate(rangos):
        columna = f'Columna{i + 1}'
        df_filtrado = df[(df[columna] >= rango_min) & (df[columna] <= rango_max)]
            

    # Guardar los datos filtrados en un nuevo archivo de Excel
    archivo_filtrado = 'archivo_filtrado.xlsx'
    df_filtrado.to_excel(archivo_filtrado, index=False)
    
    # Mostrar la gráfica de los datos
    plt.plot(df_filtrado['Columna1'], df_filtrado['Columna2'])
    plt.xlabel('Columna1')
    plt.ylabel('Columna2')
    plt.title('Gráfica de Columna1 vs Columna2')
    plt.show()
    
    # Generar la gráfica de los datos filtrados
    plt.figure(figsize=(8, 6))
    plt.plot(df_filtrado['Columna1'], label='Columna1')
    plt.plot(df_filtrado['Columna2'], label='Columna2')
    plt.xlabel('Índice')
    plt.ylabel('Valores')
    plt.title('Gráfica de datos filtrados')
    plt.legend()
    plt.grid(True)
    plt.savefig('grafica_datos_filtrados.png')  # Guardar la gráfica como imagen

    # Resaltar los valores numéricos en el archivo Excel filtrado
    wb = load_workbook(archivo_filtrado)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            if cell.value is not None and isinstance(cell.value, (int, float)):
                if rango_columna1_min <= cell.value <= rango_columna1_max or rango_columna2_min <= cell.value <= rango_columna2_max:
                    cell.font = Font(bold=True, underline='single')  # Resaltar en negrita y subrayado

    # Guardar el DataFrame seleccionado en un nuevo archivo de Excel
    df_filtrado.to_excel(archivo_filtrado, index=False)
     
    # Mostrar los datos con formato de fecha y hora
    for index, row in df.iterrows():
        fecha = row['FECHA'] if 'FECHA' in df.columns else ''
        hora = row['TIMESTAMP'] if 'TIMESTAMP' in df.columns else ''
        fecha_formateada = fecha.strftime("%Y-%m-%d") if isinstance(fecha, datetime) else ''
        hora_formateada = hora.strftime("%H:%M:%S") if isinstance(hora, datetime) else ''
        print(f"Fecha: {fecha_formateada}, Hora: {hora_formateada}")



    # Guardar el DataFrame seleccionado en un nuevo archivo de Excel
    ###df_filtrado.to_excel('print(df.columns).xlsx', index=False)
    archivo_filtrado = f'archivo_filtrado.xlsx'
    df.to_excel(archivo_filtrado, index=False)
    
    
    # Mostrar mensaje de éxito
    print("El archivo se ha guardado exitosamente.")

    # Abrir el archivo Excel guardado automáticamente
    ruta_archivo = os.path.abspath(archivo_filtrado)
    os.system(f'start {ruta_archivo}')

    # Mostrar un mensaje de éxito
    tk.messagebox.showinfo('Procesamiento completado', 'Se han seleccionado y guardado los datos correctamente.')

# Crear la ventana principal
ventana = tk.Tk()
ventana.title('Selección de datos de Excel')
ventana.geometry('400x250')

# Etiqueta y campo de entrada para la ruta del archivo
label_ruta_archivo = tk.Label(ventana, text='Archivo de Excel:')
label_ruta_archivo.pack()
entry_ruta_archivo = tk.Entry(ventana, width=40)
entry_ruta_archivo.pack()
button_cargar_archivo = tk.Button(ventana, text='Cargar', command=cargar_archivo)
button_cargar_archivo.pack()

# Etiquetas y campos de entrada para la selección de datos
label_columnas = tk.Label(ventana, text='Columnas seleccionadas (separadas por coma):')
label_columnas.pack()
entry_columnas = tk.Entry(ventana, width=40)
entry_columnas.pack()

label_fecha_inicio = tk.Label(ventana, text='Fecha de inicio:')
label_fecha_inicio.pack()
entry_fecha_inicio = tk.Entry(ventana, width=40)
entry_fecha_inicio.pack()

label_fecha_fin = tk.Label(ventana, text='Fecha de fin:')
label_fecha_fin.pack()
entry_fecha_fin = tk.Entry(ventana, width=40)
entry_fecha_fin.pack()

# Botón para procesar los datos
button_procesar = tk.Button(ventana, text='Procesar', command=procesar_datos)
button_procesar.pack()

# Iniciar el bucle principal de la ventana
ventana.mainloop()