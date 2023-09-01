import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import Font
import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
from tkinter import *

def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    if ruta_archivo:
        entry_ruta_archivo.delete(0, tk.END)
        entry_ruta_archivo.insert(tk.END, ruta_archivo)


def procesar_datos():
    archivo = entry_ruta_archivo.get()
    columnas_seleccionadas = entry_columnas.get().split(',')
    rango_fechas = pd.date_range(start=entry_fecha_inicio.get(), end=entry_fecha_fin.get())

    # Leer el archivo de Excel original
    df = pd.read_excel(archivo)

    # Filtrar los datos por las columnas seleccionadas y el rango de fechas
    #df_seleccionado = df[df['Fecha'].isin(rango_fechas)][columnas_seleccionadas]
    
    # Acceder a la columna de fechas por su posición (en este ejemplo, la primera columna)
    df_filtrado = df.iloc[:, 0]  # Utiliza el índice 0 para la primera columna
    
    # Convertir las columnas al formato datetime
    for col in columnas_seleccionadas:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    # Crear la interfaz gráfica
    root = Tk()
    root.title("Aplicación")
    
    
    
    # Definir cajas de entrada para los rangos de las columnas
    caja_rango_columna1_min = Entry(root)
    caja_rango_columna1_max = Entry(root)
    caja_rango_columna2_min = Entry(root)
    caja_rango_columna2_max = Entry(root)
    caja_rango_columna3_min = Entry(root)
    caja_rango_columna3_max = Entry(root)
    caja_rango_columna4_min = Entry(root)
    caja_rango_columna4_max = Entry(root)
    
    # Obtener los parámetros ingresados desde la interfaz gráfica
    rango_columna1_min_text = caja_rango_columna1_min.get()
    rango_columna1_max_text = caja_rango_columna1_max.get()
    rango_columna2_min_text = caja_rango_columna2_min.get()
    rango_columna2_max_text = caja_rango_columna2_max.get()
    rango_columna3_min_text = caja_rango_columna3_min.get()
    rango_columna3_max_text = caja_rango_columna3_max.get()
    rango_columna4_min_text = caja_rango_columna4_min.get()
    rango_columna4_max_text = caja_rango_columna4_max.get()

    # Validar y convertir los valores de las cajas de entrada
    try:
        rango_columna1_min = float(rango_columna1_min_text) if rango_columna1_min_text else 0.0
    except ValueError:
        rango_columna1_min = 0.0

    try:
        rango_columna1_max = float(rango_columna1_max_text) if rango_columna1_max_text else float('inf')
    except ValueError:
        rango_columna1_max = float('inf')
        
    try:
        rango_columna2_min = float(rango_columna2_min_text) if rango_columna1_min_text else 0.0
    except ValueError:
        rango_columna2_min = 0.0

    try:
        rango_columna2_max = float(rango_columna2_max_text) if rango_columna1_max_text else float('inf')
    except ValueError:
        rango_columna2_max = float('inf')
        
    try:
        rango_columna3_min = float(rango_columna3_min_text) if rango_columna1_min_text else 0.0
    except ValueError:
        rango_columna3_min = 0.0

    try:
        rango_columna3_max = float(rango_columna3_max_text) if rango_columna1_max_text else float('inf')
    except ValueError:
        rango_columna3_max = float('inf')
        
    try:
        rango_columna4_min = float(rango_columna4_min_text) if rango_columna1_min_text else 0.0
    except ValueError:
        rango_columna4_min = 0.0

    try:
        rango_columna4_max = float(rango_columna4_max_text) if rango_columna1_max_text else float('inf')
    except ValueError:
        rango_columna4_max = float('inf')  
        
    # Filtrar los datos según los rangos de valores
    columnas_avg = [f'TtarRC_Avg({i})' for i in range(1, 9)]
    df_filtrado = df[
        (df[columnas_avg].astype(float) >= rango_columna1_min) & (df[columnas_avg].astype(float) <= rango_columna1_max) &
        (df[columnas_avg].astype(float) >= rango_columna2_min) & (df[columnas_avg].astype(float) <= rango_columna2_max) &
        (df[columnas_avg].astype(float) >= rango_columna3_min) & (df[columnas_avg].astype(float) <= rango_columna3_max) &
        (df[columnas_avg].astype(float) >= rango_columna4_min) & (df[columnas_avg].astype(float) <= rango_columna4_max)
        # Repite el mismo patrón para las otras columnas
        ]
        
            

    # Guardar los datos filtrados en un nuevo archivo de Excel
    archivo_filtrado = 'archivo_filtrado.xlsx'
    df_filtrado.to_excel(archivo_filtrado, index=False)
    
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