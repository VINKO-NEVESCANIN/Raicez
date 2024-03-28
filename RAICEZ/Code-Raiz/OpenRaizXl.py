import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime

class DataProcessor:
    def __init__(self, ventana):
        self.ventana = ventana
        self.df = None
        self.cajas_rango_columnas = []
def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    if ruta_archivo:
        entry_ruta_archivo.delete(0, tk.END)
        entry_ruta_archivo.insert(tk.END, ruta_archivo)

def procesar_datos():
    global cajas_rango_columnas
    archivo = entry_ruta_archivo.get()
    columnas_seleccionadas = entry_columnas.get().split(',')
    rangos_columnas = []

    # Crear cajas de entrada y etiquetas para cada columna en dos columnas
    # Dentro de la función procesar_datos(), donde se crean las cajas de entrada para los rangos de columnas
    for i, nombre_columna in enumerate(columnas_seleccionadas):
        label = tk.Label(ventana, text=f'Rango para {nombre_columna}:')
        label.grid(row=i * 2, column=0, columnspan=2)

        caja_min = tk.Entry(ventana)
        caja_max = tk.Entry(ventana)

        caja_min.grid(row=i * 2 + 1, column=0)
        caja_max.grid(row=i * 2 + 1, column=1)

        cajas_rango_columnas.append((caja_min, caja_max))

    # Botón para realizar el filtrado
    boton_filtrar = tk.Button(ventana, text="Filtrar", command=filtrar_datos)
    boton_filtrar.grid(row=len(columnas_seleccionadas) * 2, column=0, columnspan=2)


def filtrar_datos(archivo, columnas_seleccionadas, rangos_columnas):
    df = pd.read_excel(archivo)
    for columna, (caja_min, caja_max) in zip(columnas_seleccionadas, rangos_columnas):
        rango_min = float(caja_min.get() or 0)
        rango_max = float(caja_max.get() or float('inf'))
        df = df[(df[columna] >= rango_min) & (df[columna] <= rango_max)]

    # Procesar el DataFrame filtrado...
    print(df)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title('Selección de datos de Excel')

# Etiqueta y campo de entrada para la ruta del archivo
label_ruta_archivo = tk.Label(ventana, text='Archivo de Excel:')
label_ruta_archivo.grid(row=0, column=0)
entry_ruta_archivo = tk.Entry(ventana, width=40)
entry_ruta_archivo.grid(row=0, column=1)
button_cargar_archivo = tk.Button(ventana, text='Cargar', command=cargar_archivo)
button_cargar_archivo.grid(row=0, column=2)

# Etiquetas y campos de entrada para la selección de datos
label_columnas = tk.Label(ventana, text='Columnas seleccionadas (separadas por coma):')
label_columnas.grid(row=1, column=0)
entry_columnas = tk.Entry(ventana, width=40)
entry_columnas.grid(row=1, column=1, columnspan=2)

# Botón para procesar los datos
button_procesar = tk.Button(ventana, text='Procesar', command=procesar_datos)
button_procesar.grid(row=2, column=0, columnspan=3)

# Iniciar el bucle principal de la ventana
ventana.mainloop()