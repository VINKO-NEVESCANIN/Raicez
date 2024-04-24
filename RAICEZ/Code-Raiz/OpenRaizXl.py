import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.simpledialog import askstring

import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    if ruta_archivo:
        entry_ruta_archivo.delete(0, tk.END)
        entry_ruta_archivo.insert(tk.END, ruta_archivo)

def procesar_datos():
    archivo = entry_ruta_archivo.get()
    if archivo:
        columnas_seleccionadas = entry_columnas.get().split(',')
        if columnas_seleccionadas:
            mismo_valor = messagebox.askyesno("Mismo valor", "¿Desea ingresar los mismos valores para todos los mínimos y máximos?")
            if mismo_valor:
                rango_min = float(askstring("Input", "Introduce el valor mínimo para todas las columnas:"))
                rango_max = float(askstring("Input", "Introduce el valor máximo para todas las columnas:"))
                filtrar_datos(archivo, columnas_seleccionadas, rango_min, rango_max)
            else:
                filtrar_datos(archivo, columnas_seleccionadas)
        else:
            messagebox.showinfo("Información", "Por favor, introduce las columnas.")
    else:
        messagebox.showinfo("Información", "Por favor, carga un archivo.")

def filtrar_datos(archivo, columnas_seleccionadas, rango_min=None, rango_max=None):
    df = pd.read_excel(archivo)
    condiciones = {}

    for columna in columnas_seleccionadas:
        if rango_min is not None and rango_max is not None:
            condiciones[columna] = (rango_min, rango_max)
        else:
            rango_min = float(askstring("Input", f"Introduce el mínimo para {columna}:"))
            rango_max = float(askstring("Input", f"Introduce el máximo para {columna}:"))
            condiciones[columna] = (rango_min, rango_max)
        
        # Convertir la columna a números flotantes
        df[columna] = pd.to_numeric(df[columna], errors='coerce')

    nombre_archivo = os.path.splitext(os.path.basename(archivo))[0]
    archivo_salida = f'{nombre_archivo}_datos_filtrados.xlsx'
    df.to_excel(archivo_salida, index=False)

    wb = load_workbook(archivo_salida)
    ws = wb.active
    fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for col_idx, column in enumerate(df.columns, 1):
        if column in columnas_seleccionadas:
            idx = columnas_seleccionadas.index(column)
            for row_idx, valor in enumerate(df[column], start=2):
                if isinstance(valor, (int, float)) and (valor < condiciones[column][0] or valor > condiciones[column][1]):
                    ws.cell(row=row_idx, column=col_idx).fill = fill
                    ws.cell(row=row_idx, column=col_idx).font = Font(bold=True)


    wb.save(archivo_salida)
    print(f"Los datos filtrados se han guardado en '{archivo_salida}'")
    os.startfile(archivo_salida)

ventana = tk.Tk()
ventana.title('Selección de datos de Excel')

label_ruta_archivo = tk.Label(ventana, text='Archivo de Excel:')
label_ruta_archivo.grid(row=0, column=0)
entry_ruta_archivo = tk.Entry(ventana, width=40)
entry_ruta_archivo.grid(row=0, column=1)
button_cargar_archivo = tk.Button(ventana, text='Cargar', command=cargar_archivo)
button_cargar_archivo.grid(row=0, column=2)

label_columnas = tk.Label(ventana, text='Columnas seleccionadas (separadas por coma):')
label_columnas.grid(row=1, column=0)
entry_columnas = tk.Entry(ventana, width=40)
entry_columnas.grid(row=1, column=1, columnspan=2)

button_procesar = tk.Button(ventana, text='Procesar', command=procesar_datos)
button_procesar.grid(row=2, column=0, columnspan=3)

ventana.mainloop()
