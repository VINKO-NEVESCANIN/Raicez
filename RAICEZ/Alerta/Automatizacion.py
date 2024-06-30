import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import matplotlib.pyplot as plt
import schedule
import time
from threading import Thread

epsilon = 1.001  # Valor pequeño para ajustar los límites

# Variables globales para la automatización
archivo_global = None
columnas_seleccionadas_global = None
valor_minimo_global = None
valor_maximo_global = None

def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    if ruta_archivo:
        entry_ruta_archivo.delete(0, tk.END)
        entry_ruta_archivo.insert(tk.END, ruta_archivo)

def procesar_datos():
    global archivo_global, columnas_seleccionadas_global, valor_minimo_global, valor_maximo_global
    
    archivo = entry_ruta_archivo.get()
    if archivo:
        columnas_seleccionadas = [col.strip() for col in entry_columnas.get().split(',')]
        if columnas_seleccionadas:
            respuesta = messagebox.askyesno("Input", "¿Desea ingresar los mismos valores para todas las columnas?")
            if respuesta:  # Si el usuario selecciona "Sí"
                try:
                    valor_minimo = float(simpledialog.askstring("Input", "Ingrese el valor mínimo para todas las columnas:"))
                    valor_maximo = float(simpledialog.askstring("Input", "Ingrese el valor máximo para todas las columnas:"))
                    archivo_global = archivo
                    columnas_seleccionadas_global = columnas_seleccionadas
                    valor_minimo_global = valor_minimo
                    valor_maximo_global = valor_maximo
                    filtrar_datos(archivo, columnas_seleccionadas, valor_minimo, valor_maximo)
                except (TypeError, ValueError):
                    messagebox.showerror("Error", "Valor mínimo o máximo no válido. Operación cancelada.")
            else:  # Si el usuario selecciona "No"
                try:
                    archivo_global = archivo
                    columnas_seleccionadas_global = columnas_seleccionadas
                    filtrar_datos(archivo, columnas_seleccionadas)
                except (TypeError, ValueError):
                    messagebox.showerror("Error", "Valores no válidos. Operación cancelada.")
        else:
            messagebox.showinfo("Información", "Por favor, introduce las columnas.")
    else:
        messagebox.showinfo("Información", "Por favor, carga un archivo.")

def filtrar_datos(archivo, columnas_seleccionadas, valor_minimo=None, valor_maximo=None):
    df = pd.read_excel(archivo)
    conditions = []
    columnas_no_encontradas = []
    datos_fuera_de_rango = False

    if valor_minimo is not None and valor_maximo is not None:
        for columna in columnas_seleccionadas:
            if columna in df.columns:
                df[columna] = pd.to_numeric(df[columna], errors='coerce')
                condition = (df[columna] < valor_minimo) | (df[columna] > valor_maximo + epsilon)
                conditions.append((columna, condition))
                if df[columna][condition].count() > 0:
                    datos_fuera_de_rango = True
                print(f"Límites para la columna {columna}: Mínimo = {valor_minimo}, Máximo = {valor_maximo}")
                print(f"Valores fuera de rango para la columna {columna}:")
                print(df[columna][condition])
            else:
                columnas_no_encontradas.append(columna)
    else:
        for columna in columnas_seleccionadas:
            if columna in df.columns:
                try:
                    rango_min = float(simpledialog.askstring("Input", f"Introduce el mínimo para {columna}:"))
                    rango_max = float(simpledialog.askstring("Input", f"Introduce el máximo para {columna}:"))
                    df[columna] = pd.to_numeric(df[columna], errors='coerce')
                    condition = (df[columna] < rango_min) | (df[columna] > rango_max + epsilon)
                    conditions.append((columna, condition))
                    if df[columna][condition].count() > 0:
                        datos_fuera_de_rango = True
                    print(f"Límites para la columna {columna}: Mínimo = {rango_min}, Máximo = {rango_max}")
                    print(f"Valores fuera de rango para la columna {columna}:")
                    print(df[columna][condition])
                except (TypeError, ValueError):
                    messagebox.showerror("Error", f"Valores no válidos para la columna {columna}. Operación cancelada.")
                    return
            else:
                columnas_no_encontradas.append(columna)

    if columnas_no_encontradas:
        messagebox.showinfo("Información", f"Columnas no encontradas en el archivo: {', '.join(columnas_no_encontradas)}")

    if not datos_fuera_de_rango:
        messagebox.showinfo("Información", "No hay datos fuera de rango para las columnas seleccionadas.")
        return

    nombre_archivo = os.path.splitext(os.path.basename(archivo))[0]
    archivo_salida = f'{nombre_archivo}_datos_filtrados.xlsx'
    df.to_excel(archivo_salida, index=False)

    wb = load_workbook(archivo_salida)
    ws = wb.active
    fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for col_idx, column in enumerate(df.columns, 1):
        if column in columnas_seleccionadas:
            idx = columnas_seleccionadas.index(column)
            for row_idx, valor in enumerate(df[column], 2):
                if conditions[idx][1][row_idx - 2]:
                    ws.cell(row=row_idx, column=col_idx).fill = fill

    wb.save(archivo_salida)
    print(f"Los datos filtrados se han guardado en '{archivo_salida}'")
    os.startfile(archivo_salida)

    # Generar gráfico para las columnas seleccionadas
    generar_grafico(df, columnas_seleccionadas, conditions)

    print("Columnas disponibles en el DataFrame:", df.columns)

def generar_grafico(df, columnas_seleccionadas, conditions):
    fig, ax = plt.subplots(figsize=(10, 6))
    
    porcentajes_error = []
    for idx, columna in enumerate(columnas_seleccionadas):
        if columna in df.columns:
            condition = conditions[idx][1]
            porcentaje = df[condition].shape[0] / df[columna].shape[0] * 100
            porcentajes_error.append(porcentaje)
        else:
            print(f"Columna {columna} no encontrada en el DataFrame.")
            porcentajes_error.append(0)
    
    ax.bar(columnas_seleccionadas, porcentajes_error, color='skyblue')
    ax.set_ylabel('Porcentaje de Error')
    ax.set_title('Porcentaje de Error para las Columnas Seleccionadas')
    ax.set_xlabel('Columnas')
    ax.set_yticks(range(0, 101, 10))  # Ajustar el eje y para mostrar más valores de porcentaje

    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

def tarea_automatica():
    if archivo_global and columnas_seleccionadas_global:
        try:
            if valor_minimo_global is not None and valor_maximo_global is not None:
                filtrar_datos(archivo_global, columnas_seleccionadas_global, valor_minimo_global, valor_maximo_global)
            else:
                filtrar_datos(archivo_global, columnas_seleccionadas_global)
        except Exception as e:
            print(f"Error en la tarea automática: {e}")

def iniciar_automatizacion():
    # Programar la tarea para que se ejecute cada día a una hora específica (por ejemplo, a las 8:00 AM)
    schedule.every().day.at("08:00").do(tarea_automatica)
    
    def run_scheduler():
        while True:
            schedule.run_pending()
            time.sleep(1)
    
    # Ejecutar el programador en un hilo separado
    t = Thread(target=run_scheduler)
    t.daemon = True
    t.start()

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

button_automatizar = tk.Button(ventana, text='Iniciar Automatización', command=iniciar_automatizacion)
button_automatizar.grid(row=3, column=0, columnspan=3)

ventana.mainloop()
