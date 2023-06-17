import tkinter as tk
from tkinter import filedialog
import pandas as pd

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
    df_seleccionado = df[df['Fecha'].isin(rango_fechas)][columnas_seleccionadas]

    # Guardar el DataFrame seleccionado en un nuevo archivo de Excel
    df_seleccionado.to_excel('datos_seleccionados.xlsx', index=False)

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