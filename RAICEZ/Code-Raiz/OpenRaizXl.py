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

    def cargar_archivo(self):
        ruta_archivo = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
        if ruta_archivo:
            self.entry_ruta_archivo.delete(0, tk.END)
            self.entry_ruta_archivo.insert(tk.END, ruta_archivo)

    def procesar_datos(self):
        archivo = self.entry_ruta_archivo.get()
        columnas_seleccionadas = self.entry_columnas.get().split(',')
        rango_fechas = pd.date_range(start=self.entry_fecha_inicio.get(), end=self.entry_fecha_fin.get())

        try:
            self.df = pd.read_excel(archivo)
            self.crear_interfaz_grafica()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {str(e)}")

    def crear_interfaz_grafica(self):
        nombres_columnas = self.df.columns.tolist()

        for widget in self.ventana.winfo_children():
            widget.destroy()

        for i, nombre_columna in enumerate(nombres_columnas):
            label = tk.Label(self.ventana, text=f'Rango para {nombre_columna}:')
            label.grid(row=i, column=0)

            caja_min = tk.Entry(self.ventana)
            caja_max = tk.Entry(self.ventana)

            caja_min.grid(row=i, column=1)
            caja_max.grid(row=i, column=2)

            self.cajas_rango_columnas.append((caja_min, caja_max))

        boton_filtrar = tk.Button(self.ventana, text="Filtrar", command=self.filtrar_datos)
        boton_filtrar.grid(row=len(nombres_columnas), column=0, columnspan=2)

    def filtrar_datos(self):
        rangos = []

        for caja_min, caja_max in self.cajas_rango_columnas:
            rango_min_text = caja_min.get()
            rango_max_text = caja_max.get()

            try:
                rango_min = float(rango_min_text) if rango_min_text else float('-inf')
            except ValueError:
                rango_min = float('-inf')

            try:
                rango_max = float(rango_max_text) if rango_max_text else float('inf')
            except ValueError:
                rango_max = float('inf')

            rangos.append((rango_min, rango_max))

        try:
            for columna in self.df.select_dtypes(include='number'):
                for rango_min, rango_max in rangos:
                    self.df = self.df[(self.df[columna] >= rango_min) & (self.df[columna] <= rango_max)]

            if not self.df.empty:
                archivo_filtrado = 'archivo_filtrado.xlsx'
                self.df.to_excel(archivo_filtrado, index=False)

                plt.plot(self.df.iloc[:, 0], self.df.iloc[:, 1])  # Suponiendo que se grafican las primeras dos columnas
                plt.xlabel('Columna1')
                plt.ylabel('Columna2')
                plt.title('Gráfico de datos filtrados')
                plt.show()

                messagebox.showinfo("Procesamiento completado", "Se han seleccionado y guardado los datos correctamente.")
            else:
                messagebox.showwarning("Advertencia", "No se encontraron datos que coincidan con los criterios de filtrado.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo filtrar los datos: {str(e)}")

# Crear la ventana principal
ventana = tk.Tk()
ventana.title('Selección de datos de Excel')
ventana.geometry('400x250')

# Crear una instancia de la clase DataProcessor
procesador_datos = DataProcessor(ventana)

# Etiqueta y campo de entrada para la ruta del archivo
label_ruta_archivo = tk.Label(ventana, text='Archivo de Excel:')
label_ruta_archivo.pack()
procesador_datos.entry_ruta_archivo = tk.Entry(ventana, width=40)
procesador_datos.entry_ruta_archivo.pack()
button_cargar_archivo = tk.Button(ventana, text='Cargar', command=procesador_datos.cargar_archivo)
button_cargar_archivo.pack()

# Etiquetas y campos de entrada para la selección de datos
label_columnas = tk.Label(ventana, text='Columnas seleccionadas (separadas por coma):')
label_columnas.pack()
procesador_datos.entry_columnas = tk.Entry(ventana, width=40)
procesador_datos.entry_columnas.pack()

label_fecha_inicio = tk.Label(ventana, text='Fecha de inicio:')
label_fecha_inicio.pack()
procesador_datos.entry_fecha_inicio = tk.Entry(ventana, width=40)
procesador_datos.entry_fecha_inicio.pack()

label_fecha_fin = tk.Label(ventana, text='Fecha de fin:')
label_fecha_fin.pack()
procesador_datos.entry_fecha_fin = tk.Entry(ventana, width=40)
procesador_datos.entry_fecha_fin.pack()

# Botón para procesar los datos
button_procesar = tk.Button(ventana, text='Procesar', command=procesador_datos.procesar_datos)
button_procesar.pack()

# Iniciar el bucle principal de la ventana
ventana.mainloop()