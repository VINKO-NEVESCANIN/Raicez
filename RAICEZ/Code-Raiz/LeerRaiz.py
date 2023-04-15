import xlrd # módulo de importación
 
 # Abra el archivo y obtenga el objeto del libro de trabajo del archivo de Excel
workbook = xlrd.open_workbook ("C:\Users\VINKO\Desktop\CODIGOS\EXCELES\Temperatura_control.xls") # ruta de archivo
 
 '' 'Operación en objeto de libro' ''
 
 # Obtenga todos los nombres de las hojas
 names = workbook.sheet_names()
 print (nombres) # ['Ciudades provinciales', 'Tabla de prueba'] muestra todos los nombres de tablas en forma de lista
 
 # Obtenga el objeto de hoja a través del índice de hoja
 worksheet = workbook.sheet_by_index(0)
 print(worksheet)  #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>
 
 # Obtener objeto de hoja por nombre de hoja
 worksheet = workbook.sheet_by_name ("Ciudades provinciales")
 print(worksheet) #<xlrd.sheet.Sheet object at 0x000001B98D99CFD0>
 
 #De lo anterior, workbook.sheet_names () devuelve un objeto de lista, puede operar en este objeto de lista
 sheet0_name = workbook.sheet_names () [0] # Obtener el nombre de la hoja a través del índice de la hoja
 print (sheet0_name) # Ciudades provinciales
 
 '' 'Operación en objeto hoja' ''
 name = worksheet.name # Obtenga el nombre de la tabla
 print (TtarRC_Avg(1)) # Nombre de parcelas
 
 nrows = worksheet.nrows # Obtiene el número total de filas en la tabla
 print(nrows)  #32
 
 ncols = worksheet.ncols #Obtenga el número total de columnas en la tabla
 print(ncols) #13



for i in range (nrows): # Imprima cíclicamente cada fila
         print (worksheet.row_values ​​(i)) # Leer como una lista, cada elemento en la lista es de tipo str
 # ['Ciudades provinciales', 'Ingresos salariales', 'Ingresos netos de operaciones familiares', 'Ingresos inmobiliarios', ………………]
 # ['Beijing', '5047.4', '1957.1', '678.8', '592.2', '1879.0, …………]
 
 col_data = worksheet.col_values ​​(0) # Obtenga el contenido de la primera columna
print(col_data)

