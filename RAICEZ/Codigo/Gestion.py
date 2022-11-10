from openpyxl import Workbook
from openpyxl.styles import Font

book = Workbook()
libro = Workbook()
sheet = book.active
hoja = libro.active

sheet["A1"] = 3.3456446
sheet["A2"] = 5.676767

book.save("TemperaturasMachigai.xlsx")
