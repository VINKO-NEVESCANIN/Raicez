import panda as pd
import matplotlib.pyplot as plt

workbook1 = "C:\ProyectosGIT\Raicez/RAICEZ/Excel/Temperatura control 15 feb- 13 abr 2020_OFICIAL.xlsx"

df = pd.read_excel(workbook1)

print (df.head())

#valores = df[["TtarRC_Avg(1)", "TtarHC_Avg(1)"]]
#print(valores)
