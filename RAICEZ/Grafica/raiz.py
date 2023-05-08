import pandas as pd
import matplotlib.pyplot as plt

workbook1 =r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'

df = pd.read_excel(workbook1)

print(df.head())

valores = df[["TtarRC_Avg(1)", "TtarRC_Avg(2)","TtarRC_Avg(3)", "TtarRC_Avg(4)", "TtarRC_Avg(5)", "TtarRC_Avg(6)", "TtarRC_Avg(7)", "TtarRC_Avg(8)", "TtarHC_Avg(1)", "TtarHC_Avg(2)", "TtarHC_Avg(3)", "TtarHC_Avg(4)", "TtarHC_Avg(5)", "TtarHC_Avg(16)", "TtarHC_Avg(7)", "TtarHC_Avg(8)"]]

print(valores)

#ax = valores.plot.bar(x="Date", y="TtarHC_Avg(1)",rot = 0)

#plt.show()
 