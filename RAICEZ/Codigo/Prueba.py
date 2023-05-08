import pandas as pd
import matplotlib.pyplot as plt

workbook1 = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'
df = pd.read_excel(workbook1)
 
print(df.head())
valores = df[['TIMESTAMP','RECORD','DateTime(1)','DateTime(9)','DateTime(4)','DateTime(5)','Target','TtarRC_Avg(1)','TtarRC_Avg(2)','TtarRC_Avg(3)','TtarRC_Avg(4)','TtarRC_Avg(5)','TtarRC_Avg(6)','TtarRC_Avg(7)','TtarRC_Avg(8)','TtarHC_Avg(1)','TtarHC_Avg(2)','TtarHC_Avg(3)','TtarHC_Avg(4)','TtarHC_Avg(5)','TtarHC_Avg(6)','TtarHC_Avg(7)','TtarHC_Avg(8)']]
 
print(valores)
ax = valores.plot.bar(x='DateTime(1)', y='TtarRC_Avg(1)', rot=0)

plt.title("Porcentaje de Riego")
plt.xlabel('DateTime(1)')
plt.ylabel('TtarRC_Avg(1)')
plt.xticks(x='DateTime(1)', rotation=40)
plt.show()

