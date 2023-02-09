import panda as pd
import matplotlib.pyplot as plt

fig, ax = plt.subplots()

workbook1 = 'C:\ProyectosGIT\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'
df = pd.read_excel(workbook1)

datos = df['TtarRC_Avg(1)','TtarRC_Avg(2)','TtarRC_Avg(3)', 'TtarRC_Avg(4)', 'TtarRC_Avg(5)', 'TtarRC_Avg(6)', 'TtarRC_Avg(7)', 'TtarRC_Avg(8)',
           'TtarHC_Avg(1)', 'TtarHC_Avg(2)', 'TtarHC_Avg(3)', 'TtarHC_Avg(4)', 'TtarHC_Avg(5)', 'TtarHC_Avg(6)', 'TtarHC_Avg(7)', 'TtarHC_Avg(8)']
Temp = []
bar_labels = ['red', 'blue', '_red', 'orange']
bar_colors = ['tab:red', 'tab:blue', 'tab:red', 'tab:orange']

#print (df.head())
print(df.datos)
