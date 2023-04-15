import pandas as pd

impunt_cols = [42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59]

df = pd.read_excel(r"C:\Users\VINKO\Desktop\CODIGOS\EXCELES\Temperatura_control.xls",
sheet_name="Temp control 15feb 13Abr 2020", header = 0, usecols= impunt_cols)

print(df.shape)

df = df[df["TtarRC_Avg(1)"]==""]
df_cols = df.columns

for col in df_cols:
    print(df[col].head(5))

print(df["TtarRC_Avg(1)"].head(5))
