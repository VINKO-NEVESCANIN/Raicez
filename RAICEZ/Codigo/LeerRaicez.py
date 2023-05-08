import pandas as pd

def main():
    df = leer_archivos()
    df = agregar_filtros(df)
    
def leer_archivos():
    print("Leyendo Archivo")
    import os
    impunt_cols = [42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57]
    
    path = "..\Input"
    filename = inpunt("Ingrese el nombre del archivo") + "xlsx"\
    fullpath = os.path.join(path, filename)

    df = pd.read_excel(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx',
                       sheet_name="Temp control 15feb 13Abr 2020", header = 0, usecols= impunt_cols)


    return df



def agregar_filtros(df):
    print("Agregando Filtros")

df = df[df["TtarRC_Avg(1)"]==""]

df_cols = df.columns

for col in df_cols:
    print(df[col].head(5))

#print(df["TtarRC_Avg(1)"].head(5))

#print(df["TIMESTAMP"].head(5))

df.to_csv("Temperaturas_2022.xlsx",
          sep= ",",
          header= True,
          index=False)
