import pandas as pd
import os


def main():
    df = leer_archivos()
    df = agregar_filtros(df)
    
    visualizar_datos(df)
    exportar_datos(df)
    
def leer_archivos():
    #print("Leyendo archivo")

    input_cols = [42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59]
    # impunt_cols = [42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57]
    
    path = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'
    filename = input ("") + "xlsx"
    fullpath = os.path.join(path, filename)

    df = pd.read_excel(
    r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', sheet_name="Temp control 15feb 13Abr 2020", header = 0, usecols = input_cols)

    return df


def agregar_filtros(df):
    print("Agregando filtros")
    
    df = df[df["TtarRC_Avg(1)"]== 12]
    
    return df

#print(df.shape)
#df = df[df["TtarRC_Avg(1)"]== 12]
def visualizar_datos(df):
    print("Visualizando los primeros 5 registros")
    df_cols = df.columns

    for col in df_cols:
        print(df[col].head(5))
        
        
def exportar_datos(df):
    print("Exportando archivo procesado")
    #print(df["TtarRC_Avg(1)"].head(5))

    df.to_csv(r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx', sep = ",", header = True, index = False)
    
      
if __name__ == "__main__":
    main()
    input("\tPROCESO FINALIZADO. Presionar enter para salir")    

