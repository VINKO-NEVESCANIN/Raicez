import pandas as pd


file_name = r'C:\Users\VINKO\Documents\GitHub\Raicez\RAICEZ\Excel\Temperatura_control.xlsx'

if file_name.endswith('.xlsx'):
    df = pd.read_excel(
        file_name,
        engine='openpyxl'
    )
    print(df)
elif file_name.endswith('.xls'):
    df = pd.read_excel(
        file_name,
        engine='xlrd'
    )
    print(df)
elif file_name.endswith('.csv'):
    df = pd.read_csv(file_name)
    print(df)