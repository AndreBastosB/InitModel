import pandas as pd

csvToXlsx1 = r'C:\Users\Take4\Desktop\1651676165304.csv'
csvToXlsx2 = r'C:\Users\Take4\Desktop\1651676165304.xlsx'
transform = pd.read_csv(csvToXlsx1, sep=";", encoding = 'latin-1')
transform.to_excel(csvToXlsx2, index=None)