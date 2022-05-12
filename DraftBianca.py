import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os 


# def transformCSVtoXLSX():
csvToXlsx1 = r'C:\Users\Take4\Desktop\1651676165304.csv'
csvToXlsx2 = r'C:\Users\Take4\Desktop\1651676165304.xlsx'

transform = pd.read_csv(csvToXlsx1, sep="\t", encoding = 'latin-1')
testandoPush = 'push é empurrar'
transform.to_excel(csvToXlsx2, index=None)

# transformCSVtoXLSX()

# book = load_workbook (csvToXlsx2)
# ws = book.worksheets[0]
# for cell in ws["A"]:
#     if cell.value is None:
#         print (cell.row)
#         break
# else:
#     print (cell.row + 1)

#INDICE -- INDICE -- INDICE -- INDICE -- INDICE -- INDICE -- INDICE -- INDICE -- INDICE
AmericanasEstudioShoptime = 0
AmericanasEscritorio = 1
aMercadoEscritorio = 2
AmeEscritorio = 3
LetsCDs = 4
LetsSAC = 5
AmericanasLoja = 6
LocalLoja = 7
AmericanasOutros = 8
LetsHUB = 9
LetsTransporte = 10
LestBase = 11
HortifrutiLoja = 12
HortifrutiEscritorio = 13
GrupoUnicoEscritorio = 14
VemEscritorio = 15
IF = 16
aDelivery = 17
BRMania = 18
Puket = 19
imaginarium = 20
Mind = 21
Lovebrands = 22
bitCapital = 23
Parati = 24
skoob = 25
AmeOutros = 26
AmeSAC = 27

listaDeProtocolos = [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],]

linha = 2

data1 = load_workbook (csvToXlsx2)
data2 = data1 ['Sheet1'] 
valor = data2.cell(row=linha, column=1).value

# while valor is not None:
while linha <= 739:

    valor = data2.cell(row=linha, column=1).value
    protocolo1 = valor.split(';', 1)[1]
    protocolo = protocolo1.split(';', 1)[0]
    # print (protocolo)
    if "Americanas - Loja" in valor:
        listaDeProtocolos[0].append(protocolo)
    elif "Let's - CDs" in valor:
        listaDeProtocolos[1].append(protocolo)
    else:
        pass
    linha += 1
    
print ((listaDeProtocolos[0]))
print (listaDeProtocolos[1])











