#pip install pandas
#pip install tqdm
#pip install xlrd==1.2.0
#pip install openpyxl

from tqdm import tqdm
import pandas as pd
import os, fnmatch
import shutil
from os import path, system

system('cls')

folder = os.path.dirname(os.path.abspath(__file__))

print('Pasta: ', folder)

#funcao para verificar/criar pasta e mover .xlsx
def checkCreateFolder():
    newFolder = ('xlsx')
    pastaCriar = os.path.join(folder, newFolder)

    if not os.path.isdir(pastaCriar):
        os.mkdir(pastaCriar)
    
    for x in tqdm(resultadosXLSX):
        if x.endswith ('.xlsx'):
            path = findFile(x, folder)
            #os.path.dirname(path)
            shutil.move(path,pastaCriar)

#funcao para ter lista de resultados
def find(pattern, path):
    result = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                result.append(name)
    return result

#funcao para ter caminho do ficheiro
def findFile(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)

resultadosXLSX = (find('*.xlsx', folder))

#verifica se há ficheiros .xlsx
if (len(resultadosXLSX) == 0):
    print('Erro: Ficheiros .XLSX não encontrados')
    exit()
else:
    print('Ficheiros .XLSX encontrados:' ,len(resultadosXLSX))


#xlsx to csv
print("\nA tranformar ficheiros .XLSX em .CSV...")
for x in tqdm(resultadosXLSX):
    if x.endswith ('.xlsx'):
        xlsxToCsvOk = True
        #print(x)
        nomeFinal = x.replace(".xlsx", ".csv")

        path = findFile(x, folder)

        data_xls = pd.read_excel(path, dtype=str, index_col=None)
        data_xls.to_csv(os.path.join(os.path.dirname(path), nomeFinal), encoding='utf-8', index=False, sep=';')   
        

print("\nFicheiros .XLXS transformados")
resposta = ''
while resposta == '':
    resposta = input ("Mover .XLSX para uma pasta?(1-sim 2-nao)")
    if resposta == '1':
        print("a mover")
        checkCreateFolder()
    elif resposta == '2':
        exit()
    elif resposta != '1':
        resposta = ''






