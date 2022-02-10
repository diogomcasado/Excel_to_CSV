from tqdm import tqdm
import pandas as pd
import os
import fnmatch
import shutil
from os import path, system

system('cls')

folder = os.path.dirname(os.path.abspath(__file__))

print('Pasta: ', folder)


# funcao para verificar/criar pasta e mover .xlsx
def checkCreateFolder():
    newFolder = ('xlsx')
    pastaCriar = os.path.join(folder, newFolder)

    if not os.path.isdir(pastaCriar):
        os.mkdir(pastaCriar)

    for x in resutados_excel:
        if x.endswith('.xls') or x.endswith('.xlsx'):
            path = findFile(x, folder)
            shutil.move(path, pastaCriar)


# funcao para ter lista de resultados
def find(pattern, path):
    result = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                result.append(name)
    return result


# funcao para ter caminho do ficheiro
def findFile(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)

resultados_XLSX = (find('*.xlsx', folder))
resultados_XLS = (find('*.xls', folder))

resutados_excel = resultados_XLS + resultados_XLSX


# verifica se há ficheiros .xlsx
if (len(resutados_excel) == 0):
    print('Erro: Ficheiros .XLSX/.XLS não encontrados')
    exit()
else:
    if not len(resultados_XLSX) == 0:
        print('Ficheiros .XLSX encontrados:', len(resultados_XLSX))
    if not len(resultados_XLS) == 0:
        print('Ficheiros .XLS encontrados:', len(resultados_XLS))


# transforma xlsx e xls para csv
print("\nA tranformar ficheiros .XLSX/.XLS em .CSV...")
for x in tqdm(resutados_excel):

    if x.endswith('.xlsx'):
        nome_final = x.replace(".xlsx", ".csv")
    elif x.endswith('.xls'):
        nome_final = x.replace(".xls", ".csv")
    else:
        continue

    xlsxToCsvOk = True
    path = findFile(x, folder)

    data_xls = pd.read_excel(path, dtype=str, index_col=None)
    data_xls.to_csv(os.path.join(os.path.dirname(path), nome_final),
                    encoding='utf-8', index=False, sep=';')

print("\nFicheiros transformados")


# pergunta ao utilizador
resposta = ''
while resposta == '':
    resposta = input("Mover .XLSX/.XLS para uma pasta?(1-sim 2-nao)")
    if resposta == '1':
        print("A mover")
        checkCreateFolder()
    elif resposta == '2':
        exit()
    elif resposta != '1':
        resposta = ''
