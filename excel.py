#Parte do programa responsavel por copiar parte do arquivo original para outro excel que sera enviado por email

import pandas as pd
from pandas.core.indexes.period import PeriodIndex


#Le o arquivo
arquivo = pd.read_excel(r"exemplo.xlsx")

#Coleta as linhas que serão separadas
linhaI = int(input("digite a primeira linha: "))
linhaF = int(input("digite a ultima linha: "))


#Define as variaveis
a = []
e = linhaI
n = 0

#Separa as linhas em lista
while e <= linhaF:
    
    a.append(e)
    n = n+1
    e = e+1


#Define "selecionadas" como as linhas escolhidas
selecionadas = arquivo.loc[a]


#Copia as linhas escolhidas em outro arquivo
selecionadas.to_excel('exemplo2.xlsx')

