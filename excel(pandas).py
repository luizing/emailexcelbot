import pandas

#Le o arquivo
arquivo = pandas.read_excel(r"arquivo_original.xlsx")

#Coleta as linhas que serão separadas
linhaI = int(input("digite a primeira linha: "))
linhaF = int(input("digite a ultima linha: "))


#Define as variaveis
linhasCopiar = []
e = linhaI

#Separa as linhas em lista
while e <= linhaF :
    
    linhasCopiar.append(e)
    e = e+1


#Define "selecionadas" como as linhas escolhidas
selecionadas = arquivo.loc[linhasCopiar]


#Copia as linhas escolhidas em outro arquivo
selecionadas.to_excel('arquivo_enviado.xlsx',index=False)


print(selecionadas)