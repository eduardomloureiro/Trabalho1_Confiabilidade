# -*- coding: utf-8 -*-
"""
Created on Tue Oct  3 16:02:03 2023

@author: Davi
"""

import numpy as np
from scipy.stats import binom
import pandas as pd
from itertools import combinations

carga = 2850


class Agrupamento:
    def __init__(self, pot, qtd, prob):
        self.pot = pot
        self.qtd = qtd
        self.prob = prob
        self.geraTabela()

    def geraTabela(self):
        self.resultados = []
        self.iterado = []
        for i in range(int(self.qtd) + 1):
            self.resultados.append((i * self.pot, binom.pmf(i, int(self.qtd), self.prob)))
            self.iterado.append(0)
        self.usina = pd.DataFrame(self.resultados,columns=['Pot','Prob'])
        self.usina['Iter'] = self.iterado
        self.usina['IterAg'] = np.nan
            
            
def pu(x):
    return x/100


data = pd.read_excel('C:/Users/Davi/OneDrive/Mestrado/Confiabilidade/Trabalho 1/Gerac.xlsx')
data['FOR'] = data['FOR'].apply(pu)




agrupamento = []
usinas = []
geracao_total = 0
for i in range(len(data['Usina'])):
    agrupamento.append(Agrupamento(data.loc[i]['Pot. Ativa Max'], data.loc[i]['Unidades'], data.loc[i]['FOR']))
    usinas.append(agrupamento[i].usina)
    geracao_total += data.loc[i]['Unidades']*data.loc[i]['Pot. Ativa Max']

potencia_para_falta = geracao_total - carga

lista = []
for i in range(len(usinas)):
    for x in range(len(usinas[i])):
        lista.append((i,x))

combinacoes = []
for i in combinations(lista,2):
    combinacoes.append(i)
    


resultados_final = []

for i in range(len(combinacoes)):
    soma = 0
    produto = 1
    for j in combinacoes[i]:
        soma += agrupamento[j[0]].resultados[j[1]][0]
        produto *= agrupamento[j[0]].resultados[j[1]][1]
    resultados_final.append((soma,produto))



copt = pd.DataFrame(resultados_final, columns=['Pot indisponivel', 'Probabilidade'])
copt = copt.groupby(by = ['Pot indisponivel'], as_index = False).sum()

perda_de_carga = []

for i in range(len(copt)):
    if copt.loc[i]['Pot indisponivel'] <=  potencia_para_falta:
        perda_de_carga.append(0)
    else:
        perda_de_carga.append(abs(copt.loc[i]['Pot indisponivel'] - potencia_para_falta))


copt['Perda de Carga'] = perda_de_carga
copt['x.p'] = copt['Perda de Carga'] * copt['Probabilidade']

EPNS = sum(copt['x.p'])
LOLP = sum(copt[copt['Perda de Carga'] > 0]['Probabilidade'])
LOLE = LOLP * 8760
EENS = EPNS * 8760


print('A LOLP do problema é: {}'.format(LOLP))
print('A EPNS do problema é: {} MW'.format(EPNS))
print('A LOLE do problema é: {} horas'.format(LOLE))
print('A EENS do problema é: {} MWh'.format(EENS))
















