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
          


def Calcula_COPT(data, carga):
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
    resul_copt = [EPNS, EENS, LOLP, LOLE]
    return resul_copt





def pu(x):
    return x/100


data = pd.read_excel('Gerac.xlsx')
data['FOR'] = data['FOR'].apply(pu)

curva_de_carga = pd.read_excel('curva de carga.xlsx', skiprows=[1,2,3])
curva_de_carga.rename(columns={'Inverno ': 'Inverno sem','Unnamed: 2':'Inverno fds',
                               'Verão': 'Verao sem','Unnamed: 4':'Verao fds',
                               'Primavera/Outono': 'Prim/Out sem','Unnamed: 6':'Prim/Out fds'}, inplace=True)


curva_de_carga.drop(columns=['Unnamed: 0'],inplace=True)

valores_de_carga = curva_de_carga.stack()
valores_de_carga = pd.unique(valores_de_carga)
valores_de_carga = np.sort(valores_de_carga)

print(curva_de_carga)

def Calcula_Prob(patamar):
    inver_sem = curva_de_carga[curva_de_carga['Inverno sem'] == patamar]['Inverno sem'].size    *5*(8-1+52-44+2)
    inver_fds = curva_de_carga[curva_de_carga['Inverno fds'] == patamar]['Inverno fds'].size    *2*(8-1+52-44+2)
    verao_sem = curva_de_carga[curva_de_carga['Verao sem'] == patamar]['Verao sem'].size        *5*(30-18+1)
    verao_fds = curva_de_carga[curva_de_carga['Verao fds'] == patamar]['Verao fds'].size        *2*(30-18+1)
    priout_sem = curva_de_carga[curva_de_carga['Prim/Out sem'] == patamar]['Prim/Out sem'].size *5*(17-9+43-31+2)
    priout_fds = curva_de_carga[curva_de_carga['Prim/Out fds'] == patamar]['Prim/Out fds'].size *2*(17-9+43-31+2)
    return sum([inver_sem, inver_fds, verao_sem, verao_fds, priout_sem, priout_fds])
    
    
inv_sem = 5*(8-1+52-44+2)*24
inv_fds = 2*(8-1+52-44+2)*24
ver_sem = 5*(30-18+1)*24
ver_fds = 2*(30-18+1)*24
po_sem =  5*(17-9+43-31+2)*24
po_fds =  2*(17-9+43-31+2)*24

horasano = inv_fds + inv_sem + ver_sem + ver_fds + po_sem + po_fds

EPNS = []
EENS = []
LOLP = []
LOLE = []
probab = []

for i in valores_de_carga:
    copt = Calcula_COPT(data, carga*i/100)
    probab_i = Calcula_Prob(i)/horasano
    probab.append(probab_i)
    EPNS.append(copt[0]*probab_i)
    EENS.append(copt[1]*probab_i)
    LOLP.append(copt[2]*probab_i)
    LOLE.append(copt[3]*probab_i)
    
    # a cada iteração, plotar o valor de carga e a probabilidade
    print('Carga: {} MW'.format(i), 'Probabilidade: {}'.format(probab_i))

print('A LOLP do problema é: {}'.format(sum(LOLP)))
print('A EPNS do problema é: {} MW'.format(sum(EPNS)))
print('A LOLE do problema é: {} horas'.format(sum(LOLE)))
print('A EENS do problema é: {} MWh'.format(sum(EENS)))
