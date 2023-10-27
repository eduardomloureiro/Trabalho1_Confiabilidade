# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 18:32:05 2023

@author: Davi
"""

import numpy as np
import pandas as pd

def pu(x):
    return x/100


def Calcula_Horas(patamar):
    inver_sem = curva_de_carga[curva_de_carga['Inverno sem'] == patamar]['Inverno sem'].size*5*(8-1+52-44+2)
    inver_fds = curva_de_carga[curva_de_carga['Inverno fds'] == patamar]['Inverno fds'].size*2*(8-1+52-44+2)
    verao_sem = curva_de_carga[curva_de_carga['Verao sem'] == patamar]['Verao sem'].size*5*(30-18+1)
    verao_fds = curva_de_carga[curva_de_carga['Verao fds'] == patamar]['Verao fds'].size*2*(30-18+1)
    priout_sem = curva_de_carga[curva_de_carga['Prim/Out sem'] == patamar]['Prim/Out sem'].size*5*(17-9+43-31+2)
    priout_fds = curva_de_carga[curva_de_carga['Prim/Out fds'] == patamar]['Prim/Out fds'].size*2*(17-9+43-31+2)
    return sum([inver_sem, inver_fds, verao_sem, verao_fds, priout_sem, priout_fds])



inv_sem = 5*(8-1+52-44+2)*24
inv_fds = 2*(8-1+52-44+2)*24
ver_sem = 5*(30-18+1)*24
ver_fds = 2*(30-18+1)*24
po_sem = 5*(17-9+43-31+2)*24
po_fds = 2*(17-9+43-31+2)*24

horasano = inv_fds + inv_sem + ver_sem + ver_fds + po_sem + po_fds

data = pd.read_excel('C:/Users/Davi/OneDrive/Mestrado/Confiabilidade/Trabalho 1/Gerac.xlsx')
#===============CASO QUEIRA AGRUPAR AS USINAS COM MESMA GERAÇÃO==========
# data2 = data.groupby(by = ['Pot. Ativa Max','FOR'], as_index=False, sort=False,group_keys=False)['Unidades'].sum().reset_index()
# data2['Usina'] = data['Usina']
# data = data2
#========================================================================
data['FOR'] = data['FOR'].apply(pu)
data['Capacidade'] = data['Unidades']*data['Pot. Ativa Max']

curva_de_carga = pd.read_excel('C:/Users/Davi/OneDrive/Mestrado/Confiabilidade/Trabalho 1/curva de carga.xlsx', skiprows=[1,2,3])
curva_de_carga.rename(columns={'Inverno ': 'Inverno sem','Unnamed: 2':'Inverno fds',
                               'Verão': 'Verao sem','Unnamed: 4':'Verao fds',
                               'Primavera/Outono': 'Prim/Out sem','Unnamed: 6':'Prim/Out fds'}, inplace=True)


curva_de_carga.drop(columns=['Unnamed: 0'],inplace=True)

valores_de_carga = curva_de_carga.stack()
valores_de_carga = pd.unique(valores_de_carga)
valores_de_carga = np.sort(valores_de_carga)

patamares_de_carga = pd.DataFrame(valores_de_carga, columns = ['Patamar'])

patamares_de_carga['Prob'] = patamares_de_carga['Patamar'].apply(Calcula_Horas)/horasano
acumulada = []


for i in range(len(patamares_de_carga)):
    if i == 0:
        acumulada.append(patamares_de_carga.loc[i]['Prob'])
    else:
        
        acumulada.append(patamares_de_carga.loc[i]['Prob'] + acumulada[i-1])

patamares_de_carga['Acumulada'] = acumulada


PLOAD = 2850

tol = 0.01
NSmax = 1e4
NSmin = 100

beta_LOLP = tol + 1
NS = 0
soma_LOLP = 0
soma_quad_LOLP = 0
soma_EPNS = 0

prob = 0
horas = 0



while beta_LOLP > tol and NS < NSmax:
    NS += 1
    
    capacidade = sum(data['Capacidade'])
    
    U_carga = np.random.rand()

    carga = PLOAD
    for patamar in range(len(valores_de_carga)):
        if U_carga < patamares_de_carga.loc[patamar]['Acumulada']:
            carga = PLOAD * patamares_de_carga.loc[patamar]['Patamar']/100
            break


    for usi in range(len(data['Usina'])):
        for unid in range(int(data.loc[usi]['Unidades'])):
            Ui = np.random.rand()
            if Ui < data.loc[usi]['FOR']:
                capacidade -= data.loc[usi]['Pot. Ativa Max']
                
    if capacidade < carga:
        # LOLP
        soma_LOLP += 1
        soma_quad_LOLP += 1 ** 2  # p/ convergencia

        # EPNS
        soma_EPNS += (carga - capacidade)

    v_esperado_LOLP = soma_LOLP / NS  # --> RESPOSTA
    v_esperado_EPNS = soma_EPNS / NS

    # convergencia
    if NS > NSmin:
        variancia_LOLP = (soma_quad_LOLP - NS * v_esperado_LOLP**2) / (NS - 1)
        variancia_do_valor_esperado_LOLP = variancia_LOLP / NS
        beta_LOLP = np.sqrt(variancia_do_valor_esperado_LOLP) / v_esperado_LOLP
        if np.isnan(beta_LOLP):
            beta_LOLP = np.inf

precisao = 5

LOLP = v_esperado_LOLP
EPNS = v_esperado_EPNS
LOLE = LOLP * 8760
EENS = EPNS * 8760

print("LOLP: ", round(LOLP, precisao),
"\nEPNS: ", round(EPNS, precisao), " MW",
"\nLOLE: ", round(LOLE, precisao), " h/ano",
"\nEENS: ", round(EENS, precisao), " MWh/ano")    
            
            