#PARTE 1 - Refinamento da tabela do SENAC / Construção de Dataframe


# Importação do arquivo Excel (ficha do Senac) para organização dos dados.

import pandas as pd

df = pd.read_excel(r'C:\Users\CTP-ADM\Desktop\PyWP\PyWP-main\teste1.xlsx')

df.head(100)

df.columns = [ 'col0','col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7', 'col8','col9', 'col10', 'col11', 'col12', 'col13', 'col14', 'col15', 'col16', 'col17', 'col18', 'col19', 'col20', 'col21', ]

#Armazenamento de informações importantes

curso = df.at[6, 'col0']
periodo = df.at[7, 'col0']
horario =  df.at[7, 'col10']
telefones = df['col17']
nomes = df['col6']

print(nomes)

#Correção das colunas nomes e telefones / remoção dos espaços em branco


nomes = nomes.drop(nomes.index[0:11])
telefones = telefones.drop(telefones.index[0:11])

#telefones = telefones.str.replace(" ", "")
aluno = nomes[11]

print(periodo)

print(nomes)

print(telefones)

#Mesclagem das colunas dentro de um único dataframe;

df_nomes = pd.DataFrame(nomes)
df_telefones = pd.DataFrame(telefones)

df_final = pd.concat([df_nomes, df_telefones], axis=1)
df_final['Mensagem'] = ''

df_final.columns = ['Nome','Telefone','Mensagem']

df_final.head(30)

df_final = df_final.reset_index(drop=True)

print(df_final)

# Função de preenchimento automático da mensagem;

def mensageiro(row):
    return f"Prezado(a) {row['Nome']}, bom dia! Você se matriculou para o {curso} do programa Minha Vez, e nossas aulas acontecerão no {periodo} ({horario}), no Centro de Treinamento Profissional, localizado na Rua Itagiba de Oliveira, 410 - Barra (Depois da UPA, ao lado do Cras Barra). Estamos fazendo contato para confirmamos a sua participação no curso. Pedimos a gentileza de nos responder confirmando ou não a sua participação. Aguardamos o seu retorno!"

df_final['Mensagem'] = df_final.apply(mensageiro, axis=1)

#@title Grava o dataframe em um novo arquivo .csv e .xlsx

df_final.to_csv("df_final.csv")
df_final.to_excel("df_final.xlsx")

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com")

# Aguarda o carregamento do WhatsApp

while len(navegador.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)
time.sleep(2) 

import urllib
import time
import os

for linha in df_final.index:
    nome = df_final.loc[linha, "Nome"]
    mensagem = df_final.loc[linha, "Mensagem"]
    telefone = df_final.loc[linha, "Telefone"]
    
    texto = urllib.parse.quote(mensagem)

    # Cria a mensagem
    link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
    
    navegador.get(link)
    
    # Aguarda o carregamento do WhastApp
    while len(navegador.find_elements(By.ID, 'side')) < 1: # -> lista for vazia -> que o elemento não existe ainda
        time.sleep(1)
    time.sleep(2) 
    
    # você tem que verificar se o número é inválido
    if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
        # enviar a mensagem
        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        
        time.sleep(5)