#PARTE 1 - Refinamento da tabela do SENAC / Construção de Dataframe


# Importação do arquivo Excel (ficha do Senac) para organização dos dados.

import pandas as pd

df = pd.read_excel("teste1.xlsx")

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

#df_final.to_csv("df_final.csv")
#df_final.to_excel("df_final.xlsx")

"""PARTE 2 - Construção do Envio Automático / WhatsApp"""

#contatos = pd.read_excel("df_final.xlsx")

# Importação do Selenium / Chrome WebDriver / Time / URL Lib

from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.by import By
import time
import urllib

#Cria a instância do navegador Chrome

minichrome = webdriver.Chrome()
minichrome.get("https://web.whatsapp.com/")

#Coloca um While para que o algoritmo aguarde 1 segundo e verifique se o elemento "pane-side" (cartão de contatos do WhatsApp) está ativo.

while len(minichrome.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)

for i, mensagem in enumerate(df_final["Mensagem"]):
  pessoa = df_final.loc[i, "Nome"]
  telefone = df_final.loc[i, "Telefone"]
  texto = urllib.parse.quote(mensagem)
  link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
  minichrome.get(link)
  while len(minichrome.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)
  minichrome.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
  time.sleep(12)


