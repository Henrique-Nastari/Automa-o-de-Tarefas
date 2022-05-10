#!/usr/bin/env python
# coding: utf-8

# In[1]:


# !pip install pyautogui
# !pip install pyperclip


# In[23]:


import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

#Entrar no sistema:
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#demora alguns segundos para algum lugar em específico:
time.sleep(5)

#Navegar no sistema e encontrar a base de dados:
pyautogui.click(x=459, y=381, clicks = 2)
time.sleep(2)

#Exportar e fazer downloads:
pyautogui.click(x=500, y=425) #clicar no arquivo
pyautogui.click(x=1661, y=230) #clicar com o botão direito ou 3 pontinhos
pyautogui.click(x=1375, y=745) #fazer o download

#Importar base de dados:
tabela = pd.read_excel(r"C:\Users\nasta\Downloads\Vendas - Dez.xlsx") #sempre que usar read_"tipo do arquivo"
display(tabela) #ou print se estiver usando outra IDE

#Calcular os indicadores
#Faturamento (soma da coluna valor final):
faturamento = tabela ["Valor Final"].sum()

#Quantidade de produtos (soma da coluna quantidade):
quantidade = tabela ["Quantidade"].sum()

display (faturamento)
display (quantidade)

#Enviando e-mail de relatório:
#Abrir o e-mail desejado (link: email)
#Clicar no escrever, digite o destinatário, digite o assunto, escrever o corpo do e-mail e enviar:
pyautogui.hotkey("ctrl", "t") #vai abrir o email
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)
pyautogui.click(x=134, y=255) #clicar em escrever
time.sleep(2)
pyautogui.write("nastari.henrique@gmail.com") #detalhe que você deverá escrever aqui todos os e-mails desejados
pyautogui.press("tab") #selecionando o destinatário
pyautogui.press("tab") #passar para o campo do assunto
pyperclip.copy("Relatório")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") #passar para o corpo do e-mail

#Escrevendo o e-mail:
texto = f"""""
Prezados, bom dia.
O faturamento de ontem foi de: R${faturamento:,.2f} reais e a quantidade de produtos foi de: {quantidade:,} 

abs 
Henrique Nastari Corrêa
"""""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl", "enter") #envia o e-mail


# In[ ]:
