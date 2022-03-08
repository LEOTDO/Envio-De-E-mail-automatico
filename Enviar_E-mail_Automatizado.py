#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[23]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# pyautogui -> automatizar o seu mouse, seu teclado e a sua tela

# escrever o passo a passo em português

# Passo 1: Entrar no sistema da empresa (no caso: drive)
pyautogui.hotkey("ctrl","t") #abre uma nova aba
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter") # botao do teclado

time.sleep(3)

# Passo 2: Navegar no sistema até encontrar a base de dados
pyautogui.click(x=345, y=306, clicks=2)
time.sleep(3)

# Passo 3: Exportar a base de vendas
pyautogui.click(x=453, y=402)
pyautogui.click(x=1234, y=195)
pyautogui.click(x=1013, y=589)
time.sleep(5)


# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos vendidos)


# Passo 5: Enviar em email para a diretoria com os indicadores


# In[9]:


get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[33]:


# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos vendidos)
# pandas, numpy e openxL
import pandas as pd

tabela = pd.read_excel(r"C:\Users\55479\Downloads\Vendas - Dez.xlsx")
display(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# ### Vamos agora enviar um e-mail pelo gmail

# In[36]:


# Passo 5: Enviar em email para a diretoria com os indicadores
pyautogui.hotkey("ctrl","t") # nova aba
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox") # copia o link
pyautogui.hotkey("ctrl","v") # colar o link
pyautogui.press("enter") # entrar no site
time.sleep(5)

#clicar no botão escrever
pyautogui.click(x=136, y=201)
time.sleep(3)
#escrever destinatario (quem vai receber o e-mail)
pyautogui.write("pythonimpressionador+diretoria@gmail.com")
pyautogui.press("tab")

#escrever o assunto
pyautogui.press("tab") #pular pro campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")

#escrever o corpo do e-mail
pyautogui.press("tab") # pula pro campo de corpo do email
text = f"""
Prezados, Bom dia!

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Por: Leonardo Torres
"""
pyperclip.copy(text)
pyautogui.hotkey("ctrl","v")

#enviar o e-mail
pyautogui.hotkey("ctrl","enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[27]:


time.sleep(10)
pyautogui.position()


# In[ ]:





# In[ ]:




