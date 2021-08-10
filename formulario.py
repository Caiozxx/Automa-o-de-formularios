import pyautogui
import pyperclip
import time
import pandas as pd
# passo 1- entrar no nosso sistema (entrar no link : https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga )
pyautogui.hotkey('ctrl', 't')
pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
time.sleep(5)
pyautogui.press('enter')
# passo 2 - navegar no sistema (entrar na pasta exportar e fazer o download do arquivo)
time.sleep(8)
pyautogui.click(x=347, y=271, clicks=2)
# passo 3 - baixar o arquivo de venda
time.sleep(5)
pyautogui.click(x=350, y=322)
time.sleep(2)
pyautogui.click(x=1163, y=145)
time.sleep(3)
pyautogui.click(x=987, y=554)
time.sleep(10)
# passo 4 - calcular faturamento e quantidade de produtos vendidos 
tabela = pd.read_excel(r"C:\Users\Caioz\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["quantidade"].sum()
display(tabela)
# passo 5 - enviar o email para a diretoria 
pyautogui.hotkey('ctrl', 't')
pyautogui.write('https://mail.google.com/mail/u/1/#inbox')
time.sleep(5)
pyautogui.press('enter')
# clicando no botao escrever
pyautogui.click(x=206, y=153)
# escrevendo para quem eu mando o email
pyautogui.write('codex@dnmx.org')
pyautogui.press('tab')
pyautogui.press('tab')
#escrever assunto
pyautogui.write('relatorio de vendas')
pyautogui.press('tab')
#escrever corpo do email
texto = f"""Prezados, Bom dia
o faturamento foi de R${faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}
abs 
caio
"""
pyautogui.write(texto)
#enviar email
pyautogui.hotkey('enter')
