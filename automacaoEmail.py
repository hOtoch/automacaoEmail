import pyautogui
import pyperclip
import time
import pandas as pd
import numpy
import openpyxl

pyautogui.PAUSE = 1

# Abrindo o Opera
pyautogui.hotkey('winleft','s')
pyautogui.write('opera')
pyautogui.press('enter')

# Entrando no drive
pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.press('enter')

time.sleep(3)

# Baixando a planilha
pyautogui.click(x=390, y=299,clicks = 2)
time.sleep(1)
pyautogui.click(x=390, y=299)
pyautogui.click(x=1476, y=183)
pyautogui.click(x=1237, y=621)
time.sleep(5)

# Lendo a planilha com Panda

tabela = pd.read_excel(r"C:\Users\hugoo\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Enviando o email

time.sleep(3)
pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://outlook.live.com/mail/0/')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=220, y=175)
pyautogui.write("pythonimpressionador+diretoria@gmail.com")
pyautogui.press('tab')
assunto = "Relatório de vendas"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')

texto = f"Prezados, \n\nO faturamento do mês foi de R$ {faturamento:,.2f}\n\n A quantidade de vendas foi de {quantidade:,}."

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')

pyautogui.hotkey('ctrl','enter')

