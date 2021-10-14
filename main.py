import time
import pyautogui
import pyperclip
import pandas as pd

time.sleep(3)

pyautogui.PAUSE = 2
# tempo para dar tempo entre um comando e outro do pyautogui

# primeiro vamos ir ate o drive com automação do Chrome
pyautogui.hotkey('ctrl', 't')  # atalho para abrir uma nova página do Chrome
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
#endereço do navegador onde esta o arquivo a baixar
pyautogui.hotkey('ctrl', 'v')
# este comandos são necessários, pois o pyautogui não é capaz de ler caracteres especiais
pyautogui.press("enter")

# baixando o arquivo do drive
pyautogui.click(x=424, y=311, clicks=2)
time.sleep(2)
pyautogui.click(x=413, y=393, clicks=1)
pyautogui.click(x=1164, y=192, clicks=1)
pyautogui.click(x=969, y=623, clicks=1)

# # abrindo o arquivo baixado com pandas
vendas = pd.read_excel("digite aqui o caminho do seu arquivo a ser aberto")
print(vendas)
# somando o valor de faturamento e quantidade vendidas
faturamento = vendas["Valor Final"].sum()
quantidade = vendas["Quantidade"].sum()

# entrando no email com Gmail
time.sleep(3)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('digite aqui o endereço do navegador do seu e-mail')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(x=107, y=225, clicks=1)

#escrevendo o email
time.sleep(5)
pyautogui.click(x=986, y=309, clicks=1)
pyautogui.write("digite aqui o e-mail destinatário")
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write("Faturamento empresa")
pyautogui.press('tab')
#corpo do email
assunto =(f"""Bom dia
O faturamento foi de: R$ {faturamento: ,.2f},
A quantidade vendida foi de: {quantidade: }""")

time.sleep(2)
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.click(x=843, y=700, clicks=1)

#codigo para descobrir a localização para clicar
# x, y = pyautogui.position()
# print("x = "+str(x)+" y = "+str(y))
