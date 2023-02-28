import pyautogui
import pyperclip
import time
import pandas as pd
import openpyxl
import displayfunction as display

pyautogui.PAUSE = 1

#Passo 1 abrir o sistema
pyautogui.press('win') # Pressiona a tecla windows
pyautogui.write('chrome') # Escreve a palavra chrome
pyautogui.press('Enter') # Pressiona Enter
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl','v') # Aperta ctrl+c
pyautogui.press('Enter') # Aperta Enter

time.sleep(2) # Delay de 2 segundos para carregar a pagina
pyautogui.hotkey('win','up')
print(pyautogui.position())
pyautogui.click(x=355, y=267, clicks=2)
time.sleep(1)
pyautogui.click(x=355, y=287, button='right')
time.sleep(1)
pyautogui.click(x=463, y=640)

# Ler os dados de uma tabela em excel
time.sleep(5)
tabela = pd.read_excel(r'C:\Users\flavio.escobar\Downloads\Vendas - Dez.xlsx')
print(tabela)

# Calcular o faturamento
faturamento = tabela['Valor Final'].sum() #esse .sum() server pra somar
qtde_produtos = tabela['Quantidade'].sum()

print(faturamento)
print(qtde_produtos)

pyautogui.click(x=370, y=53)
pyautogui.press('backspace')
pyautogui.write('gmail.com')
pyautogui.press('enter')
time.sleep(4)
pyautogui.click(x=87, y=173)
pyautogui.click(x=900, y=692)
pyautogui.write('flavioescobar93@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write('Faturamento de vendas')
pyautogui.press('tab')

texto = f'''Boa noite,
O faturamento de hoje foi: {faturamento:,.2f}
O total de produtos vendidos hoje foi: {qtde_produtos:,}
'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
time.sleep(1)
pyautogui.hotkey('ctrl','enter')