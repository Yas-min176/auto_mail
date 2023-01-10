import pyautogui
import pyperclip
import time
import pandas


pyautogui.PAUSE = 2.5

# comandos pyautogui
# pyautogui.click -> clicar
# pyautogui.write -> escrever
# pyautogui.press -> pressionar
# pyautogui.hotkey -> atalho

# pegar posição do mouse
# import time
# time.sleep(5)
# print(pyautogui.position())

# # Passo 1: entrar no sistema da empresa (no link do drive)
pyautogui.press("win")
pyautogui.write("opera")
pyautogui.press("enter")
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://www.youtube.com/redirect?event=video_description&redir_token=QUFFLUhqbDlYbFRXM1QxbHBRc0xUMk5uXzJpcmVHWjZNUXxBQ3Jtc0trMUpIa2dyWVBDUEtCUUpPVHRSMmhQYmJZVFR5NmY4Tkdqa25fWjVoZWlMTTZWTk4xRjREZmtvY2s1TFdhSXl2RjNjV0F3WWFHN2hSalVpN1NpMXA4U0YtZFl2eS1XWjNlZHByOF90NmMzMUNBRm9KTQ&q=https%3A%2F%2Fdrive.google.com%2Fdrive%2Ffolders%2F149xknr9JvrlEnhNWO49zPcw0PW5icxga%3Fusp%3Dsharing&v=JKOLBw5sHCw")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)

# Passo 2: navegar até o local do relatório (entrar na pasta exportar)
pyautogui.click(x=490, y=279, clicks=2)


# Passo 3: exportar o relatório (fazer o download)
pyautogui.click(x=490, y=279, clicks=1)
pyautogui.click(x=1158, y=191, clicks=1)
pyautogui.click(x=1045, y=621, clicks=1)


# Passo 4: calcular os indicadores (faturamento e quantidade de produtos)
tabela = pandas.read_excel(r"C:\Users\yasmi\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# passo 5: enviar um email para a diretoria
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/2/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=227, y=202, clicks=1)
pyautogui.write("101972017@eniac.edu.br")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
texto = f"""Prezados, bom dia
O faturamento de ontem foi de : R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("ctrl","enter")