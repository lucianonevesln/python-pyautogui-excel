import pyautogui
import time

# Acessar o Excel
pyautogui.press('win')
time.sleep(4)
pyautogui.write('excel')
time.sleep(4)
pyautogui.press('enter')
time.sleep(6)

# Acessar pasta de trabalho em branco
pyautogui.click(x=552, y=255)
time.sleep(4)

# Selecionando a planilha inteira
pyautogui.click(x=14, y=245)
time.sleep(4)

# Diminuir tamanho da fonte
pyautogui.click(x=296, y=79)
time.sleep(4)

# Alinhar ao Meio
pyautogui.click(x=348, y=80)
time.sleep(4)

# Centralizar
pyautogui.click(x=349, y=110)
time.sleep(4)

# Acessar aba de exibir
pyautogui.click(x=579, y=47)
time.sleep(4)

# Clicar em flag de Linhas de Grade
pyautogui.click(x=331, y=114)
time.sleep(4)

# Clicar em celula
pyautogui.click(x=55, y=267)
time.sleep(2)
pyautogui.write('pronto para uso')
time.sleep(4)

# Clique duplo para ajuste
pyautogui.click(x=93, y=250)
pyautogui.doubleClick()