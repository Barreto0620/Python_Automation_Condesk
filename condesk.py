# Passo a passo do projeto 

import pyautogui
import time
import pandas as pd

# Pause Global

pyautogui.PAUSE = 1.3   

# Passo 1: Entrar no Condesk
pyautogui.press("win")
pyautogui.write("firefox")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=367, y=64)
pyautogui.write("https://hemc.consulters.com.br/")
pyautogui.press("enter")

# Passo 2: Fazer login
#pyautogui.write("sigiloso")
#pyautogui.press("tab")
#pyautogui.write(str("sigiloso"))
pyautogui.press("enter")
time.sleep(2)
 
# Passo 3: Acessar o Painel de Criação do Usuário
pyautogui.click(x=96, y=274)
pyautogui.click(x=52, y=312)
time.sleep(1.5)
pyautogui.click(x=301, y=264)
time.sleep(1.5)

# Passo 3: Importando a base de dados para cadastro
tabela = pd.read_excel("condesk_register.xlsx")
print(tabela)

# Passo 4: Cadastrando Usuário
for linha in tabela.index:
    time.sleep(1)
    pyautogui.write(tabela.loc[linha, "nome_completo"])
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "matricula"]))
    pyautogui.press("tab")
    pyautogui.write(tabela.loc[linha, "login"])
    pyautogui.press("tab", presses=2)
    pyautogui.write(str(tabela.loc[linha, "telefone"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "celular"]))
    pyautogui.press("tab")
    pyautogui.write(tabela.loc[linha, "email"])
    pyautogui.press("tab", presses=2)
    pyautogui.write(tabela.loc[linha, "area"])
    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey("ctrl","c")
    time.sleep(0.20)
    pyautogui.press("down")
    pyautogui.press("tab")
    pyautogui.write("HEMC") 
    time.sleep(0.20)
    pyautogui.press("tab")
    pyautogui.hotkey("ctrl","v")
    time.sleep(0.20)
    pyautogui.press("tab")
    pyautogui.write("Usuario")
    time.sleep(0.20)
    pyautogui.press("tab")
    pyautogui.write("Usuario")
    time.sleep(0.20)
    pyautogui.press("tab", presses=3)
    pyautogui.click(x=554, y=604)
    pyautogui.press("tab", presses=3)
    pyautogui.write(str(tabela.loc[linha, "senha"]))
    pyautogui.press("tab", presses=2)
    pyautogui.press("enter")
    time.sleep(1.5)
    pyautogui.click(x=1092, y=612)
    time.sleep(2)
    #Usuário Criado

    # Passo 5: Fechar o Navegador
    pyautogui.alert('Cadastro Finalizado !!')
    time.sleep(2.5)
    pyautogui.press("enter")
    pyautogui.hotkey('alt', 'f4') 

#Abrindo planilha
pyautogui.press("win")
pyautogui.write(r"C:\Users\gabriel.oliveira\Documents\condesk_register.xlsx")
time.sleep(1)
pyautogui.press("enter")
time.sleep(4)
pyautogui.hotkey('win', 'up')

#Atualizando informações
pyautogui.hotkey('ctrl', 'pageup')
pyautogui.click(x=16, y=268)
pyautogui.dragTo(x=16, y=598, duration=1)
pyautogui.hotkey("ctrl","x")
pyautogui.hotkey('ctrl', 'pagedown')
pyautogui.hotkey("ctrl","home")
pyautogui.hotkey("ctrl","down")
pyautogui.press("down")
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey('ctrl', 'pageup')
pyautogui.hotkey("ctrl","home")
pyautogui.hotkey("ctrl","b")
pyautogui.hotkey("alt","f4")

pyautogui.alert('Atualização Finalizada !!')

#FIM