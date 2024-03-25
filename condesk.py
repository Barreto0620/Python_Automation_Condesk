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
pyautogui.write("sigiloso")
pyautogui.press("tab")
pyautogui.write(str("sigiloso"))
pyautogui.press("enter")
time.sleep(2)
 
# Passo 3: Acessar o Painel de Criação do Usuário
pyautogui.click(x=82, y=245)
pyautogui.click(x=61, y=287)
time.sleep(1.5)
pyautogui.click(x=311, y=237)
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
    time.sleep(0.10)
    pyautogui.press("down")
    pyautogui.press("tab")
    pyautogui.write("HEMC") 
    time.sleep(0.10)
    pyautogui.press("tab")
    pyautogui.write("000000 - Genérico")
    time.sleep(0.10)
    pyautogui.press("tab")
    pyautogui.write("Usuario")
    time.sleep(0.10)
    pyautogui.press("tab")
    pyautogui.write("Usuario")
    time.sleep(0.10)
    pyautogui.press("tab", presses=3)
    pyautogui.click(x=553, y=573)
    pyautogui.press("tab", presses=3)
    pyautogui.write(str(tabela.loc[linha, "senha"]))
    pyautogui.press("tab", presses=2)
    pyautogui.press("enter")
    time.sleep(1.5)
    pyautogui.click(x=1131, y=598)
    time.sleep(2)
    pyautogui.click(x=327, y=234)
    time.sleep(2)
    #Usuário Criado

pyautogui.alert('Cadastro Finalizado !!') 

#Abrindo planilha
pyautogui.press("win")
pyautogui.write("Documentos: condesk_register.xlsx")
time.sleep(1)
pyautogui.press("enter")
time.sleep(4)

#Atualizando informações
pyautogui.hotkey('ctrl', 'pageup')
pyautogui.click(x=17, y=256)
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