# coding: utf-8

import pandas as pd
import pyautogui
from time import sleep
from IPython.display import display
import pyperclip

#abre google
pyautogui.click(705,1054,duration=1)
# #abre usuario google
pyautogui.click(1070,611,duration=1)
# #abre gmail
#pyautogui.click(851,494,duration=3)
pyautogui.click(233,50,duration=2)
pyautogui.write('https://mail.google.com/mail/u/5/#inbox')
pyautogui.click(289,73,duration=2)
tabela = pd.read_excel("Teste do bot python.xlsx")
display(tabela)
for linha in range(0,15):
    nome = tabela.loc[linha,"Nome"]
    gmail = tabela.loc[linha,"Email"]
    cod = tabela.loc[linha,"Código"]
    #clica em escrever email
    if (linha == 0):
        pyautogui.click(88,204,duration=4)
    else:
        pyautogui.click(88,204,duration=0.1)
    #clica em destinatario
    pyautogui.click(256,234,duration=2)
    pyautogui.write(gmail)
    #clica em assunto
    assunto = 'Processo Seletivo Milhagem UFMG- Instruções para a realização da prova'
    pyperclip.copy(assunto)
    pyautogui.click(172,276,duration=0.3)
    with pyautogui.hold('ctrl'):
        pyautogui.press(['v'])
    #tira a marca do mailtrack
    pyautogui.click(368,576,duration=2)
    pyautogui.click(1365,318,duration=3)
    

    corpo_email = f"""Olá, {nome}! Esperamos que esteja bem.

As instruções para a realização da Prova estão no PDF em anexo.
O código que você vai precisar para realizá-la no dia 12/04 é: {cod}

Qualquer dúvida, não hesite em nos perguntar.

Atenciosamente.
Equipe Milhagem UFMG.
    """
    pyperclip.copy(corpo_email)
    #clica no corpo
    pyautogui.click(176,281,duration=0.1)
    with pyautogui.hold('ctrl'):
        pyautogui.press(['v'])
    #clica em anexar
    pyautogui.click(419,979,duration=0.1)
    #clica em pesquisar
    pyautogui.click(532,770,duration=1)
    pyperclip.copy('Instruções para realização das Provas - PS 2023.pdf')
    with pyautogui.hold('ctrl'):
        pyautogui.press(['v'])
    #seleciona
    pyautogui.click(1040,800,duration=0.1)
    #envia email
    pyautogui.click(199,976,duration=4)
    sleep(1)