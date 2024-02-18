import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('Escala.xlsx')
pagina_Escala = workbook['Sheet1']

for linha in pagina_Escala.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    dia = linha[2].value
    posição = linha[3].value
    
    mensagem = f'Paz Do Senhor, passando para lembrar que você está escalado (a) hoje. Função *{posição}*.     *Portanto, meus amados irmãos, sede firmes e constantes, sempre abundantes na obra do Senhor, sabendo que o vosso trabalho não é vão no Senhor*.*Coríntios 16:58*' 
    
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(60)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(60)
        pyautogui.click(seta[0],seta[1])
        sleep(60)
        pyautogui.hotkey('ctrl','w')
        sleep(60)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
