import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

try:
    workbook = openpyxl.load_workbook('Escala.xlsx')  # Carrega o arquivo Excel
    pagina_Escala = workbook['Sheet1']

    for linha in pagina_Escala.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        dia = linha[2].value
        posição = linha[3].value
            
        mensagem = f'Paz Do Senhor, passando para lembrar que você está escalado(a) no culto dia {dia.strftime('%d/%m/%Y')}. Função *{posição}* \n\n*Portanto, meus amados irmãos, sede firmes e constantes, sempre abundantes na obra do Senhor, sabendo que o vosso trabalho não é vão no Senhor.\n1 Coríntios 16:58*'
        
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)   # Abre a mensagem do Whats
        
        input('')

except Exception as e:
    print("Ocorreu um erro:", e)
