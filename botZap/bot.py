import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os


webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_cliente = workbook['folha']

for linha in pagina_cliente.iter_rows(min_row=2):
    nome = linha[0].value
    movel = linha[1].value
    data = linha[2].value
    valor = linha[3].value

    mensagem = f'Olá {nome} seu pagamento vence no dia {data.strftime("%d/%m/%Y")}. Favor faça o pagamento do seguinte {valor}€ valor'

    try:
        link_mensagem = f'https://web.whatsapp.com/send?phone={movel}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem)
        sleep(10)
        enter = pyautogui.locateCenterOnScreen('enter.png')
        sleep(5)
        pyautogui.click(enter[0], enter[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{movel}{os.linesep}')
