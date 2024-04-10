import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import urllib
from time import sleep
import datetime
import os

navegador = webdriver.Chrome()
navegador.get('https://web.whatsapp.com')

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_cliente = workbook['folha']

for linha in pagina_cliente.iter_rows(min_row=2):
    nome = linha[0].value
    movel = linha[1].value
    data = linha[2].value
    valor = linha[3].value

    texto = f'Olá {nome} seu pagamento vence no dia {data.strftime("%d/%m/%Y")}. Favor faça o pagamento dos {valor}€'
    texto = urllib.parse.quote(texto)

    link = f'https://web.whatsapp.com/send?phone={movel}&text={texto}'
    navegador.get(link)

    while len(navegador.find_elements(By.ID, 'side')) < 1:
        sleep(2)
    sleep(4)

    navegador.find_element(
        By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
    sleep(5)
