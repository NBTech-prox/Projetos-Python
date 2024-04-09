import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os


webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_cliente = workbook['Clientes']
