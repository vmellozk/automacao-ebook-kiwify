#Importando Bibliotecas
import webbrowser
import openpyxl
import time
import pyautogui

#Antes de tudo, precisa carregar a planilha com o OpenPyxl
def dados_planilha(arquivo):
    workbook = openpyxl.load_workbook(arquivo)
    sheet = workbook.active
    dados = [] #--> para ficar numa lista de memória e ser utilizado depois
    for row in sheet.iter_rows(min_row=2, values_only=True):
        dados.append(row)

#Depois de ler os dados, é necessário abrir o Kiwifi com o WebBrowser
def abrir_kiwifi():
    kiwifi_url = 'https://dashboard.kiwify.com.br'
    webbrowser.open(kiwifi_url)
    time.sleep(15)
