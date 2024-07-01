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
    time.sleep(6)

#Agora começa o cadastro dos produtos
def cadastrar_produto(nome_produto='Pastel',
                      descricao='Frito ',
                      link_vendas='instagram.com/@devxxctor',
                      preco=1000,
                      email_suporte= 'pasteldocriador@gmail.com',
                      nome_criador_produto= 'Joao da Silva'):
    pyautogui.click(pyautogui.locateCenterOnScreen('essential/produtos_botao.png'))
    time.sleep(1)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/criar_produto.png'))
    time.sleep(1)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/continuar.png'))
    time.sleep(1)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/nome_produto.png'))
    time.sleep(1)
    pyautogui.write(nome_produto)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/descricao.png'))
    time.sleep(1)
    pyautogui.write(descricao * 20)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/link_vendas.png'))
    time.sleep(1)
    pyautogui.write(link_vendas)

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/preco.png'))
    time.sleep(1)
    pyautogui.write(str(preco))

    pyautogui.click(pyautogui.locateCenterOnScreen('essential/criar_produto_e_avancar.png'))
    time.sleep(4)