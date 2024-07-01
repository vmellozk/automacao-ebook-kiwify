#Importando Bibliotecas
import webbrowser
import openpyxl
import time
import pyautogui

class KiwifiAutomation():
    def __init__(self):
        self.nome_produto = None # --> para garantir que nome_produto possa existir em qualquer instancia da classe

    #Antes de tudo, precisa carregar a planilha com o OpenPyxl
    def dados_planilha(self, arquivo):
        workbook = openpyxl.load_workbook(arquivo)
        sheet = workbook.active
        dados = [] #--> para ficar numa lista de memória e ser utilizado depois
        for row in sheet.iter_rows(min_row=2, values_only=True):
            dados.append(row)
        return dados

    #Depois de ler os dados, é necessário abrir o Kiwifi com o WebBrowser
    def abrir_kiwifi(self):
        kiwifi_url = 'https://dashboard.kiwify.com.br'
        webbrowser.open(kiwifi_url)
        time.sleep(6)

    #Agora começa o cadastro dos produtos
    def cadastrar_produto(self,
                          nome_produto='Pastel',
                          descricao='Frito ',
                          link_vendas='instagram.com/@devxxctor',
                          preco=1000):
        self.nome_produto = nome_produto
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

    #Agora começa a personalizar o produto
    def personalizar_produto(self,
                             email_suporte= 'pasteldocriador@gmail.com',
                             nome_criador_produto= 'Joao da Silva'):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_uma_categoria.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/culinaria_e_gastronomia.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_do_computador.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/mostrar_todos_arquivos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/este_computador.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/imagens.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/pesquisar_imagens.png'))
        time.sleep(1)
        pyautogui.write(self.nome_produto)
        time.sleep(3)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/miniatura_foto_produto.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/abrir.png'))
        time.sleep(5)
        pyautogui.scroll(-1000)
        time.sleep(2)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/email_suporte.png'))
        time.sleep(1)
        pyautogui.write(email_suporte)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/nome_criador_produto.png'))
        time.sleep(1)
        pyautogui.write(nome_criador_produto)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/salvar_produto.png'))
        time.sleep(3)
        pyautogui.scroll(+1000)

    def abrir_canva():
        url_canva = 'https://www.canva.com'
        webbrowser.open(url_canva)
        time.sleep(6)

#Função principal
def main():
     automation = KiwifiAutomation()
     automation.abrir_kiwifi()
     automation.cadastrar_produto()
     automation.personalizar_produto()

#
if __name__ == '__main__':
        main()
