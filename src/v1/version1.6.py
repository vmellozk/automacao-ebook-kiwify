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
        time.sleep(4.5)

    #Agora começa o cadastro dos produtos
    def cadastrar_produto(self,
                          nome_produto='Pastel',
                          descricao='Temos com os sabores de: carne, 4 queijos, pizza, camarão, carne seca e queijo e presunto. Caldo de Cana Refil! Estamos localizados na Av. Central, nº 113.',
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
        time.sleep(0.5)
        pyautogui.write(nome_produto)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/descricao.png'))
        time.sleep(0.5)
        pyautogui.write(descricao)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/link_vendas.png'))
        time.sleep(0.5)
        pyautogui.write(link_vendas)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/preco.png'))
        time.sleep(0.5)
        pyautogui.write(str(preco))
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/criar_produto_e_avancar.png'))
        time.sleep(3)

    #Agora começa a personalizar o produto
    def personalizar_produto(self,
                             email_suporte= 'pasteldocriador@gmail.com',
                             nome_criador_produto= 'Joao da Silva'):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_uma_categoria.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/culinaria_e_gastronomia.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_do_computador.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/mostrar_todos_arquivos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/disco_local_g.png'))
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
        time.sleep(6)
        pyautogui.scroll(-1000)
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/email_suporte.png'))
        time.sleep(0.5)
        pyautogui.write(email_suporte)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/nome_criador_produto.png'))
        time.sleep(0.5)
        pyautogui.write(nome_criador_produto)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/salvar_produto.png'))
        time.sleep(3)
        pyautogui.scroll(+1000)

    def abrir_canva(self):
        url_canva = 'https://www.canva.com'
        webbrowser.open(url_canva)
        time.sleep(3.5)

    def baixar_ebook_canva(self):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/projetos_canva.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/busque_conteudo_canva.png'))
        time.sleep(0.5)
        pyautogui.write(self.nome_produto + ' ebook')
        pyautogui.hotkey('enter')
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/apresentacao_responsiva.png'))
        time.sleep(5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/compartilhar.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/baixar_ebook.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/formato_ebook.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/video_mp4.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/formato_ebook.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/pdf_padrao.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/baixar.png'))
        time.sleep(7)

    def salvar_ebook(self):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/disco_loca_g.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/documentos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/salvar.png'))
        time.sleep(3)

#Função principal
def main():
    automation = KiwifiAutomation()
    automation.abrir_kiwifi()
    automation.cadastrar_produto()
    automation.personalizar_produto()
    automation.abrir_canva()
    automation.baixar_ebook_canva()

#
if __name__ == '__main__':
        main()
