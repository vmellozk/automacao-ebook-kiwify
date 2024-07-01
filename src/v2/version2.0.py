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
        time.sleep(5.5)

    #Agora começa o cadastro dos produtos
    def cadastrar_produto(self,
                          nome_produto='Hackers do Amanha',
                          descricao='Nao se iluda com esses gurus de internet, eles não sabem nada! Aqui voce encontra o melhor do Hacking Brasil. Conteudos atualizados diariamente e exercicios reais propostos. Seja sua propria evolucao!',
                          link_vendas='instagram.com/@devxxctor',
                          preco=60000):
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
        time.sleep(4)

    #Agora começa a personalizar o produto
    def personalizar_produto(self,
                             email_suporte= 'hackingbrasil@gmail.com',
                             nome_criador_produto= 'Victor Mello'):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_uma_categoria.png'))
        time.sleep(0.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/internet.png'))
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
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(1)
        pyautogui.press('down')
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
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/apresentacao_responsiva.png'))
        time.sleep(6.5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/compartilhar.png'))
        time.sleep(2.5)
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
        time.sleep(8)

    def salvar_ebook_canva(self):
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/disco_local_g.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/documentos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/salvar.png'))
        time.sleep(5)

    def adicionar_ebook_ao_produto(self,
                                   titulo='Seja um Hacker, do bem!'):
        self.titulo = titulo
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(1.5)

        pyautogui.click(pyautogui.locateCenterOnScreen('essential/area_de_membros.png'))
        time.sleep(5)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/adicionar.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/nome_modulo.png'))
        time.sleep(1)
        pyautogui.write(self.nome_produto + ' v1')
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/adicionar_modulo.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/botao_mais.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/adicionar_conteudo.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/titulo.png'))
        time.sleep(1)
        pyautogui.write(self.titulo)
        time.sleep(1)
        pyautogui.scroll(-1000)
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/selecione_do_computador_.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/mostrar_todos_arquivos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/disco_local_g.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/documentos.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/pesquisar_documentos.png'))
        time.sleep(1)
        pyautogui.write(self.nome_produto + ' ebook')
        time.sleep(3)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/pdf.png'))
        time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/abrir.png'))
        time.sleep(6)
        pyautogui.scroll(+1000)
        pyautogui.click(pyautogui.locateCenterOnScreen('essential/criar_e_publicar.png'))
        time.sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(10)

#Função principal
def main():
    automation = KiwifiAutomation()
    automation.abrir_kiwifi()
    automation.cadastrar_produto()
    automation.personalizar_produto()
    automation.abrir_canva()
    automation.baixar_ebook_canva()
    automation.salvar_ebook_canva()
    automation.adicionar_ebook_ao_produto()

#Condição de execução
if __name__ == '__main__':
    for _ in range(100):
        main()  
        time.sleep(10)
