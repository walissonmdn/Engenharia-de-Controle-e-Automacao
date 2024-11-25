from byd.constants import *
from byd.selenium import By
from byd.worksheets.fuels_worksheet import FuelsWorksheet
from byd.worksheets.results_worksheet import ResultsWorksheet
from byd.worksheets.stores_worksheet import StoresWorksheet
from openpyxl import load_workbook
import byd.worksheets.results_worksheet as results_worksheet
import time


class Dealer:
    """Classe para interação com o dealer."""
    
    def __init__(self, selenium, maestro, log) :
        """Inicialização de variáveis."""
        self.selenium = selenium
        self.maestro = maestro
        self.log = log

        # Criação de objetos para manipulação de planilhas.
        self.results_worksheet = ResultsWorksheet()
        self.stores_worksheet = StoresWorksheet()
        self.fuel_worksheet = FuelsWorksheet()

        # As variáveis abaixo serão alteradas durante a execução do programa.
        self.worksheet_current_row = None
        self.datapool_item = None
        self.xml_inserted = None

    def login(self, try_again=False):
        """Realiza login no Dealer."""
        self.selenium.get_page("https://dealer.gruposaga.com.br/")
        self.selenium.wait_for_element_to_appear("input#vUSUARIO_IDENTIFICADORALTERNATIVO")
        time.sleep(1)
        if try_again == False:
            self.selenium.fill("input#vUSUARIO_IDENTIFICADORALTERNATIVO", 
                               self.maestro.get_credential("dealer", "soulbot_login"))
        self.selenium.fill("input#vUSUARIOSENHA_SENHA", 
                           self.maestro.get_credential("dealer", "soulbot_password"))
        self.selenium.click("input#IMAGE3")

        # Aguarda até a loja aparecer, pois isso signica que o login foi bem-sucedido.
        self.selenium.wait_for_element_to_appear("tr.x-toolbar-right-row > td:nth-child(4) > table.x-btn.x-btn-noicon")
        self.log.info(message="Login realizado com sucesso.")

    def change_store(self):
        """Busca a marca e a loja a serem selecionadas no Dealer, verifica se já estão selecionadas
        e se não estiverem, seleciona-as. Se a loja não estiver na planilha de relação de lojas ou 
        houver erro ao selecionar, lança exceção.
        """
        # Retorna o nome da marca e da loja que serão selecionadas no Dealer.
        make_and_store = self.stores_worksheet.get_make_and_store(self.datapool_item["cnpj"])
        make_to_be_selected = make_and_store[0]
        store_to_be_selected = make_and_store[1]

        # Verifica se foi possível encontrar a marca e a loja na planilha.
        if store_to_be_selected == None:
            self.results_worksheet.write(results_worksheet.STATUS_COLUMN, 
                                     self.worksheet_current_row, 
                                     "CNPJ não encontrado na planilha de relação de lojas para selecionar a loja no dealer.")
            raise Exception("CNPJ não encontrado na planilha de relação de lojas para selecionar a loja no dealer.")
        
        # Verifica se a loja já está selecionada.
        self.selenium.switch_to_default_content()
        selected_store_tr_element = self.selenium.find_elements("tr.x-toolbar-right-row")[-1]
        selected_store = selected_store_tr_element.find_element("button").text()
        if selected_store.lower().strip() == store_to_be_selected.strip().lower():
            self.log.info(message="A loja correta já está selecionada.")
            return

        # Inicizalização de variáveis.
        max_attempts = 4
        correct_store_selected = False
            
        for _ in range(max_attempts):
            # Clique para garantir que a lista de lojas será fechada para fazer uma nova tentativa.
            self.selenium.click("div#W5-timeout-menu")
            time.sleep(0.5)

            # Expande a lista de marcas.
            self.selenium.click("tr.x-toolbar-right-row > td:nth-child(4) > table.x-btn.x-btn-noicon")
            self.selenium.wait_for_element_to_appear("li:nth-child(1) > a.x-menu-item.x-menu-item-arrow.x-unselectable > span")
            
            # Procura pela marca.
            makes = self.selenium.find_elements("li > a.x-menu-item.x-menu-item-arrow.x-unselectable > span")
            for make in makes:
                if make.text().strip().lower() == make_to_be_selected.strip().lower():
                    make.hover()
                    break
            time.sleep(0.5)

            # Procura pela loja.
            stores_div = self.selenium.find_elements("div.x-menu.x-menu-floating")[-1]
            stores = stores_div.find_elements("span.x-menu-item-text")
            for store in stores:
                if store.text().strip().lower() == store_to_be_selected.strip().lower():
                    store.click()
                    break  

            # Verifica a loja selecionada..
            selected_store = self.selenium.find_element("tr.x-toolbar-right-row > td:nth-child(4) > table.x-btn.x-btn-noicon > tbody > tr:nth-child(2) > td:nth-child(2) > em > button.x-btn-text").text()
            if selected_store.lower().strip() == store_to_be_selected.strip().lower():
                correct_store_selected = True
                break
        
        # Armazena no log e na planilha se a loja foi selecionada corretamente ou não.
        if correct_store_selected:
            self.log.info(message="A loja foi selecionada no Dealer.")
        else:
            self.results_worksheet.write(results_worksheet.STATUS_COLUMN, 
                                     self.worksheet_current_row, 
                                     "Não foi possível selecionar a loja no Dealer.")
            raise Exception("Não foi possível selecionar a loja no Dealer.")

    def import_xml(self, xml_path):
        """Navega até a página de importação do xml e realiza a importação."""
        # Navega até a página de importação do xml.
        self.selenium.click("table[id*= 'Integração']") 
        self.selenium.hover("a[id*= 'XML-Importação']")
        self.selenium.click("a[id*= 'NotaFiscaldeVeículo']")
        self.selenium.switch_to_frame("div.x-window-body > iframe")
        
        # A variável self.xml_inserted serve como indicativo a respeito da importação do xml. Caso 
        # o xml tenha sido importado corretamente e o processo tenha retornado ao início em 
        # decorrência de algum problema de interação com o dealer, o processo será feito do início,
        # mas sem importar o xml novamente.
        if self.xml_inserted == False: 
            # Botão para direcionar à pagina de importação.
            self.selenium.click("input#IMAGE1")
            
            # Importação.
            self.selenium.click("input#IMAGE2")
            self.selenium.switch_to_frame("iframe#gxp0_ifrm")
            self.selenium.wait_for_element_to_appear("input#uploadfiles")
            time.sleep(0.5)
            self.selenium.upload_file("input#uploadfiles", xml_path)
            self.selenium.switch_to_default_content()
            self.selenium.switch_to_frame("div.x-window-body > iframe")
            self.selenium.click("input#TRN_ENTER")
            self.selenium.wait_for_element_to_appear("span#TEXTBLOCKDOWNLOAD > a > text")

            # Verifica se a importação ocorreu com sucesso.
            message = self.selenium.find_element("span#TEXTBLOCKDOWNLOAD > a > text").text()
            if "1 de 1" in message:
                self.xml_inserted = True
                self.log.info(message="Nota importada com sucesso.")
            else:
                self.results_worksheet.write(results_worksheet.STATUS_COLUMN, 
                                             str(self.worksheet_current_row), 
                                             "Não foi possível importar a nota.")
                raise Exception("Não foi possível importar a nota.")

            # Volta para a página anterior.
            self.selenium.click("input#TRN_CANCEL")
        else:
            self.log.info(message="Nota já importada.")

    def fill_in_informations(self):
        """Verifica em qual linha da tabela do dealer se encontra a nota importada 
        e preenche na página de dados do veículo os dados obtidos do xml.
        """
        # Aguarda até que a primeira linha da tabela apareça.
        self.selenium.wait_for_element_to_appear("table#GridintxmlContainerTbl") 
        self.selenium.delete_element("table#GridintxmlContainerTbl")
        self.selenium.wait_for_element_to_disappear("table#GridintxmlContainerTbl")
        self.selenium.clear("input#vINTEGRACAOXMLNF_DATAEMISSAOINICIAL")
        self.selenium.wait_for_element_to_appear("table#GridintxmlContainerTbl") 

        # Encontra a linha com a nota importada e clica no ícone para editar as informações.
        row_found = False # Inicialização de variável.
        table = self.selenium.find_elements("table#GridintxmlContainerTbl > tbody > tr")
        for table_row in table:
            document_number = table_row.find_element("td:nth-child(4) > span").text()
            if document_number == self.datapool_item["nota_fiscal"]:
                row_found = True
                time.sleep(0.5)
                table_row.find_element("td:nth-child(9) > a > img").click()
                break

        # Se a nota não tiver aparecido na tabela, informar na planilha e prosseguir para a próxima.
        if row_found == False:
            self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                         self.worksheet_current_row,
                                         "Linha com a nota inserida não encontrada na tabela do dealer após importação do xml.")
            raise Exception("Linha com a nota inserida não encontrada na tabela do dealer após importação do xml.")
        
        # Prossegue para a página "Dados Veículo".
        self.selenium.click("input#vUPDATE_0001")

        # Página: "Dados Veículo";
        self.selenium.wait_for_element_to_appear("select#vMARCA_CODIGO")
        self.log.info(message="Inserindo os dados do veículo.")

        # Marca.
        self.selenium.select("//select[@id='vMARCA_CODIGO']", "BYD", By.XPATH)

        # Procura o modelo pelo código.
        self.selenium.click("img#IMAGEMODELO")
        self.selenium.switch_to_frame("iframe#gxp0_ifrm")
        self.selenium.fill("input#vMODELOVEICULO_MODELOMARCA", self.datapool_item["codigo_do_produto"])

        # Aguarda até que a quinta linha desapareça, pois se desaparecer, significa que a pesquisa 
        # retornou algum ou nenhum resultado.
        self.selenium.wait_for_element_to_disappear("tr#GridContainerRow_0005")  
        
        # Verifica se o código do modelo retornou resultado.
        codigo_do_produto_found = False
        try:
            codigo_do_produto_browser = self.selenium.find_element("span#span_MODELOVEICULO_MODELOMARCA_0001").text()
        except:
            pass
        else:
            if ((len(codigo_do_produto_browser) != 0) and 
                (codigo_do_produto_browser.lower() == self.datapool_item["codigo_do_produto"].lower())
                ):
                codigo_do_produto_found = True
                self.selenium.click("input#vLINKSELECTION_0001")
                self.log.info(message="Modelo do veículo encontrado pelo código do produto.")
        
        # Pesquisa o modelo pela descrição se a pesquisa pelo código não tiver retornado o resultado.
        if codigo_do_produto_found == False:
            self.selenium.find_element("input#vMODELOVEICULO_MODELOMARCA").clear()
            self.selenium.fill("input#vMODELOVEICULO_MODELOMARCA", " ")
            self.selenium.wait_for_element_to_appear("tr#GridContainerRow_0005")
            time.sleep(0.5)
            self.selenium.fill("input#vMODELOVEICULO_DESCRICAO", self.datapool_item["descricao_do_produto"].replace(" ", "%"))

            # Aguarda até que a quinta linha desapareça, pois se desaparecer, significa que a 
            # pesquisa retornou algum ou nenhum resultado.
            self.selenium.wait_for_element_to_disappear("tr#GridContainerRow_0005")
            
            # Tenta selecionar o modelo, mas se não tiver aparecido, informa na planilha.
            try:
                descricao_do_produto_browser = self.selenium.find_element("span#span_MODELOVEICULO_DESCRICAO_0001").text()
            except:
                self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                         self.worksheet_current_row,
                                         "Modelo do veículo não encontrado no Dealer.")
                raise Exception("Modelo do veículo não encontrado no Dealer.")
            else:
                if ((len(descricao_do_produto_browser) != 0) and 
                    (descricao_do_produto_browser.lower() == self.datapool_item["descricao_do_produto"].lower())
                    ):
                    self.selenium.click("input#vLINKSELECTION_0001")
                    self.log.info(message="Modelo do veículo encontrado pela descrição do produto.")
                else:
                    self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                             self.worksheet_current_row,
                                             "Modelo do veículo não encontrado no Dealer.")
                    raise Exception("Modelo do veículo não encontrado no Dealer.")
  
        # Volta ao frame anterior.
        self.selenium.switch_to_default_content()
        self.selenium.switch_to_frame("div.x-window-body > iframe")
        time.sleep(1)
        self.selenium.wait_for_element_to_appear("#BTNCONFIRMAR")

        # Cor interna (Será sempre Preto/Cinza para a BYD).
        self.selenium.select("//select[@id='vINTEGRACAO_CORCODINTERNA']", "PRETO/CINZA", By.XPATH)

        # Cor externa (Tenta selecionar por até 10 segundos para garantir que as cores terão sido 
        # atualizadas após selecionado o modelo).
        selection_result = self.selenium.select("//select[@id='vINTEGRACAO_CORCODEXTERNA']",
                                                self.datapool_item["cor_externa"].upper(),
                                                By.XPATH,
                                                wait_time=10,
                                                return_exception=True)
        # Verifica se retornou exceção.
        if selection_result == "Timeout.":
            self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                         self.worksheet_current_row,
                                         "Cor externa não encontrada no Dealer.")
            raise Exception("Cor externa não encontrada no Dealer.")

        # Combustível.
        fuel = self.fuel_worksheet.get_fuel(self.datapool_item["tipo_de_combustivel"])
        if fuel != None:
            self.selenium.select("//select[@id='vCOMBUSTIVEL_CODIGO']", fuel, By.XPATH)
        else:
            raise Exception("Código do combustível não encontrado na planilha.")
    
        # Procedência.
        self.selenium.select("//select[@id='vINTEGRACAOXMLNFVEICULO_PROCEDENCIACOD']", 
                             "0-PRODUTO NACIONAL", 
                             By.XPATH)        
        # Observação.
        descricao = self.selenium.find_element("textarea#vINTEGRACAOXMLNFVEICULO_INFOADICIONAL").text()
        self.selenium.fill("#vINTEGRACAOXMLNF_INFADICIONAIS", descricao)

        # Clica no botão de confirmação.
        self.selenium.click("input#BTNCONFIRMAR")

        # Inicialização de variável.
        error_message = None
        
        # Verifica se houve erro.
        shown_element = self.selenium.wait_for_one_of_the_elements_to_appear("table[id^='TABLE'] > tbody > tr > td > div > span#gxErrorViewer > div",
                                                                             "input#BTNPROCESSAR")
        if shown_element == 1:
            error_message = self.selenium.find_element("table[id^='TABLE'] > tbody > tr > td > div > span#gxErrorViewer > div").text()
            self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                 self.worksheet_current_row,
                                 f'Mensagem de erro: "{error_message}"')
        else:
            self.log.info(message="Dados preenchidos com sucesso.")
        
        return error_message
        
    def process_data(self):
        """Campos que aparecem para preenchimento ao clicar no botão "Processar"."""
        # Navegando até a página.
        self.selenium.click("input#BTNPROCESSAR")

        self.log.info(message="Página de processamento de dados.")
        
        # Grupo Movimento.
        self.selenium.select("//select[@id='vNOTAFISCAL_GRUPOMOVIMENTO']", "Compra", By.XPATH)
        
        # Natureza Operação.
        natureza_operacao = self.selenium.find_element("span#span_vNOTAFISCAL_NATUREZAOPERACAODES").text()
        if ("compra de veiculo novo" in natureza_operacao.lower()) == False:
            self.selenium.click("input#NATOPE")
            self.selenium.switch_to_frame("iframe#gxp0_ifrm")
            self.selenium.select("//select[@id='vNOTAFISCAL_NATUREZAOPERACAOCOD']", 
                                 "COMPRA DE VEICULO NOVO", 
                                 By.XPATH)
            self.selenium.click("input#CONFIRMAR")
            self.selenium.switch_to_default_content()
            self.selenium.switch_to_frame("div.x-window-body > iframe")
        
        # Tipo Pessoa.
        self.selenium.select("//select[@id='vPESSOA_TIPOPESSOA']", "Montadora", By.XPATH)

        # Tipo Documento.
        tipo_documento = self.selenium.find_element("span#span_vNOTAFISCAL_TIPODOCUMENTODES").text()
        if ("nf-e - nota fiscal eletronica" in tipo_documento.lower()) == False:
            self.selenium.click("input#TIPODOC")
            self.selenium.switch_to_frame("iframe#gxp0_ifrm")
            self.selenium.select("//select[@id='vNOTAFISCAL_TIPODOCUMENTOCOD']", 
                                 "NF-E - NOTA FISCAL ELETRONICA", 
                                 By.XPATH)
            self.selenium.click("input#CONFIRMAR")
            self.selenium.switch_to_default_content()
            self.selenium.switch_to_frame("div.x-window-body > iframe")

        # Conta Gerencial.
        self.selenium.click("input#vNOTAFISCAL_CONTAGERENCIALCOD")
        conta_gerencial = self.selenium.find_element("input#vNOTAFISCAL_CONTAGERENCIALCOD")
        conta_gerencial.clear()
        conta_gerencial.send_keys("68")

        # Departamento
        self.selenium.select("//select[@id='vNOTAFISCAL_DEPARTAMENTOCOD']", 
                             "VEICULOS NOVOS", 
                             By.XPATH)

        # Estoque.
        self.selenium.select("select#vNOTAFISCALITEM_ESTOQUECOD", "VN - VEICULOS NOVOS", wait_time=5)
        
        # Condição Pagamento.
        self.selenium.select("//select[@id='vNOTAFISCAL_CONDICAOPAGAMENTOCOD']",
                             f"VEICULOS NOVOS BYD",
                             By.XPATH)

        # Agente Cobrador.
        self.selenium.select("//select[@id='vAGENTECOBRADOR_CODIGO']", "VEICULOS", By.XPATH)

        # Tipo de Faturamento.
        self.selenium.select("//select[@id='vVEICULOTIPOFATURAMENTO_CODIGO']",
                             f"FLOOR PLAN - BYD", By.XPATH)

        # Regra dos Tributos.
        considerar_regras_dos_tributos = self.selenium.find_element("input#vISUTILIZAREGRATRIBUTOICMSPISCOFINS")
        if considerar_regras_dos_tributos.is_selected() == True:
            considerar_regras_dos_tributos.click()
  
        # Finaliza o processo.
        self.selenium.click("input#IMGPROCESSAR")

        # Aguarda a mensagem de retorno e armazena.
        self.selenium.wait_for_element_to_appear("#TABLE1 > tbody > tr > td > div >span#gxErrorViewer > div")
        messages = self.selenium.find_elements("table#TABLE1 > tbody > tr > td > div >span#gxErrorViewer > div")

        # Formata a mensagem para que apareça uma mensagem por linha.
        result = ""
        for i in range(len(messages)):
            result += messages[i].text()
            if i != (len(messages) -1):
                result += "\n"

        self.log.info(message=f'Mensagem de resultado: "{result}"')
        
        # Armazena a mensagem na plinha.
        self.results_worksheet.write(results_worksheet.STATUS_COLUMN,
                                 self.worksheet_current_row,
                                 result) 
        input("Pressione ENTER para prosseguir.")