from kia.constants import *
from kia.dealer import Dealer
from kia.log import Log
from kia.maestro import Maestro, AutomationTaskFinishStatus
from kia.onedrive import OneDrive
from kia.os import OS
from kia.outlook import Outlook
from kia.selenium import Selenium
from kia.worksheets.results_worksheet import ResultsWorksheet
from kia.xml_file import XMLFile
import datetime as dt
import kia.worksheets.results_worksheet as results_worksheet
import os
import shutil
import sys
import traceback


class Main:
    """Execução do processo para a marca KIA."""

    def __init__(self, execution_start_time):
        """Inicialização de objetos."""
        self.execution_start_time = execution_start_time
        self.maestro = Maestro()
        self.log = Log(self.maestro, "132_BOT_Importacao_de_XML_3_KIA")
        self.results_worksheet = ResultsWorksheet()    
        self.xml_file = XMLFile()
        self.selenium = Selenium()
        self.os = OS()
        self.onedrive = OneDrive(self.selenium, self.maestro, self.os, self.log)
        self.outlook = Outlook(self.log, self.maestro)
        self.dealer = Dealer(self.selenium, self.maestro, self.log)

        self.configure_maestro()
        self.create_folders()
        self.create_results_worksheet()
        #self.upload_previous_worksheet_to_maestro()
        self.download_models_worksheet()
        self.download_attachments_from_email()
        self.attachments_validation()
        self.send_data_to_datapool()
        self.dealer_process()
        #self.send_email()
        self.finish_automation()

    def configure_maestro(self):
        """Realiza as configurações necessárias do Maestro."""
        self.maestro.runner()
        self.log.info(message="")
        self.log.info(message="O login no Maestro foi realizado.")
        self.log.info(message="Iniciando o robô.")

    def create_folders(self):
        """Cria os diretórios necessários caso já não existam."""
        self.os.create_folder(WORKBOOK_DOWNLOAD_FOLDER_PATH)
        self.os.create_folder(DOCUMENTS_DOWNLOADS_FOLDER_PATH)
        self.os.create_folder(DOCUMENTS_FOLDER_PATH)
        self.os.create_folder(RESULTS_WORKBOOK_FOLDER_PATH)

    def create_results_worksheet(self):
        """Cria a planilha de resultados do dia atual fazendo uma cópia da planilha de modelo."""
        if os.path.exists(RESULTS_WORKBOOK_PATH) == False:
            shutil.copy2(RESULTS_WORKBOOK_MODEL_PATH, RESULTS_WORKBOOK_PATH)
            self.log.info(message="Planilha de resultados do dia atual gerada.")

    def upload_previous_worksheet_to_maestro(self):
        """Envia a planilha do dia anterior para os arquivos de resultado do Maestro, caso já não
        tenha sido enviada.
        """
        # Inicialização de variável.
        previous_worksheet_exists = False

        # Procura pela planilha de resultados da última execução.
        for days in range(1, 30):
            previous_worksheet_name = f"Resultado da Importação de {(dt.date.today() - dt.timedelta(days)).strftime("%d-%m-%Y")} - KIA.xlsx"
            previous_results_worksheet_path = f"{RESULTS_WORKBOOK_FOLDER_PATH}\\{previous_worksheet_name}"

            if os.path.exists(previous_results_worksheet_path) == True:
                previous_worksheet_exists = True
                break

        # Se não houver uma planilha anterior, sai do método.
        if not previous_worksheet_exists:
            return
        
        # Inicialização de variável.
        artifact_found = False

        # Verifica se a planilha já existe nos artefatos do Maestro.
        for artifact in self.maestro.list_artifacts():
            if artifact.name.strip().lower() == previous_worksheet_name.lower().strip():
                artifact_found = True
                break

        # Se a planilha não tiver sido encontrada no maestro, realiza o upload.
        if artifact_found == False:
            self.maestro.post_artifact(f"{previous_worksheet_name}", previous_results_worksheet_path)
            self.log.info("A planilha da última execução foi adicionada aos arquivos de resultado do Maestro.")

    def download_models_worksheet(self):
        """Baixa a planilha que possui os modelos dos veículos e as cores a serem selecionadas como
        opcional e em cor externa no Dealer."""
         # Inicializa o navegador.
        self.selenium.initialize_browser(headless=False, 
                                         default_download_folder=WORKBOOK_DOWNLOAD_FOLDER_PATH)

        # Inicialização de variável.
        max_attempts = 3
        login = True

        # Realiza até três tentativas para baixar a planilha com a relação de modelos.
        for attempt in range(max_attempts):
            try:
                self.onedrive.get_page()
                if login:
                    self.onedrive.login()
                    login=False
                self.os.delete_all_files(WORKBOOK_DOWNLOAD_FOLDER_PATH)
                self.onedrive.download_models_workbook()
                break
            except Exception as exception:
                # Trata a exceção personalizada "Timeout.".
                if ((str(type(exception)) == "<class 'Exception'>") and 
                    ((exception.args[0] == "Timeout."))
                    ):
                    self.log.warning(f"Problema ao identificar ou interagir com algum elemento ou ao baixar a planilha - Tentativa {str(attempt+1)} de 3.")
                    self.log.warning(traceback.format_exc())
                    if attempt < 2:
                        continue
                    else:
                        self.log.error(traceback.format_exc())
                        raise
                else:
                    self.log.error(traceback.format_exc())
                    raise
        
    def download_attachments_from_email(self):
        """Verifica os e-mails recebidos na caix "XML-Veiculos" e realiza o download."""
        self.log.info(message="Iniciando verificação de e-mail.")
        self.outlook.download_attachments()
        self.log.info(message="Verificação de e-mail concluída.")
    
    def attachments_validation(self):
        """Verifica cada arquivo baixado se é de veículo e se o destinatário é pessoa física."""
        self.log.info(message="Iniciando validação de documentos baixados.")
        
        # Percorre os arquivos na pasta e valida.
        for filename in os.listdir(DOCUMENTS_DOWNLOADS_FOLDER_PATH):
            # Validação.
            self.log.info(message=f"Validando o documentao {filename}.")
            self.xml_file.file_path = f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\{filename}"
            vehicle = self.xml_file.is_about_vehicle()
            recipient = self.xml_file.verify_recipient()

            # Remoção de arquivos que não serão importados.
            if (not vehicle) or (recipient != "legal person"):
                os.remove(f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\{filename}")
                self.log.info(message=f"Pelas regras de validação, o documento {filename} não será importado, portanto foi deletado.")

        self.log.info(message="Validação concluída.")

    def send_data_to_datapool(self):
        """Método para enviar os dados dos arquivos para o datapool e mover os arquivos. """
        self.log.info(message="Etapa de envio de itens para o datapool.")
        # Retora o datapool
        self.datapool = self.maestro.get_datapool("132_BOT_Importacao_de_XML_3_KIA")

        # Envia os itens ao datapool.
        for filename in os.listdir(DOCUMENTS_DOWNLOADS_FOLDER_PATH):
            filename = filename[:-4]
            self.log.info(message=filename)

            # Extrai dados do xml.
            self.xml_file.file_path = f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\{filename}.xml"
            xml_data = self.xml_file.get_data()

            # Envia dados ao datapool.
            self.datapool.create_entry({
                "nome_do_arquivo": filename, # Nome do arquivo sem extensão.
                "data_de_emissao": xml_data["data_de_emissao"],
                "nota_fiscal": xml_data["nota_fiscal"],
                "cnpj": xml_data["cnpj"],
                "valor": xml_data["valor"],
                "descricao_do_produto": xml_data["descricao_do_produto"],
                "chassi": xml_data["chassi"],
                "cor_externa": xml_data["cor_externa"],
                "codigo_do_modelo": xml_data["codigo_do_modelo"],
                "pintura": xml_data["pintura"],
                "sigla_da_pintura": xml_data["sigla_da_pintura"],
                "tipo_de_combustivel": xml_data["tipo_de_combustivel"],
                "informacoes_do_pedido": xml_data["informacoes_do_pedido"]
            })

            # Se o xml não existir na pasta de notas, é copiado. Se exisitir. é deletado da pasta
            # de download.
            if not os.path.exists(f"{DOCUMENTS_FOLDER_PATH}\\{filename}.xml"):
                # Envia os arquivos para a pasta de notas.
                shutil.move(f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\{filename}.xml", DOCUMENTS_FOLDER_PATH)
            else:
                os.remove(f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\{filename}.xml")

        self.log.info(message="Finalizada a etapa de envio de dados para o datapool.")

    def dealer_process(self):
        """Execução do processo de importação de acordo com as linhas da planilha. 
        
        Realiza leitura dos arquivos, passa os dados para a planilha de resultado e realiza o 
        processo de importação. Caso haja algum erro aleatório de interação com algum elemento, o 
        processo será recomeçado com a mesma nota, podendo ser realizadas até três tentativas 
        consecutivas. Caso a conexão do usuário tenha sido derrubada, o login será realizado 
        novamente até cinco vezes.
        """
        self.log.info(message="Iniciando etapa de importação no Dealer.")

        # Realiza o login no dealer.
        self.dealer.login()

        # Retorna o número de itens pendentes no datapool.
        number_of_items_in_datapool = self.datapool.summary()["countPending"]

        # Percorre os itens no datapool e importa os arquivos no Dealer.
        for _ in range(number_of_items_in_datapool):
            # Busca o próximo item do datapool.
            datapool_item = self.datapool.next()   
            
            # Número da nota fiscal que aparecerá no log durante a importação.
            self.log.document = datapool_item["nota_fiscal"]
            self.log.info(message=f"Nota a ser importada: {datapool_item["nome_do_arquivo"]}.")

            # Passa o item para o objeto da classe Dealer.
            self.dealer.datapool_item = datapool_item

            # Linha da planilha.
            worksheet_current_row = self.results_worksheet.get_xml_row(datapool_item["nome_do_arquivo"])
     
            # Verifica se a nota já não foi processada anteriormente.
            status = self.results_worksheet.get(results_worksheet.STATUS_COLUMN, 
                                                worksheet_current_row)
            if ((status != "Cor externa não encontrada no Dealer.") and
                (status != "Não foi possível selecionar a loja no Dealer.") and
                (status != "Modelo não encontrado na planilha de relação de modelos.") and
                (status != None)
                ):
                self.log.info("Nota já importada anteriormente.")
                datapool_item.report_done()
                continue

            # Preenche a planilha com os dados do datapool.
            self.results_worksheet.fill_in(datapool_item, worksheet_current_row)

            # Passa a linha da planilha para o objeto da classe Dealer.
            self.dealer.worksheet_current_row = worksheet_current_row


            # Inicialização de variável que indica sucesso na importação.
            self.dealer.xml_inserted = False       

            # Variáveis para realizar nova tentativas em caso de exceção.
            attempt = 0
            login_again_attempt = 0
            login_attempt = False

            # Serão realizadas até 3 tentativas consecutivas de realização do processo em caso de 
            # problema com interação com elementos.
            while attempt < 3: 
                try:
                    # Se a conexão cair, uma nova tentativa de login é feita (limite de cinco vezes).
                    if login_attempt == True:
                        self.dealer.login(try_again=True)
                        login_attempt = False
                        login_again_attempt += 1

                    # Página inicial para que cada processamento inicie do mesmo ponto de partida.
                    self.selenium.get_page("https://dealer.gruposaga.com.br/Portal/default.html")

                    # Modifica a loja na página do dealer.
                    self.dealer.change_store()

                    # Realiza a importação do xml do veículo.
                    xml_path = f"{DOCUMENTS_FOLDER_PATH}\\{datapool_item['nome_do_arquivo']}.xml"
                    self.dealer.import_xml(xml_path)

                    # Preenche os campos com as informações do veículo.
                    error_message = self.dealer.fill_in_informations()

                    # Preenche os campos na página de processamento de dados.
                    if error_message == None:
                        self.dealer.process_data()

                    # Informa ao datapool que o processamento do item foi concluído.
                    datapool_item.report_done()

                    # Retorna ao "frame" principal e fecha a janela de processamento.
                    self.selenium.switch_to_default_content()
                    self.selenium.click("div.x-tool.x-tool-close")
                    break

                except Exception as exception:
                    print(exception)

                    # Identifica se a conexão do usuário foi derrubada e em caso afirmativo, a 
                    # variável login será o indicador para que a conexão seja reestabelecida.
                    try:
                        error = self.selenium.find_element("span#span_vOCORRENCIA").text()
                        print(error)
                        if "ativo em outra estação" in error.lower():
                            if login_again_attempt < 5:
                                self.log.warning("Conexão derrubada, pois o usuário foi logado em outro local. Realizando o login novamente...")
                                login_attempt = True
                                continue
                            else:
                                attempt = 3 # Força a finalização.
                                self.log.error("O limite de vezes em que o login foi realizado após perder a conexão foi atingido.")
                                exit()
                    except:
                        pass                                         
                    
                    # Trata a exceção personalizada "Timeout.".
                    if ((str(type(exception)) == "<class 'Exception'>") and 
                        ((exception.args[0] == "Timeout."))
                        ):
                        self.log.warning(f"Problema ao identificar ou interagir com elemento - Tentativa {str(attempt+1)} de 3.")
                        self.log.warning(traceback.format_exc())

                        attempt +=1
                        if attempt < 3:
                            continue
                        else:
                            pass
                            # Reporta erro para o item no datapool.
                            datapool_item.report_error()

                            # Insere novamente no datapool para ser processado posteriormente.
                            self.datapool.create_entry({
                                "nome_do_arquivo": datapool_item["nome_do_arquivo"],
                                "data_de_emissao": datapool_item["data_de_emissao"],
                                "nota_fiscal": datapool_item["nota_fiscal"],
                                "cnpj": datapool_item["cnpj"],
                                "valor": datapool_item["valor"],
                                "descricao_do_produto": datapool_item["descricao_do_produto"],
                                "chassi": datapool_item["chassi"],
                                "cor_externa": datapool_item["cor_externa"],
                                "codigo_do_modelo": datapool_item["codigo_do_modelo"],
                                "pintura": datapool_item["pintura"],
                                "sigla_da_pintura": datapool_item["sigla_da_pintura"],
                                "tipo_de_combustivel": datapool_item["tipo_de_combustivel"],
                                "informacoes_do_pedido": datapool_item["informacoes_do_pedido"]
                            })


                    else:
                        self.log.document = None

                        # Reporta erro para o item no datapool.
                        datapool_item.report_error()

                        # Insere novamente no datapool para ser processado posteriormente.
                        self.datapool.create_entry({
                                "nome_do_arquivo": datapool_item["nome_do_arquivo"],
                                "data_de_emissao": datapool_item["data_de_emissao"],
                                "nota_fiscal": datapool_item["nota_fiscal"],
                                "cnpj": datapool_item["cnpj"],
                                "valor": datapool_item["valor"],
                                "descricao_do_produto": datapool_item["descricao_do_produto"],
                                "chassi": datapool_item["chassi"],
                                "cor_externa": datapool_item["cor_externa"],
                                "codigo_do_modelo": datapool_item["codigo_do_modelo"],
                                "pintura": datapool_item["pintura"],
                                "sigla_da_pintura": datapool_item["sigla_da_pintura"],
                                "tipo_de_combustivel": datapool_item["tipo_de_combustivel"],
                                "informacoes_do_pedido": datapool_item["informacoes_do_pedido"]
                            })
                        attempt = 3 # Forçar finalização.


                    # Se algum programa no UIPATH finalizar o navegador, aparece a palavra 
                    # "disconnected" na mensagem de erro. Será utilizada para reiniciar o programa.
                    if "disconnected" in str(exception).lower():
                        self.log.warning("O navegador pode ter sido encerrado a força.")
                        self.__init__()
                        sys.exit()

                    self.log.error(message=str(exception))
                    self.log.error(message=traceback.format_exc())
                    self.log.document = None
    
    def send_email(self):
        """Realiza o envio da planilha e do log por e-mail."""
        # Envia a planilha por e-mail para a área quando o processamento começar nos intervalos de
        # horários informados.
        if (((self.execution_start_time > dt.time(12, 0, 0)) and (self.execution_start_time < dt.time(13, 0, 0))) or  # Entre 12h e 13h.
            ((self.execution_start_time > dt.time(15, 0, 0)) and (self.execution_start_time < dt.time(16, 0, 0))) or  # Entre 15 e 16.
            ((self.execution_start_time > dt.time(18, 0, 0)) and (self.execution_start_time < dt.time(19, 0, 0)))  # Entre 18h e 19h.
            ): 
            self.outlook.send_partial_report()

        # Envia a planilha para a área e para a equipe de RPA.
        elif self.execution_start_time > dt.time(23, 0, 0): # Após 23h.
            self.outlook.send_complete_report()

    def finish_automation(self):
        """Finaliza o robô."""
        self.log.info("Automação finalizada.")
        self.maestro.finish_task(AutomationTaskFinishStatus.SUCCESS, "Tarefa finalizada com sucesso.")
