from hyundai.constants import *
import time
from hyundai.selenium import Keys


class OneDrive:
    def __init__(self, selenium, maestro, os, log):
        """Inicialização de variáveis."""
        self.selenium = selenium
        self.maestro = maestro
        self.os = os
        self.log = log

    def get_page(self):
        """Navega até a página da pasta compartilhada."""
        self.selenium.get_page("https://saganet-my.sharepoint.com/:f:/r/personal/saga_rpa_gruposaga_com_br/Documents/132_BOT_Importacao_de_XML/HYUNDAI?csf=1&web=1&e=FUHA5n")
        self.log.info("Página do OneDrive aberta.")

    def login(self):
        """Realiza o login no OneDrive."""
        # Login.
        self.selenium.fill("input[type='email']", 
                           self.maestro.get_credential("outlook", "saga.rpa_gruposaga_login") + Keys.ENTER)

        # Password.
        self.selenium.click("input[type='password']")
        self.selenium.fill("input[type='password']", 
                           self.maestro.get_credential("outlook", "saga.rpa_gruposaga_password") + Keys.ENTER)
        
        # Confirmação.
        self.selenium.click("input#idBtn_Back")
        self.log.info("Login realizado.")
    
    def download_models_workbook(self):
        """Baixa a planilha e aguarda até que tenha sido completamente baixada."""
        # Seleciona e baixa a planelha.
        self.selenium.click("div[aria-label*='Mapeamento - HYUNDAI.xlsx'] > div:nth-child(1)")
        self.selenium.wait_for_element_to_appear("div[aria-label*='Mapeamento - HYUNDAI.xlsx'][aria-selected='true']")
        self.selenium.click("button[data-id='download']")
        self.log.info("Iniciando download da planilha.")

        # Inicialização de variável.
        download_finished = False 

        # Aguarda até que a planilha tenha sido baixada.
        end_time = time.monotonic() + 120
        while True:
            download_finished = self.os.path_exists(f"{WORKBOOK_DOWNLOAD_FOLDER_PATH}\\{MAPPING_WORKBOOK_NAME}")
            if download_finished:
                self.log.info(f'A planilha {MAPPING_WORKBOOK_NAME} foi baixada.')
                break
            elif time.monotonic() > end_time:
                self.log.warning(f'Não foi possível baixar a planilha "{MAPPING_WORKBOOK_NAME}"')
                raise Exception("Timeout.")
