"""Há duas classes neste arquivo.

A classe Selenium possui métodos criados para facilitar a realização de interações com o navegador, considerando possíveis erros/exceções que venham a surgir.

A classe Element tem por objetivo permitir que os métodos padrões do Selenium sejam utilizados no desenvolvimento sem que sejam chamados diretamente, pois caso haja atualização na bilioteca
e nas funções, as mudanças podem ser feitas apenas neste arquivo.
"""

from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException, NoSuchElementException, NoSuchFrameException, StaleElementReferenceException, WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time


class Selenium:
    """Classe para interação com métodos do selenium com tratamentos de exceções."""
    
    def initialize_browser(self, headless=True, default_download_folder=None):
        """Realiza configuração e inicialização do navegador."""
        # Configurações do navegador.
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--force-device-scale-factor=0.9")
        options.add_argument("--inprivate")

        if headless:
            options.add_argument("--headless=new")
        if default_download_folder != None:
            prefs = {"download.default_directory": default_download_folder}
            options.add_experimental_option("prefs", prefs)

        # Inicialização.
        self.driver = webdriver.Edge(options=options)

    def clear(self, selector, by=By.CSS_SELECTOR, wait_time=60):
        """Apaga o texto presente em um elemento."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                self.driver.find_element(by=by, value=selector).clear()
                break
            except (ElementNotInteractableException, NoSuchElementException, 
                    StaleElementReferenceException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")
        
    def click(self, selector, by=By.CSS_SELECTOR, wait_time=60):
        """Clica em um elemento de acordo com o seletor css."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                element = self.driver.find_element(by=by, value=selector)
                if element.is_displayed() and element.is_enabled():
                    element.click()
                    break
                else:
                    time.sleep(0.5)
                    if time.monotonic() > end_time:
                        raise Exception("Timeout.")
            except (ElementClickInterceptedException, ElementNotInteractableException, 
                    NoSuchElementException, StaleElementReferenceException, WebDriverException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")
                
    def delete_element(self, css_selector):
        """Deleta o elemento da página."""
        self.driver.execute_script(f"document.querySelector('{css_selector}').remove()")         
        
    def element_exists(self, selector, by=By.CSS_SELECTOR):
        """Verifica se elemento existe na página."""
        try:
            self.driver.find_element(by=by, value=selector)
            return True
        except:
            return False       
    
    def fill(self, selector, text, by=By.CSS_SELECTOR, wait_time=60, clear=False):
        """Envia caracteres para determinado elemento. alterando a variável clear para verdadeiro,
        o texto do campo é apagado antes de enviar o texto."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                field = self.driver.find_element(by=by, value=selector)
                if clear == True:
                    field.clear()
                field.send_keys(text)
                break
            except (ElementNotInteractableException, NoSuchElementException, 
                    StaleElementReferenceException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")
                
    def find_element(self, selector, by=By.CSS_SELECTOR):
        """Busca por um elemento pelo seletor css."""
        element = self.driver.find_element(by=by, value=selector)
        element_obj = Element(self.driver, element)
        return element_obj

    def find_elements(self, css_selector):
        """Busca por todos os elementos existentes pelo seletor css."""
        list_of_elements = []
        elements = self.driver.find_elements(by=By.CSS_SELECTOR, value=css_selector)
        for element in elements:
            element_obj = Element(self.driver, element)
            list_of_elements.append(element_obj)

        return list_of_elements
    
    def get_page(self, page_url):
        """Carrega a página da url passada como argumento."""
        self.driver.get(page_url)

    def get_text(self, selector, by=By.CSS_SELECTOR, wait_time=60):
        """Retorna o texto de um elemento."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                text = self.driver.find_element(by=by, value=selector).text
                break
            except (ElementNotInteractableException, NoSuchElementException, 
                    StaleElementReferenceException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")
        return text
    
    def hover(self, selector, by=By.CSS_SELECTOR):
        """Simula a interção de quando o mouse é colocado sobre um elemento sem realizar clique."""
        actions = ActionChains(self.driver)
        actions.move_to_element(self.driver.find_element(by=by, value=selector)).perform()
    
    def refresh(self):
        """Atualiza a página atual."""
        self.driver.refresh()

    def script(self, code):
        """Envia um código em javascript para o navegador."""
        message = self.driver.execute_script(code)
        return message
    
    def select(self, selector, text, by = By.CSS_SELECTOR, wait_time=60, return_exception=False):
        """Tenta selecionar o item do campo de seleção de acordo com o texto por um tempo."""
        end_time = time.monotonic() + wait_time               
        while True:
            try:
                element = self.driver.find_element(by, selector)
                Select(element).select_by_visible_text(text)
                if by == By.XPATH:
                    selected = self.driver.find_element(By.XPATH, f"{selector}/*[text()='{text}']").get_attribute("selected")
                    if selected != "false":
                        break
                else:
                    break
                break
            except (NoSuchElementException, StaleElementReferenceException):
                time.sleep(0.5)
                if (time.monotonic() > end_time) and not return_exception:
                    raise Exception("Timeout.")
                elif (time.monotonic() > end_time) and return_exception:
                    return "Timeout."

    def selected_element_text(self, css_selector):
        """Retorna o texto do elemento que está selecionado em um <select>."""
        element = self.driver.find_element(by=By.CSS_SELECTOR, value=css_selector)
        selected_element_text = Select(element).first_selected_option.text
        return selected_element_text

    def switch_to_default_content(self):
        """Retorna ao frame inicial (tag <iframe>)."""
        self.driver.switch_to.default_content()

    def switch_to_frame(self, frame, by=By.CSS_SELECTOR, wait_time=60):
        """Troca de frame (tag <iframe>)."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                if "str" in str(type(frame)):
                    self.driver.switch_to.frame(self.driver.find_element(by, frame))
                elif "int" in str(type(frame)):
                    self.driver.switch_to.frame(frame)
                else:
                    raise Exception("Tipo de dado não suportado.")
                break
            except (NoSuchElementException, NoSuchFrameException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")

    def switch_to_window(self, window_index):
        """Muda de aba."""
        self.driver.switch_to.window(self.driver.window_handles[window_index])

    def upload_file(self, selector, file_path, by=By.CSS_SELECTOR, wait_time=60):
        """Realiza o upload de um arquivo em um campo especificado."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                self.driver.find_element(by=by, value=selector).send_keys(file_path)
                break
            except (NoSuchElementException, StaleElementReferenceException):
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")

    def wait_for_element_to_appear(self, selector, by=By.CSS_SELECTOR, wait_time=60):
        """Aguarda até que o elemento informado apareça na página."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                self.driver.find_element(by=by, value=selector)
                break
            except (NoSuchElementException, StaleElementReferenceException):
                time.sleep(0.5)
                if time.monotonic() > end_time: 
                    raise Exception("Timeout.")

    def wait_for_element_to_disappear(self, selector, by=By.CSS_SELECTOR, wait_time=60):
        """Aguarda até que o elemento informado desapareça da página."""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                self.driver.find_element(by=by, value=selector)
                time.sleep(0.5)
                if time.monotonic() > end_time:
                    raise Exception("Timeout.")
            except (NoSuchElementException, StaleElementReferenceException):
                break

    def wait_for_one_of_the_elements_to_appear(self, selector_1, selector_2, by_1 = By.CSS_SELECTOR, by_2 = By.CSS_SELECTOR, wait_time=60):
        """Aguarde até que um dos elementos apareça na página"""
        end_time = time.monotonic() + wait_time
        while True:
            try:
                self.driver.find_element(by=by_1, value=selector_1)
                return 1
            except:
                try:
                    self.driver.find_element(by=by_2, value=selector_2)
                    return 2
                except:
                    if time.monotonic() > end_time:
                        raise Exception("Timeout.")

class Element:
    """Classe para utilização dos métodos do selenium sem tratamento de exceção."""
    def __init__(self, driver, element):
        self.driver = driver
        self.element = element

    def clear(self):
        """Apaga texto do elemento."""
        self.element.clear()

    def click(self):
        """Clica no elemento."""
        self.element.click()

    def element_exists(self, css_selector, by=By.CSS_SELECTOR):
        """Verifica se elemento existe no contexto do elemento já obtido."""
        try:
            self.element.find_element(by, css_selector)
            return True
        except:
            return False

    def find_element(self, selector, by=By.CSS_SELECTOR):
        """Procura por um elemento."""
        element = self.element.find_element(by, selector)
        element_obj = Element(self.driver, element)
        return element_obj
    
    def find_elements(self, selector, by=By.CSS_SELECTOR):
        """Procura por elementos."""
        list_of_elements = []
        elements = self.element.find_elements(by, selector)
        
        for element in elements:
            element_obj = Element(self.driver, element)
            list_of_elements.append(element_obj)

        return list_of_elements
    
    def element_exists(self, css_selector, by=By.CSS_SELECTOR):
        """Verifica se elemento existe na página."""
        try:
            self.element.find_element(by, css_selector)
            return True
        except:
            return False
    
    def get_attribute(self, attribute):
        """Retorna o valor de um atributo do elemento."""
        return self.element.get_attribute(attribute)

    def hover(self):
        """Realiza a mesma interação de colocar o mouse sobre um elemento."""
        actions = ActionChains(self.driver)
        actions.move_to_element(self.element).perform()

    def is_selected(self):
        """Verifica se o elemento está selecionado."""
        return self.element.is_selected()

    def send_keys(self, text):
        """Envia caracteres."""
        self.element.send_keys(text)

    def text(self):
        """Retorna o texto."""
        return self.element.text