from PyPDF2 import PdfReader
import re


class PDFFile:
    """Classe para leitura de arquivo pdf."""

    def __init__(self, log):
        """Inicialização de variável."""
        self.log = log
        self.file_path =  None

    def get_data(self):
        """Busca o dado de modalidade no pdf, pois não consta no xml."""
        
        # Abertura e leitura do pdf.
        with open(self.file_path, 'rb') as file:
            reader = PdfReader(file)
            text = reader.pages[0].extract_text()

        # Inicialização de variáveis.
        modalidade = ""
        possible_expressions_list = [r"modalidade destinada a: ([^\-]+?)(?=-|ICMS)",
                                     r"destinado a modalidade:([^\-]+?)(?=-|ICMS)"]
        
        # Verifica se uma das expressões regulares retorna resultado.
        for regular_expression in possible_expressions_list:
            # Procura pelo expressão nos dados extraídos do pdf.
            result = re.search(regular_expression, 
                               text.replace("\n", " "), 
                               re.IGNORECASE)
            
            # Extrai a modalidade.
            if result != None:
                modalidade = result.group(1).strip()
                break
        
        # Retorna em formato de dicionário.
        return {"modalidade": modalidade}