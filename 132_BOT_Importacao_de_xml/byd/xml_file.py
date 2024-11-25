import datetime as dt
import re
import xml.dom.minidom


class XMLFile:
    """Classe para manipulação de arquivos xml da BYD."""

    def __init__(self, log):
        """Inicialização de variáveis."""
        self.log = log
        self.file_path = None

    def is_about_vehicle(self):
        """Verifica se o xml recebido é de veículo ou não."""
        file =  xml.dom.minidom.parse(self.file_path)

        try:
            file.getElementsByTagName("veicProd")[0]
            is_about_vehicle = True
        except IndexError:
            is_about_vehicle = False
        
        return is_about_vehicle

    def get_data(self):
        """Busca no xml os dados que serão utilizados no processo de importação."""
        file =  xml.dom.minidom.parse(self.file_path)

        # Obtenção dos dados.
        data_de_emissao = file.getElementsByTagName("dhEmi")[0].firstChild.nodeValue[:10]
        nota_fiscal = file.getElementsByTagName("nNF")[0].firstChild.nodeValue
        cnpj = file.getElementsByTagName("dest")[0].getElementsByTagName("CNPJ")[0].firstChild.nodeValue
        valor = file.getElementsByTagName("vNF")[0].firstChild.nodeValue
        chassi = file.getElementsByTagName("chassi")[0].firstChild.nodeValue
        descricao_do_produto =  file.getElementsByTagName("xProd")[0].firstChild.nodeValue
        codigo_do_produto = file.getElementsByTagName("cProd")[0].firstChild.nodeValue
        cor_externa = file.getElementsByTagName("xCor")[0].firstChild.nodeValue
        tipo_de_combustivel = file.getElementsByTagName("tpComb")[0].firstChild.nodeValue

        # Formatação da data de emissão.
        data_de_emissao = dt.datetime.strptime(data_de_emissao, "%Y-%m-%d").strftime("%d/%m/%Y")

        xml_data = {
            "data_de_emissao": data_de_emissao,
            "nota_fiscal": nota_fiscal,
            "cnpj": cnpj,
            "valor": valor,
            "chassi": chassi,
            "descricao_do_produto": descricao_do_produto,
            "codigo_do_produto": re.sub(r"\s+", " ", codigo_do_produto), # Substitui vários espaços por um.
            "cor_externa": cor_externa,
            "tipo_de_combustivel": tipo_de_combustivel,
            }
        
        return xml_data
    
    def verify_recipient(self):
        """Verifica se o destinatário é pessoa física ou pessoa jurídica e se for pessoa jurídica,
        verifica se é referente à Saga Exclusive."""
        file =  xml.dom.minidom.parse(self.file_path)

        recipient = None # Inicialização de variável.
        # Verifica se destinatário é pessoa física.
        try:
            file.getElementsByTagName("dest")[0].getElementsByTagName("CPF")[0].firstChild.nodeValue
            recipient = "natural person"
            print("Destinatário é pessoa física.")
        except IndexError:
            recipient = "legal person"


        # Verifica se a loja é a saga exclusive.
        if recipient == False:
            cnpj_loja = file.getElementsByTagName("dest")[0].getElementsByTagName("CNPJ")[0].firstChild.nodeValue
            if cnpj_loja == "21333642000263":
                recipient = "saga exclusive"

        return recipient
    
    def cnpj_validation(self):
        """Validação de CNPJ."""
        file =  xml.dom.minidom.parse(self.file_path)

        # Verifica o CNPJ.
        cnpj = file.getElementsByTagName("dest")[0].getElementsByTagName("CNPJ")[0].firstChild.nodeValue
        if cnpj == "21333642000263":
            import_file = False
        else:
            import_file = True

        return import_file