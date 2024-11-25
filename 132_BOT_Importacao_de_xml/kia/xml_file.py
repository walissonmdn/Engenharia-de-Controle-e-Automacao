import datetime as dt
import re
import xml.dom.minidom

class XMLFile:
    """Classe para manipulação de arquivos xml da KIA."""

    def __init__(self):
        """Inicialização de variáveis."""
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
    
    def verify_recipient(self):
        """Verifica se o destinatário é pessoa física ou pessoa jurídica e se for pessoa jurídica,
        verifica se é referente à Saga Exclusive."""
        file =  xml.dom.minidom.parse(self.file_path)

        # Inicialização de variável.
        recipient = None 
        
        # Verifica se destinatário é pessoa física.
        try:
            file.getElementsByTagName("dest")[0].getElementsByTagName("CPF")[0].firstChild.nodeValue
            recipient = "natural person"
        except IndexError:
            recipient = "legal person"

        # Retorna o tipo de destinatário.
        return recipient

    def get_data(self):
        """Busca no xml os dados que serão utilizados no processo de importação."""
        file =  xml.dom.minidom.parse(self.file_path)
        
        # Puxa dados do xml.
        data_de_emissao = file.getElementsByTagName("dhEmi")[0].firstChild.nodeValue[:10]
        nota_fiscal = file.getElementsByTagName("nNF")[0].firstChild.nodeValue
        cnpj = file.getElementsByTagName("dest")[0].getElementsByTagName("CNPJ")[0].firstChild.nodeValue
        valor = file.getElementsByTagName("vNF")[0].firstChild.nodeValue
        chassi = file.getElementsByTagName("chassi")[0].firstChild.nodeValue
        cor_externa = file.getElementsByTagName("xCor")[0].firstChild.nodeValue
        tipo_de_combustivel = file.getElementsByTagName("tpComb")[0].firstChild.nodeValue
        descricao_do_produto =  file.getElementsByTagName("xProd")[0].firstChild.nodeValue
        
        # Formatação da data de emissão.
        data_de_emissao = dt.datetime.strptime(data_de_emissao, "%Y-%m-%d").strftime("%d/%m/%Y")

        # Busca parte do código do modelo que se encontra nas inforamções complementares.
        informacoes_complementares = file.getElementsByTagName("infCpl")[0].firstChild.nodeValue
        regular_expression_result = re.search(r"MOD\.:\s*([^\.]+\.[^\.]+)", 
                                              informacoes_complementares,
                                              re.IGNORECASE)
        codigo_do_modelo = regular_expression_result.group(1)

        # Modalidade do veículo.
        if "veic.destinado" in informacoes_complementares.lower():
            regular_expression_result = re.search(r"(VEIC\.DESTINADO.*?)#", 
                                                  informacoes_complementares,
                                                  re.IGNORECASE)
            informacoes_do_pedido = regular_expression_result.group(1)
        else:
            informacoes_do_pedido = ""

        # Busca a pintura.
        if "-perolizada" in informacoes_complementares.lower():
            pintura = "PINTURA PEROLIZADA KIA"
            sigla_da_pintura = "PPK"
        elif "-metalica" in informacoes_complementares.lower():
            pintura = "PINTURA METALICA KIA"
            sigla_da_pintura = "PMK"
        else:
            pintura = None
            sigla_da_pintura = None

        # Retorna os dados obtidos a partir do xml.
        return {
            "data_de_emissao": data_de_emissao,
            "nota_fiscal": nota_fiscal,
            "cnpj": cnpj,
            "valor": valor,
            "descricao_do_produto": descricao_do_produto,
            "chassi": chassi,
            "cor_externa": cor_externa,
            "codigo_do_modelo": codigo_do_modelo,
            "pintura": pintura,
            "sigla_da_pintura": sigla_da_pintura,
            "tipo_de_combustivel": tipo_de_combustivel,
            "informacoes_do_pedido": informacoes_do_pedido
            }
    
