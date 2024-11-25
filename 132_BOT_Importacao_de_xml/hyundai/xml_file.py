import datetime as dt
import re
import xml.dom.minidom


class XMLFile:
    """Classe para manipulação de arquivos xml da HYUNDAI."""

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
    
    def get(self, element):
        """Retorna o valor de um elemento."""
        file = xml.dom.minidom.parse(self.file_path)
        return file.getElementsByTagName(element)[0].firstChild.nodeValue

    def get_data(self):
        """Busca no xml os dados que serão utilizados no processo de importação."""
        file = xml.dom.minidom.parse(self.file_path)

        # Puxa dados do xml.
        data_de_emissao = file.getElementsByTagName("dhEmi")[0].firstChild.nodeValue[:10]
        nota_fiscal = file.getElementsByTagName("nNF")[0].firstChild.nodeValue
        cnpj = file.getElementsByTagName("dest")[0].getElementsByTagName("CNPJ")[0].firstChild.nodeValue           
        valor = file.getElementsByTagName("vNF")[0].firstChild.nodeValue
        chassi = file.getElementsByTagName("chassi")[0].firstChild.nodeValue
        codigo_do_produto = file.getElementsByTagName("cProd")[0].firstChild.nodeValue
        cor_externa = file.getElementsByTagName("xCor")[0].firstChild.nodeValue
        codigo_da_cor_externa = file.getElementsByTagName("cCor")[0].firstChild.nodeValue
        tipo_de_combustivel = file.getElementsByTagName("tpComb")[0].firstChild.nodeValue
        ano_fabricacao = file.getElementsByTagName("anoFab")[0].firstChild.nodeValue
        ano_modelo = file.getElementsByTagName("anoMod")[0].firstChild.nodeValue

        # Modelo completo do veículo, conforme no pdf.
        descricao_do_modelo =  file.getElementsByTagName("xProd")[0].firstChild.nodeValue
        modelo_do_veiculo = descricao_do_modelo
        informacoes_adicionais =  re.sub(r"\s+", " ", file.getElementsByTagName("infAdProd")[0].firstChild.nodeValue)
        informacoes_adicionais_list = informacoes_adicionais.replace(":", ": ").split("|")
        for list_item in informacoes_adicionais_list:
            if "chassi" not in list_item.lower(): 
                modelo_do_veiculo = modelo_do_veiculo + " " + list_item
            else:
                break
                
        # Busca o dado "Informações do Pedido" nas informações complementares.
        informacoes_complementares = file.getElementsByTagName("infCpl")[0].firstChild.nodeValue
        informacoes_do_pedido = re.search(r"Informações do Pedido:(.*?)\|", informacoes_complementares).group(1)
        
        # Formatação da data de emissão.
        data_de_emissao = dt.datetime.strptime(data_de_emissao, "%Y-%m-%d").strftime("%d/%m/%Y")

        # Retorna os dados obtidos a partir do xml.
        return {
            "data_de_emissao": data_de_emissao,
            "nota_fiscal": nota_fiscal,
            "cnpj": cnpj,
            "valor": valor,
            "chassi": chassi,
            "codigo_do_produto": re.sub(r"\s+", " ", codigo_do_produto), # Substitui vários espaços por um.
            "modelo_do_veiculo": modelo_do_veiculo,
            "ano_fabricacao": ano_fabricacao,
            "ano_modelo": ano_modelo,
            "cor_externa": cor_externa,
            "codigo_da_cor_externa": codigo_da_cor_externa,
            "tipo_de_combustivel": tipo_de_combustivel,
            "informacoes_do_pedido": informacoes_do_pedido
            }
    
