import datetime as dt

# Caminhos das pastas.
WORKBOOK_DOWNLOAD_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\HYUNDAI\\Planilha baixada"
DOCUMENTS_DOWNLOADS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\HYUNDAI\\Notas baixadas"
CANCELLED_DOCUMENTS_DOWNLOADS_FOLDER_PATH = f"{DOCUMENTS_DOWNLOADS_FOLDER_PATH}\\Canceladas"
DOCUMENTS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\HYUNDAI\\Notas"
CANCELLED_DOCUMENTS_FOLDER_PATH = f"{DOCUMENTS_FOLDER_PATH}\\Canceladas"
RESULTS_WORKBOOK_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\HYUNDAI\\Resultados"

# Caminhos dos arquivos.
STORES_WORKBOOK_PATH = "hyundai\\resources\\Automação Relação de Lojas.xlsx"
FUEL_WORKBOOK_PATH = "hyundai\\resources\\Combustível - Códigos.xlsx"
RESULTS_WORKBOOK_MODEL_PATH =  "hyundai\\resources\\Resultado da Importação.xlsx"
MAPPING_WORKBOOK_NAME = "Mapeamento - HYUNDAI.xlsx"
MAPPING_WORKBOOK_PATH = f"{WORKBOOK_DOWNLOAD_FOLDER_PATH}\\{MAPPING_WORKBOOK_NAME}"

RESULTS_WORKBOOK_NAME = f"Resultado da Importação de {dt.date.today().strftime("%d-%m-%Y")} - HYUNDAI"  # Data atual.
RESULTS_WORKBOOK_PATH = f"{RESULTS_WORKBOOK_FOLDER_PATH}\\{RESULTS_WORKBOOK_NAME}.xlsx" 
