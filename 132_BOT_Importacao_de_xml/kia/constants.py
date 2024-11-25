import datetime as dt

# Caminhos das pastas.
WORKBOOK_DOWNLOAD_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\KIA\\Planilha baixada"
DOCUMENTS_DOWNLOADS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\KIA\\Notas baixadas"
DOCUMENTS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\KIA\\Notas"
RESULTS_WORKBOOK_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML\\KIA\\Resultados"

# Caminhos dos arquivos.
STORES_WORKBOOK_PATH = "kia\\resources\\Automação Relação de Lojas.xlsx"
FUEL_WORKBOOK_PATH = "kia\\resources\\Combustível - Códigos.xlsx"
RESULTS_WORKBOOK_MODEL_PATH =  "kia\\resources\\Resultado da Importação.xlsx"
MAPPING_WORKBOOK_NAME = "Mapeamento - KIA.xlsx"
MAPPING_WORKBOOK_PATH = f"{WORKBOOK_DOWNLOAD_FOLDER_PATH}\\{MAPPING_WORKBOOK_NAME}"

RESULTS_WORKBOOK_NAME = f"Resultado da Importação de {dt.date.today().strftime("%d-%m-%Y")} - KIA"  # Data atual.
RESULTS_WORKBOOK_PATH = f"{RESULTS_WORKBOOK_FOLDER_PATH}\\{RESULTS_WORKBOOK_NAME}.xlsx" 

