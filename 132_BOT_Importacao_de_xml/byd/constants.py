import datetime as dt

# Caminhos das pastas.
DOWNLOADS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML-Teste\\BYD\\Downloads"
DOCUMENTS_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML-Teste\\BYD\\Notas"
RESULTS_WORKBOOK_FOLDER_PATH = "C:\\Automacao\\132_BOT_Importacao_de_XML-Teste\\BYD\\Resultados"

# Caminhos dos arquivos.
STORES_WORKBOOK_PATH = "byd\\resources\\Automação Relação de Lojas.xlsx"
FUEL_WORKBOOK_PATH = "byd\\resources\\Combustível - Códigos.xlsx"
RESULTS_WORKBOOK_MODEL_PATH =  "byd\\resources\\Resultado da Importação.xlsx"

RESULTS_WORKBOOK_NAME = f"Resultado da Importação de {dt.date.today().strftime("%d-%m-%Y")} - BYD"  # Data atual.
RESULTS_WORKBOOK_PATH = f"{RESULTS_WORKBOOK_FOLDER_PATH}\\{RESULTS_WORKBOOK_NAME}.xlsx" 


