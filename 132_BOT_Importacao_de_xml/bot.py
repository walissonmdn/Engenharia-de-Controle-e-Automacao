"""Arquivo para incialização do programa."""
from byd.main import Main as BYD
from hyundai.main import Main as HYUNDAI
from kia.main import Main as KIA
import datetime as dt
from execution_verification import ExecutionVerification

# Horário em que se o robô está iniciando.
execution_start_time = dt.datetime.now().time()

# Verifica se já há uma tarefa de execução deste robô em andamento.
ExecutionVerification()

# Processo de importação de xml por marca.
BYD(execution_start_time)
HYUNDAI(execution_start_time)
KIA(execution_start_time)