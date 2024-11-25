from botcity.maestro import *

class Maestro:
    def __init__(self):
        """Inicialização de variável."""
        self.activity_label = None # Rótulo do log de execução.

    def login(self, workspace_login, workspace_key, task_id=None):
        """Realiza o login no maestro por meio do login e da chave. O task_id precisa ser passado
        como parâmetro para poder utilizar alguns métodos. O número do task_id para testes
        atualmente é 2166642.
        """
        self.maestro = BotMaestroSDK()
        self.maestro.login(
            server = "https://developers.botcity.dev",
            login = workspace_login,
            key = workspace_key
        )
        self.task_id = task_id

    def runner(self):
        """Integração com o Maestro a partir de um runner."""
        BotMaestroSDK.RAISE_NOT_CONNECTED = False
        self.maestro = BotMaestroSDK.from_sys_args()
        self.execution = self.maestro.get_execution()
        self.task_id = self.execution.task_id
    
    def finish_task(self, status, message, total_items=None, processed_items=None, failed_items=None):
        """Finaliza a tarefa no Maestro."""
        self.maestro.finish_task(
            task_id=self.maestro.task_id,
            status=status,
            message=message,
            total_items=total_items,
            processed_items=processed_items,
            failed_items=failed_items
        )
    
    def get_credential(self, credential_label, credential_key):
        """Retorna a tarefa do Maestro."""
        return self.maestro.get_credential(label=credential_label, key=credential_key)

    def get_datapool(self, datapool_label):
        """Retorna um objeto do datapool."""
        maestro_datapool = self.maestro.get_datapool(label=datapool_label) 
        datapool = DataPool(maestro_datapool, self.maestro, self.task_id)
        return datapool
    
    def get_log(self, date):
        return self.maestro.get_log(activity_label=self.activity_label, date=date)

    def list_artifacts(self):
        """Lista todos os artefatos existentes."""
        artifacts = self.maestro.list_artifacts()
        return artifacts
    
    def log(self, values_dictionary):
        """Criação de log."""
        self.maestro.new_log_entry(
            activity_label = self.activity_label,
            values = values_dictionary
        )
    
    def post_artifact(self, artifact_name, filepath):
        """Envia um arquivo para a lista de artefatos."""
        self.maestro.post_artifact(
            task_id=self.task_id,
            artifact_name=artifact_name,
            filepath=filepath
        )
        
    
class DataPool:
    """Classe para manipulação de um datapool do Maestro"""
    def __init__(self, datapool, maestro, task_id):
        """Inicialização de variáveis."""
        self.datapool = datapool
        self.maestro = maestro
        self.task_id = task_id

    def create_entry(self, values_dictionary):
        """Adiciona um item no datapool."""
        new_item = DataPoolEntry(values=values_dictionary)
        self.datapool.create_entry(new_item)

    def has_next(self):
        """Verifica se há mais itens no datapool."""
        return self.datapool.has_next()

    def next(self):
        """Retorna o próximo item do datapool."""
        return self.datapool.next(task_id=self.task_id)
    
    def summary(self):
        """Retorna informações sobre o datapool."""
        return self.datapool.summary()