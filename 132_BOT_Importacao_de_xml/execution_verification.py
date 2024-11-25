from byd.maestro import Maestro, AutomationTaskFinishStatus
import datetime as dt


class ExecutionVerification:
    def __init__(self):
        """Inicialização de variável e chamada de métodos para verificação."""
        # Inicialização de variável.
        log_labels = ["132_BOT_Importacao_de_XML_1_BYD",
                      "132_BOT_Importacao_de_XML_2_HYUNDAI",
                      "132_BOT_Importacao_de_XML_3_KIA"]

        # Chamada de métodos.
        self.initialize_maestro()
        for log_label in log_labels:
            self.set_log(log_label)
            finish_automation = self.check_last_data()
            if finish_automation: 
                self.finish_automation()
                quit()

    def initialize_maestro(self):
        """Inicialização do Maestro."""
        self.maestro = Maestro()
        self.maestro.runner()
    def set_log(self, log_label):
        """Passa o rótulo do log que será acessado."""
        self.maestro.activity_label = log_label
    
    def check_last_data(self):
        """Verifica a data e hora do log mais recente e se for menor que o tempo estimado, 
        significa que possivelmente já existe uma tarefa deste robô sendo executada."""
        log_data = self.maestro.get_log(dt.datetime.now().strftime("%d/%m/%Y"))
        if len(log_data) == 0:
            finish_automation = False
        else:
            most_recent_date = log_data[-1]["Data"]
            most_recent_date = dt.datetime.strptime(most_recent_date, "%d/%m/%Y, %H:%M:%S")
            current_date = dt.datetime.now()

            if (current_date - most_recent_date) < dt.timedelta(minutes=4):
                finish_automation = True
            else:
                finish_automation = False

        return finish_automation
        
    def finish_automation(self):
        """Finaliza o robô."""
        self.maestro.finish_task(AutomationTaskFinishStatus.SUCCESS, "Tarefa finalizada com sucesso.")  