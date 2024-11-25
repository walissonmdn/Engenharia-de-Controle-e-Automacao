import datetime as dt

class Log:
    """Classe para realização do log de execução no Maestro."""

    def __init__(self, maestro, activity_label):
        """Inicialização de variável e configuração no Maestro."""
        self.maestro = maestro
        self.document = None
        
        self.maestro.activity_label = activity_label

    def info(self, message=None):
        """Log para informação sobre a execução."""
        if self.document == None: 
            print(f"Informação - {message}")
        else: 
            print(f"Informação - {self.document} - {message} ")
        
        self.maestro.log({
            "log_type": "Informação",
            "document": self.document,
            "message": message,
            "date": dt.datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        })
        
    def warning(self, message=None):
        """Log para aviso."""
        if self.document == None: 
            print(f"Aviso - {message}")
        else: 
            print(f"Aviso - {self.document} - {message} ")
        
        self.maestro.log({
            "log_type": "Aviso",
            "document": self.document,
            "message": message,
            "date": dt.datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        })
        
    def error(self, message=None):
        """Log para erro."""
        if self.document == None: 
            print(f"Erro - {message}")
        else: 
            print(f"Erro - {self.document} - {message} ")
        
        self.maestro.log({
            "log_type": "Erro",
            "document": self.document,
            "message": message,
            "date": dt.datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
        })
