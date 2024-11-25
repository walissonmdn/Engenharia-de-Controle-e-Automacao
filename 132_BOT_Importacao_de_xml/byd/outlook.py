from byd.constants import *
from byd.worksheets.results_worksheet import ResultsWorksheet
from email import encoders
from email.header import decode_header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime as dt
import email
import imaplib
import smtplib
import traceback


class Outlook:
    """Realiza conexão com o outlook para leitura e envio de e-mail."""

    def __init__(self, log, maestro):
        """Inicialização de variáveis e parâmetros de conexão com o Outlook."""
        self.log = log
        self.maestro = maestro

        # Conexão com o e-mail.
        self.imap_server = "imap.gmail.com"
        self.smtp_server = "smtp-mail.outlook.com: 587"
    
    def download_attachments(self):
        """Realiza leitura do e-mail e baixa os arquivos necessários."""
        # Tenta realizar até três tentativas de acesso ao outlook para baixar os arquivos 
        # necessários.
        for _ in range(3):
            try:
                # Realiza a conexão com o e-mail.
                mail = imaplib.IMAP4_SSL(self.imap_server, port=993)
                mail.login(self.maestro.get_credential("EMAIL_MANAGER", "GMAIL_LOGIN"), 
                           self.maestro.get_credential("EMAIL_MANAGER", "GMAIL_PWD"))
                self.log.info("O login no gmail foi realizado.")

                # Seleciona a caixa onde chegam os arquivos e impede que as mensagens de e-mail sejam 
                # marcadas como lidas automaticamente.
                mail.select(mailbox='"XML - Veiculos"', readonly=True)

                # Busca os índices das mensagens de e-mail não lidas.
                _, messages_index = mail.search(None, "UNSEEN")

                # Verifica as mensagens.
                for message_index in messages_index[0].split():
                    # Faz com que o modo seja de apenas leitura para que as mensagens não sejam 
                    # marcadas como lidas só de serem acessadas.
                    mail.select(mailbox='"XML - Veiculos"', readonly=True)

                    # Armazena a mensagem.
                    _, data = mail.fetch(message_index, "(RFC822)")

                    e_mail = email.message_from_bytes(data[0][1])
                
                    # Decodifica o assunto da mensagem.
                    encoding =  decode_header(e_mail.get("subject"))[0][1]
                    if encoding == None:
                        subject = decode_header(e_mail.get("subject"))[0][0]
                    else:
                        encoded_subject = decode_header(e_mail.get("subject"))[0][0]
                        subject = encoded_subject.decode(encoding)

                    # Caso esteja especificado no assunto do e-mail que se trata de nota fiscal da BYD,
                    # baixar o xml.
                    if "byd - nota fiscal eletrônica" in subject.lower():
                        for part in e_mail.walk():
                            if part.get_content_maintype() == "multipart":
                                continue
                            if part.get("Content-Disposition") is None:
                                continue
                            attachment = part.get_filename()

                            if ".pdf" in attachment.lower():
                                # Recebe arquivo para escrita e salva no computador.
                                file_object = open(f"{DOWNLOADS_FOLDER_PATH}\\{attachment}", "wb")
                                file_object.write(part.get_payload(decode = True))
                                self.log.info(f"Arquivo baixado: {attachment}.")

                            if ".xml" in attachment.lower():
                                # Recebe arquivo para escrita e salva no computador.
                                file_object = open(f"{DOWNLOADS_FOLDER_PATH}\\{attachment}", "wb")
                                file_object.write(part.get_payload(decode = True))

                                # Permite marcar a mensagem como lida e marca.
                                mail.select(mailbox='"XML - Veiculos"', readonly=False)
                                mail.store(message_index, "+FLAGS", '\\Seen') # Marca como lida.
                                self.log.info(f'Arquivo baixado: "{attachment}".')
                mail.close()
                mail.logout()
                self.log.info("A conexão com o gmail foi encerrada.")
                break
            except:
                self.log.error(f'Erro durante o processo de verificação do outlook.')
                self.log.error(str(traceback.format_exc()))

    def send_partial_report(self):
        """Envia o relatório da execução para a área."""
        recipients = ["nfebsb@gruposaga.com.br",
                      "gerentefaturamentobsb@gruposaga.com.br",
                      "caio.gmandarino@gruposaga.com.br",
                      "caroline.fmessias@gruposaga.com.br",
                      "jhennyfer.smoreira@gruposaga.com.br",
                      "nfekiabsb@gruposaga.com.br"]
        
        # Remetente, destinatário e Assunto do e-mail.
        message = MIMEMultipart()
        message["From"] = "Saga RPA <saga.rpa@gruposaga.com.br>"
        message["To"] = ", ".join(recipients)
        message["Subject"] = f"RPA - Importação de XML - BYD"

        # Mensagem (HTML).
        html_message = f"""
            <div>
                <p>Olá!</p>
                <p>Segue em anexo a planilha com os resultados da importação dos arquivos de veículos da BYD.</p>
                <p>Data de execução: {dt.date.today().strftime("%d/%m/%Y")}.</p>
            </div>
            """
        html_part = MIMEText(html_message, "html")
        message.attach(html_part)

        # Anexa a planilha de resultados.
        attachment = open(RESULTS_WORKBOOK_PATH, "rb")
        attachment_package = MIMEBase("application", "octet-stream")
        attachment_package.set_payload(attachment.read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header("content-disposition", "attachment", filename=f"{RESULTS_WORKBOOK_NAME}.xlsx")
        message.attach(attachment_package)

        # Login.
        stmp = smtplib.SMTP(self.smtp_server)
        stmp.starttls() 
        stmp.login(self.maestro.get_credential("EMAIL_MANAGER", "OFFICE365_LOGIN"), 
                   self.maestro.get_credential("EMAIL_MANAGER", "OFFICE365_PWD"))
        
        # Envio do e-mail.
        stmp.sendmail(message["From"], recipients, message.as_string().encode("utf-8"))
        self.log.info("Relatório enviado por e-mail.")

        stmp.quit()

    def send_complete_report(self):
        """Envia o relatório da execução para a área."""
        recipients = ["nfebsb@gruposaga.com.br",
                      "gerentefaturamentobsb@gruposaga.com.br",
                      "caio.gmandarino@gruposaga.com.br",
                      "caroline.fmessias@gruposaga.com.br",
                      "jhennyfer.smoreira@gruposaga.com.br",
                      "nfekiabsb@gruposaga.com.br"]
        
        cc = ["cesar.asilva@gruposaga.com.br",
              "marcos.scurvina@gruposaga.com.br",
              "maykon.ponce@gruposaga.com.br",
              "rafaell.pcandido@gruposaga.com.br",
              "walisson.mnogueira@gruposaga.com.br"]

        # Remetente, destinatário e Assunto do e-mail.
        message = MIMEMultipart()
        message["From"] = "Saga RPA <saga.rpa@gruposaga.com.br>"
        message["To"] = ", ".join(recipients)
        message["Cc"] = ", ".join(cc)
        message["Subject"] = f"RPA - Importação de XML - BYD"

        # Mensagem (HTML).
        html_message = f"""
            <div>
                <p>Olá!</p>
                <p>Segue em anexo a planilha com os resultados da importação dos arquivos de veículos da BYD.</p>
                <p>Data de execução: {dt.date.today().strftime("%d/%m/%Y")}.</p>
            </div>
            """
        html_part = MIMEText(html_message, "html")
        message.attach(html_part)

        # Anexa a planilha de resultados.
        attachment = open(RESULTS_WORKBOOK_PATH, "rb")
        attachment_package = MIMEBase("application", "octet-stream")
        attachment_package.set_payload(attachment.read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header("content-disposition", "attachment", filename=f"{RESULTS_WORKBOOK_NAME}.xlsx")
        message.attach(attachment_package)

        # Login.
        stmp = smtplib.SMTP(self.smtp_server)
        stmp.starttls() 
        stmp.login(self.maestro.get_credential("EMAIL_MANAGER", "OFFICE365_LOGIN"), 
                   self.maestro.get_credential("EMAIL_MANAGER", "OFFICE365_PWD"))
        
        # Envio do e-mail.
        stmp.sendmail(message["From"], recipients + cc, message.as_string().encode("utf-8"))
        self.log.info("Relatório enviado por e-mail.")

        stmp.quit()