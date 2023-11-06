from PyQt5.QtWidgets import *
from PyQt5 import uic

import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

import imaplib
import email
i = 1
class interface(QMainWindow):

    def __init__(self):
        super(interface, self).__init__()
        uic.loadUi("emailClient.ui", self)
        self.show()
        

        

        self.LoginButton.clicked.connect(self.login)
        self.SendButton.clicked.connect(self.send_mail)
        self.toolButton.clicked.connect(self.attach_sth)
        self.toolButton_3.clicked.connect(self.anterior)
        self.toolButton_2.clicked.connect(self.proximo)
        self.toolButton.setEnabled(False)
        self.SendButton.setEnabled(False)

    def login(self):
        try:
            global i
            imap_server = "outlook.office365.com"
            imap = imaplib.IMAP4_SSL(imap_server)
            imap.login(self.EmailAdress.text(), self.EmailPassword.text())
            imap.select("Inbox")
            _, msgnums = imap.search(None, "ALL")
            msgnum = msgnums[0].split()[len(msgnums[0].split()) - i]
            _, data = imap.fetch(msgnum, "(RFC822)")
            message = email.message_from_bytes(data[0][1])
            geral = ""
            number = f"Message Number: {msgnum}"
            self.Sender.setText(f"Remetente: {message.get('From')}")
            BCC = f"BCC: {message.get('BCC')}"
            self.label_8.setText(f"Data: {message.get('Date')}")
            self.label_7.setText(f"Assunto: {message.get('Subject')}")
            for part in message.walk():
                if part.get_content_type() == "text/plain":
                    geral = geral + part.as_string()
            self.EmailBoxText.setText(number  + "\n" + BCC + "\n" + geral + "\n")
            imap.close
            self.toolButton_3.setEnabled(True)


            self.server = smtplib.SMTP("smtp-mail.outlook.com", 587)
            self.server.ehlo()
            self.server.starttls()
            self.server.ehlo()
            self.server.login(self.EmailAdress.text(), self.EmailPassword.text())
            self.EmailAdress.setEnabled(False)
            self.LoginButton.setEnabled(False)
            self.EmailPassword.setEnabled(False)
            self.ReceiverAdress.setEnabled(True)
            self.Subject.setEnabled(True)
            self.EmailText.setEnabled(True)
            self.SendButton.setEnabled(True)
            self.toolButton.setEnabled(True)
            self.msg = MIMEMultipart()

        except smtplib.SMTPAuthenticationError:
            message_box = QMessageBox()
            message_box.setText("Dados De login Invalidos!")
            message_box.exec()

        except:
            message_box = QMessageBox()
            message_box.setText("Erro No Login!")
            message_box.exec()

    
    def send_mail(self):
        dialog = QMessageBox()
        dialog.setText("Você quer enviar esse e-mail?")
        dialog.addButton(QPushButton("Sim"), QMessageBox.YesRole)
        dialog.addButton(QPushButton("Não"), QMessageBox.NoRole)

        if dialog.exec_() == 0:
            try:
                self.msg['From'] =  "teste"
                self.msg['To'] = self.ReceiverAdress.text()
                self.msg['Subject'] = self.Subject.text()
                self.msg.attach(MIMEText(self.EmailText.toPlainText(), 'plain'))
                text = self.msg.as_string()
                self.server.sendmail(self.EmailAdress.text(), self.ReceiverAdress.text(), text)
                message_box = QMessageBox()
                message_box.setText("E-mail enviado!")
                message_box.exec()
            
            except:
                message_box = QMessageBox()
                message_box.setText("O envio Falhou!")
                message_box.exec()
    
    def attach_sth(self):
        options = QFileDialog.Options()
        filenames, _ = QFileDialog.getOpenFileNames(self, "Open File", "", "All Files (*.*)", options = options)
        if filenames != []:
            for filename in filenames:
                attachment = open(filename, 'rb')

                filename = filename[filename.rfind("/") + 1:]

                p = MIMEBase('application', 'octet-stream')
                p.set_payload(attachment.read())
                encoders.encode_base64(p)
                p.add_header("Content-Disposition", f"attachment; filename= {filename}")
                self.msg.attach(p)
                if not self.label_5.text().endswith(":"):
                    self.label_5.setText(self.label_5.text() + ",")
                self.label_5.setText(self.label_5.text() + " " + filename)
    
    def anterior(self):
        global i
        i = i + 1
        imap_server = "outlook.office365.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(self.EmailAdress.text(), self.EmailPassword.text())
        imap.select("Inbox")
        _, msgnums = imap.search(None, "ALL")
        msgnum = msgnums[0].split()[len(msgnums[0].split()) - i]
        _, data = imap.fetch(msgnum, "(RFC822)")
        message = email.message_from_bytes(data[0][1])

        geral = ""
        number = f"Message Number: {msgnum}"
        self.Sender.setText(f"Remetente: {message.get('From')}")
        BCC = f"BCC: {message.get('BCC')}"
        self.label_8.setText(f"Data: {message.get('Date')}")
        self.label_7.setText(f"Assunto: {message.get('Subject')}")
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                geral = geral + part.as_string()
        self.EmailBoxText.setText(number + "\n" + BCC + "\n" + geral + "\n")
        imap.close
        if (len(msgnums[0].split()) - i - 1) < 0:
            self.toolButton_3.setEnabled(False)
        self.toolButton_2.setEnabled(True)
    
    def proximo(self):
        global i
        i = i - 1
        imap_server = "outlook.office365.com"
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(self.EmailAdress.text(), self.EmailPassword.text())
        imap.select("Inbox")
        _, msgnums = imap.search(None, "ALL")
        msgnum = msgnums[0].split()[len(msgnums[0].split()) - i]
        _, data = imap.fetch(msgnum, "(RFC822)")
        message = email.message_from_bytes(data[0][1])

        geral = ""
        number = f"Message Number: {msgnum}"
        self.Sender.setText(f"Remetente: {message.get('From')}")
        BCC = f"BCC: {message.get('BCC')}"
        self.label_8.setText(f"Data: {message.get('Date')}")
        self.label_7.setText(f"Assunto: {message.get('Subject')}")
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                geral = geral + part.as_string()
        self.EmailBoxText.setText(number + "\n" + BCC + "\n" + geral + "\n")
        imap.close
        if (len(msgnums[0].split()) - i + 1) > (len(msgnums[0].split()) -1):
            self.toolButton_2.setEnabled(False)
        self.toolButton_3.setEnabled(True)


app = QApplication([])
window = interface()
app.exec_()
