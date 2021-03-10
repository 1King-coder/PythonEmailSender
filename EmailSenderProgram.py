from sys import argv
from SendRoot import SendEmails
from Interface import *
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt5.QtGui import QPixmap
import smtplib
from DataBase import Data
import os
from pathlib import Path


# Application class.
class EmailSenderProgram(QMainWindow, Ui_MainWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.setFixedSize(1280, 720)
        self.root = None
        self.is_image = False
        self.template = False
        self.attached = False
        self.connected = False
        self.addsrs_emails = []
        self.setWindowTitle('Enviador de Emails do Vitor')
        self.Entrar.clicked.connect(self.login)
        self.Conectar.clicked.connect(self.connect)
        self.BuscaAnexo.clicked.connect(self.walk_archives)
        self.EnviaEmail.clicked.connect(self.send)
        self.Anexar.clicked.connect(self.attach_attachments)
        self.AdicionarEmail.clicked.connect(self.add_email)
        self.Delete.clicked.connect(self.delete_item_list)

    def login(self):
        return self.verify_email_password()

    def take_user(self):
        """
        Take sender's E-mail.
        """

        self.email = self.Email.displayText().strip()
        return self.email

    def take_password(self):
        """
        Take sender's password.
        """
        self.password = self.Senha.text().strip()
        return self.password

    def check_user_server(self):
        """
        Check sender's E-mail server.
        """

        global server
        user_email = self.take_user()
        if 'hotmail' in user_email or 'outlook' in user_email:
            server = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)

        elif 'yahoo' in user_email:
            server = smtplib.SMTP(host='smtp.mail.yahoo.com', port=587)

        else:
            server = smtplib.SMTP(host='smtp.gmail.com', port=587)

        server.ehlo()
        server.starttls()

    def take_email(self):
        """
        Return addresseer E-mail
        from DataBase by it's keyword.
        """

        self.kword = self.InsereMateria.displayText().strip()

        # Change 'directory_root' to your directory root.
        self.db = Data.AddressersData(
            Path('../Email Sender Program/DataBase/PeopleData.db'),
            'PeopleEmails'
        )

        self.addresser_email = self.db.people_data(self.kword)
        self.db.close()
        return self.addresser_email

    def take_sender_name(self):
        """
        Return the sender's name.
        """

        self.name = self.Nome.displayText().strip()
        return self.name

    def take_attachment_name(self):
        """
        Return specified attachment file name.
        """

        if not self.Anexo.displayText() == '':
            self.archive_name = self.Anexo.displayText().strip()
            return self.archive_name
        else:
            self.Response.setStyleSheet('*{color: orange;}')
            self.Response.setText('Selecione um arquivo.')
            return None

    def add_email(self):
        """
        Add addresseer E-mail to QListWidget.
        """

        email = self.take_email()
        try:
            if '@' in email:
                if email not in self.addsrs_emails:  # Check if the E-mail
                    self.PessoasLista.addItem(email)  # is already in the list.
                    self.addsrs_emails.append(email)
                    self.Response.setStyleSheet('*{color: green;}')
                    self.Response.setText(
                        'E-mail adiciodado a lista com sucesso!')
                else:
                    self.Response.setStyleSheet('*{color: orange;}')
                    self.Response.setText('E-mail já adicionado!')
            else:
                self.Response.setStyleSheet('*{color: orange;}')
                self.Response.setText('Insira uma keyword válida.')
        except Exception:
            self.Response.setStyleSheet('*{color: orange;}')
            self.Response.setText('Insira uma keyword válida.')

    def delete_item_list(self):
        """
        Delete the selected addresseer from list.
        """

        item = self.PessoasLista.currentItem()
        if item:
            if item.text() in self.addsrs_emails:

                # Hide selected E-mai from QListWidget.
                item.setHidden(True)

                # Delete selected E-mail from addressers.
                self.addsrs_emails.remove(item.text())

                self.DelResponse.setText('')
            else:
                self.DelResponse.setStyleSheet('*{color: orange;}')
                self.DelResponse.setText('Nada a deletar.')
        else:
            self.DelResponse.setStyleSheet('*{color: orange;}')
            self.DelResponse.setText('Nada a deletar.')

    def verify_email_password(self):
        """ 
        Check Login fields.
        """
        email = self.take_user()
        password = self.take_password()
        if email == '' and password == '':
            self.Error.setStyleSheet('*{color: red;}')
            self.Error.setText('Campos Email e Senha vazios')

        elif email == '':
            self.Error.setStyleSheet('*{color: red;}')
            self.Error.setText('Campo Email vazio!')

        elif password == '':
            self.Error.setStyleSheet('*{color: red;}')
            self.Error.setText('Campo Senha vazio!')

        else:
            self.Conectar.setChecked(True)
            return self.connect()

    def connect(self):
        """
        Login into the server
        """

        if self.Conectar.isChecked():

            try:
                self.check_user_server()
                server.login(self.take_user(), self.take_password())
            except Exception:
                self.Error.setStyleSheet('*{color: red;}')
                self.Error.setText('Email ou Senha inválidos')
                return
            self.connected = True
            self.Error.setStyleSheet('*{color: green;}')
            self.Error.setText('Conectado!')
            self.Conectar.setStyleSheet('*{color: green;}')

        else:  # Disconnect from the server.
            server.quit()
            self.Error.setText('')
            self.connected = False
            self.Conectar.setStyleSheet('*{color: red;}')

    def open_attachment(self):
        """
        Open File dialog.
        """
        if not self.attached:
            self.archive_name, _ = QFileDialog.getOpenFileName(
                self.centralwidget,
                'Abrir arquivo',
                'C:'
            )

            # Saves the file path.
            self.archive_path = self.archive_name

            # Get file's root.
            self.root, _ = os.path.split(self.archive_path)

            # Remove file's extension.
            self.archive_name, _ = os.path.splitext(self.archive_name)

            # Remove file's root, leaving only it's name.
            self.archive_name = self.archive_name.replace(
                f'{self.root}/', '')
            self.Anexo.setText(self.archive_name.strip())

    def open_image(self):
        """
        Return image file path
        """

        return self.archive_path

    def verify_if_image(self, ext):
        """
        Check if it's a image.
        """

        if ext in ['.jpg', '.png']:  # If it's a image
            self.is_image = True
            self.image = QPixmap(self.open_image())
            self.Response.setPixmap(self.image)  # Shows it.
            return True
        else:
            return False

    def walk_archives(self):
        """
        Calls File dialog function.
        """
        if self.Anexo.displayText() == '' or not self.attached:
            self.open_attachment()

    def attach_data(self, data):  # Return attachment by SenderRoot Module.

        mime = SendEmails.anexo(data[0], data[1], data[2])
        return mime

    def first_attachment(self, data):
        """
        Create E-mail's template with the first attachment.
        """

        if not self.Anexo.displayText() == '' and self.connected:

            # Sets E-mail template.
            msg = SendEmails.set_template(
                Path('SendRoot/template.html'),
                self.take_attachment_name(),
                self.take_sender_name()
            )

            # Sets Message "templated"
            self.template = True

            # Attach attachment to Message.
            msg.attach(self.attach_data(data))

            self.Response.setText('')

            if self.verify_if_image(data[1]):
                self.verify_if_image(data[1])
            else:
                self.Response.setStyleSheet('*{color: green;}')
                self.Response.setText('Template com anexo criado!')
            return msg
        else:
            try:
                raise Exception
            except Exception:
                self.Response.setStyleSheet('*{color: red;}')
                self.Response.setText('Campo de anexo em branco!')
                return

    def attach_attachments(self):
        """
        Attach selected attachment to the Message.
        """
        try:
            global msg
            if not self.root:  # Check if user selected a file.
                self.Response.setStyleSheet('*{color: orange;}')
                self.Response.setText('Selecione um arquivo.')
                return
            for data in SendEmails.get_path(
                self.take_attachment_name(),
                self.root
            ):
                data_list = data

                # If it's the first attachment, set E-mail's template.
                if not self.template and self.connected:
                    msg = self.first_attachment(data_list)

                    return msg

                if not self.connected:  # Check if user is logged.

                    self.Response.setStyleSheet(
                        '*{color: orange; font-size: 50px}')
                    self.Response.setText('Conecte-se primeiro.')

                # If there's a template, just attach file to Message.
                else:
                    if not self.attached:
                        if self.verify_if_image(data_list[1]):
                            self.verify_if_image(data_list[1])
                        else:
                            self.Response.setStyleSheet('*{color: green;}')
                            self.Response.setText('Arquivo anexo criado!')
                        mime = SendEmails.anexo(
                            data_list[0], data_list[1], data_list[2])

                        msg.attach(mime)
                    return msg

        except Exception as e:
            self.Response.setStyleSheet('*{color: red; font-size: 50px;}')
            self.Response.setText(f'Error: {e}')

    def set_email(self):
        """
        Sets Message's structure.
        """
        msg = self.attach_attachments()
        return msg

    def send(self):
        """
        Sends Message to all addresseers in list.
        """
        try:
            self.attached = True
            message = self.set_email()
            if self.connected:  # Check if user is logged.
                try:
                    # Send Message/ E-mail to everyone in addressers list.
                    for addsr in self.addsrs_emails:
                        server.send_message(
                            message, from_addr=self.take_user(),
                            to_addrs=addsr
                        )
                    self.attached = False

                except Exception:
                    self.Response.setStyleSheet('*{color: red;}')
                    self.Response.setText('Mensagem nula!')
                    self.attached = False
                    return

                # Check if there's a addressers in list.
                if not self.addsrs_emails == []:
                    self.Response.setStyleSheet(
                        '*{color: green; font-size: 50px;}')
                    self.Response.setText('Email enviado com sucesso!')
                    self.attached = False
                else:
                    self.Response.setText('*{color: orange;}')
                    self.Response.setText('Adicione um e-mail primeiro!')
            else:
                self.Response.setStyleSheet(
                    '*{color: orange; font-size: 50px}')
                self.Response.setText('Conecte-se primeiro.')
                self.attached = False
        except Exception as e:
            self.Response.setStyleSheet('*{color: red; font-size: 50px;}')
            self.Response.setText(f'Error: {e}')


if __name__ == '__main__':
    qt = QApplication(argv)
    send = EmailSenderProgram()
    send.show()
    qt.exec_()
