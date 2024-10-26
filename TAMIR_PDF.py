import sys
import configparser
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QMessageBox, QComboBox)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import os
import smtplib
import hashlib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import logging
from config import *

# Set up logging
logging.basicConfig(filename='app.log', level=logging.INFO)

class EmailSender(QThread):
    email_sent = pyqtSignal()
    email_failed = pyqtSignal(str)

    def __init__(self, email_info):
        super().__init__()
        self.email_info = email_info

    def run(self):
        try:
            # Extract email information
            sender_email = self.email_info['from']
            sender_password = self.email_info['password']
            smtp_server = self.email_info['smtp_server']
            smtp_port = self.email_info['smtp_port']
            receiver_email = self.email_info['to']
            subject = self.email_info['subject']
            pdf_path = self.email_info['pdf_path']

            # Create email message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject
            body = "Teklifinizin PDF dosyası ektedir."
            msg.attach(MIMEText(body, 'plain'))

            # Attach PDF
            with open(pdf_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(pdf_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
                msg.attach(part)

            # Connect to the SMTP server and send the email
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)

            self.email_sent.emit()  # Emit success signal
        except Exception as e:
            self.email_failed.emit(str(e))  # Emit error signal

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        # Başlangıç ayarları
        self.setWindowTitle("Giriş")
        self.setGeometry(100, 100, 300, 200)
        self.setup_ui()
        self.load_credentials()

    def setup_ui(self):
        layout = QVBoxLayout()
        self.email_label = QLabel("E-posta adresi:")
        self.email_input = QLineEdit(self)
        self.password_label = QLabel("Şifre:")
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.submit_button = QPushButton("Giriş Yap")
        self.submit_button.clicked.connect(self.verify_credentials)
        layout.addWidget(self.email_label)
        layout.addWidget(self.email_input)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)

    def load_credentials(self):
        # Kaydedilen e-posta bilgilerini yükler
        config = configparser.ConfigParser()
        config.read("settings.ini")
        self.saved_email = config.get("CREDENTIALS", "email", fallback="")
        self.saved_password = config.get("CREDENTIALS", "password", fallback="")
    
        # Otomatik doldurma
        self.email_input.setText(self.saved_email)
        self.password_input.setText(self.saved_password)


    def verify_credentials(self):
        # Girilen şifreyi hash'e çevirip kaydedilen hash ile karşılaştırır
        entered_email = self.email_input.text()
        entered_password = self.password_input.text()

        if entered_email == self.saved_email and entered_password == self.saved_password:
            QMessageBox.information(self, "Başarılı", "Giriş başarılı.")
            self.open_main_application()
        else:
            QMessageBox.critical(self, "Hata", "E-posta veya şifre hatalı.")

    def open_main_application(self):
        # Ana uygulamayı açma
        self.main_window = RepairOfferApplication()
        self.main_window.show()
        self.close()  # Giriş penceresini kapat

class RepairOfferApplication(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Oluşturucu")
        self.setGeometry(100, 100, 850, 520)
        self.pdf_template_path = ""
        self.config_file = "settings.ini"
        self.setup_gui()
        self.load_pdf_template_path()

    def setup_gui(self):
        # Başlık
        title_label = QLabel("FAST TAMIR TEKLIF PROGRAMI", self)
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        title_label.setAlignment(Qt.AlignCenter)

        # PDF şablonu seçme butonu
        self.pdf_button = QPushButton("PDF Şablonunu Seç", self)
        self.pdf_button.clicked.connect(self.select_pdf_template)

        # Seçilen dosya etiketi
        self.label_selected_file = QLabel("Seçilen PDF: Henüz seçilmedi", self)
        self.label_selected_file.setStyleSheet("font-size: 14pt;")

        # Bayi seçimi
        self.dealer_label = QLabel("Kime:", self)
        self.dealer_label.setStyleSheet("font-size: 14pt;")
        self.dealer_combo = QComboBox(self)
        self.dealer_combo.addItems(bayiler_listesi)
        self.dealer_combo.setEnabled(False)
        self.dealer_combo.currentIndexChanged.connect(self.bayi_secildi)

        # Girdi kutuları
        self.entries = {}
        self.create_form()

        # PDF oluşturma butonu
        self.button_create_pdf = QPushButton("PDF Oluştur", self)
        self.button_create_pdf.setEnabled(False)
        self.button_create_pdf.clicked.connect(self.pdf_olustur)

        # E-posta gönderme butonu
        self.button_send_email = QPushButton("E-posta Gönder", self)
        self.button_send_email.setEnabled(False)
        self.button_send_email.clicked.connect(self.send_email_threaded)

        # Layout ayarları
        layout = QVBoxLayout()
        layout.addWidget(title_label)
        layout.addWidget(self.pdf_button)
        layout.addWidget(self.label_selected_file)
        layout.addWidget(self.dealer_label)
        layout.addWidget(self.dealer_combo)

        # Girdi alanlarını ekleyelim
        for label, entry in self.entries.items():
            layout.addWidget(QLabel(f"{label}:"))
            layout.addWidget(entry)

        layout.addWidget(self.button_create_pdf)
        layout.addWidget(self.button_send_email)
        self.setLayout(layout)

    def create_form(self):
        for label_text in ["Telefon", "Faks", "Adres", "e-Mail", "Model", "Seri No", "Tamir Fiyatı"]:
            entry = QLineEdit(self)
            entry.setEnabled(False)
            self.entries[label_text] = entry

    def select_pdf_template(self):
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getOpenFileName(self, "PDF Şablon Dosyasını Seçin", "", "PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)", options=options)
            if file_path:
                self.label_selected_file.setText(file_path)
                self.pdf_template_path = file_path
                self.activate_fields()
                self.save_pdf_template_path()

    def load_pdf_template_path(self):
        config = configparser.ConfigParser()
        # Dosyayı okuma denemesi
        try:
            config.read(self.config_file)
            if "SETTINGS" in config:
                # Ayarları okuma
                self.pdf_template_path = config.get("SETTINGS", "pdf_template_path")
                if self.pdf_template_path:
                    self.label_selected_file.setText(self.pdf_template_path)
                    self.activate_fields()
                    logging.info("PDF şablonu konfigürasyondan yüklendi.")
                else:
                    logging.warning("PDF şablon yolu boş. Konfigürasyondan yüklenemedi.")
            else:
                # Eğer 'SETTINGS' bölümü yoksa, yeni bir bölüm oluştur
                config["SETTINGS"] = {'pdf_template_path': ''}
                with open(self.config_file, 'w') as configfile:
                    config.write(configfile)
                    logging.info("Yeni 'SETTINGS' bölümü oluşturuldu.")
        except Exception as e:
            logging.error(f"Konfigürasyon dosyası okunurken hata oluştu: {str(e)}")

    def save_pdf_template_path(self):
        config = configparser.ConfigParser()
        config.read(self.config_file)

        if "SETTINGS" not in config:
            config["SETTINGS"] = {}
        
        config["SETTINGS"]["pdf_template_path"] = self.pdf_template_path

        # Dosyaya yazma denemesi
        try:
            with open(self.config_file, 'w') as configfile:
                config.write(configfile)
                logging.info("PDF şablon yolu config'e kaydedildi.")
        except Exception as e:
            logging.error(f"Konfigürasyon dosyası yazılırken hata oluştu: {str(e)}")


    def activate_fields(self):
        self.dealer_combo.setEnabled(True)
        for entry in self.entries.values():
            entry.setEnabled(True)
        self.button_create_pdf.setEnabled(True)

    def bayi_secildi(self):
        bayi = self.dealer_combo.currentText()
        self.populate_entries(bayi)

    def populate_entries(self, bayi: str):
        self.entries["Telefon"].setText(telefonlar.get(bayi, ""))
        self.entries["Faks"].setText(fakslar.get(bayi, ""))
        self.entries["Adres"].setText(adresler.get(bayi, ""))
        self.entries["e-Mail"].setText(email_adresleri.get(bayi, ""))
        self.entries["Model"].setText("")  # Model ve Seri No'yu temizle
        self.entries["Seri No"].setText("")
        self.button_send_email.setEnabled(False)

    def pdf_olustur(self):
        if not self.pdf_template_path:
            QMessageBox.critical(self, "Hata", "Lütfen önce bir PDF şablonu seçin!")
            return

        fields = [self.dealer_combo.currentText()] + [entry.text() for entry in self.entries.values()]
        if not all(fields):
            QMessageBox.critical(self, "Hata", "Lütfen tüm alanları doldurun!")
            return

        output_pdf_path = os.path.join(os.path.dirname(self.pdf_template_path), f"{self.entries['Seri No'].text()}.pdf")
        pdf_reader = PdfReader(self.pdf_template_path)
        pdf_writer = PdfWriter()

        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)

        self.write_to_pdf(c)
        c.save()

        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = pdf_reader.pages[0]
        page.merge_page(new_pdf.pages[0])
        pdf_writer.add_page(page)

        with open(output_pdf_path, "wb") as outputStream:
            pdf_writer.write(outputStream)

        QMessageBox.information(self, "Başarılı", f"{self.entries['Seri No'].text()} seri numaralı teklif oluşturuldu.")
        logging.info(f"{self.entries['Seri No'].text()} seri numaralı teklif oluşturuldu.")
        self.button_send_email.setEnabled(True)

    def write_to_pdf(self, canvas):
        canvas.drawString(102, 571, self.dealer_combo.currentText())
        canvas.drawString(102, 555, self.entries['Telefon'].text())
        canvas.drawString(102, 539, self.entries['Faks'].text())
        canvas.drawString(102, 525, self.entries['Adres'].text())
        canvas.drawString(102, 343, self.entries['Model'].text())
        canvas.drawString(320, 343, self.entries['Seri No'].text())
        canvas.drawString(480, 343, f"{self.entries['Tamir Fiyatı'].text()} TL + KDV")

    def load_password(self):
        # Kaydedilen şifreyi yükler
        config = configparser.ConfigParser()
        config.read("settings.ini")
        return config.get("CREDENTIALS", "password", fallback="")

    def send_email_threaded(self):
        email_info = {
            'from': self.entries['e-Mail'].text(),
            'to': self.entries['e-Mail'].text(),
            'subject': f"Teklif - {self.entries['Seri No'].text()}",
            'pdf_path': os.path.join(os.path.dirname(self.pdf_template_path), f"{self.entries['Seri No'].text()}.pdf"),
            'smtp_server': "proxy.uzmanposta.com",
            'smtp_port': 465,  # Port 465 SSL için
            'password': self.load_password()
        }

        self.email_sender = EmailSender(email_info)
        self.email_sender.email_sent.connect(self.on_email_sent)
        self.email_sender.email_failed.connect(self.on_email_failed)
        self.email_sender.start()  # E-posta gönderme iş parçacığını başlat

    def on_email_sent(self):
        QMessageBox.information(self, "Başarılı", "E-posta başarıyla gönderildi.")

    def on_email_failed(self, error_message):
        QMessageBox.critical(self, "Hata", f"E-posta gönderiminde hata: {error_message}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_window = LoginWindow()
    login_window.show()
    sys.exit(app.exec_())