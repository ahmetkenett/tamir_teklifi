import sys
import configparser
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QMessageBox, QComboBox
from PyQt5.QtCore import Qt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import logging
from config import *

# Set up logging
logging.basicConfig(filename='app.log', level=logging.INFO)

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Giriş")
        self.setGeometry(100, 100, 300, 200)
        self.setup_ui()
        self.load_credentials()

    def setup_ui(self):
        layout = QVBoxLayout()
        self.email_input = QLineEdit(self)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.submit_button = QPushButton("Giriş Yap")
        self.submit_button.clicked.connect(self.save_credentials)

        layout.addWidget(QLabel("E-posta adresi:"))
        layout.addWidget(self.email_input)
        layout.addWidget(QLabel("Şifre:"))
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

    def save_credentials(self):
        email, password = self.email_input.text(), self.password_input.text()
        if email and password:
            config = configparser.ConfigParser()
            config.read("settings.ini")
            config['CREDENTIALS'] = {
                'email': email,
                'password': password,
                'smtp_server': "proxy.uzmanposta.com",
                'smtp_port': "465"
            }
            with open("settings.ini", "w") as configfile:
                config.write(configfile)
            QMessageBox.information(self, "Başarılı", "Kimlik bilgileri başarıyla kaydedildi.")
            self.close()
            self.main_window = TamirTeklifUygulamasi(email, password)
            self.main_window.show()
        else:
            QMessageBox.critical(self, "Hata", "Lütfen tüm alanları doldurun.")

class TamirTeklifUygulamasi(QWidget):
    def __init__(self, email, password):
        super().__init__()
        self.email, self.password = email, password
        self.setWindowTitle("PDF Oluşturucu")
        self.setGeometry(100, 100, 850, 520)
        self.config_file = "settings.ini"
        self.setup_gui()
        self.load_pdf_template_path()

    def setup_gui(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("FAST TAMIR TEKLIF PROGRAMI", self, alignment=Qt.AlignCenter, styleSheet="font-size: 18pt; font-weight: bold;"))

        self.pdf_button = QPushButton("PDF Şablonunu Seç", self)
        self.pdf_button.clicked.connect(self.select_pdf_template)
        self.label_selected_file = QLabel("Seçilen PDF: Henüz seçilmedi", self)
        self.dealer_combo = QComboBox(self)
        self.dealer_combo.addItems(bayiler_listesi)
        self.dealer_combo.currentIndexChanged.connect(self.bayi_secildi)

        layout.addWidget(self.pdf_button)
        layout.addWidget(self.label_selected_file)
        layout.addWidget(QLabel("Kime:", self))
        layout.addWidget(self.dealer_combo)

        self.entries = {label: QLineEdit(self, enabled=False) for label in ["Telefon", "Faks", "Adres", "e-Mail", "Model", "Seri No", "Tamir Fiyatı"]}
        for label, entry in self.entries.items():
            layout.addWidget(QLabel(f"{label}:"))
            layout.addWidget(entry)

        self.button_create_pdf = QPushButton("PDF Oluştur", self, enabled=False)
        self.button_create_pdf.clicked.connect(self.pdf_olustur)
        self.button_send_email = QPushButton("E-posta Gönder", self, enabled=False)
        self.button_send_email.clicked.connect(self.send_email)
        
        layout.addWidget(self.button_create_pdf)
        layout.addWidget(self.button_send_email)
        self.setLayout(layout)

    def select_pdf_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "PDF Şablon Dosyasını Seçin", "", "PDF Dosyaları (*.pdf);;Tüm Dosyalar (*)")
        if file_path:
            self.label_selected_file.setText(file_path)
            self.pdf_template_path = file_path
            self.activate_fields()
            self.save_pdf_template_path()

    def load_pdf_template_path(self):
        config = configparser.ConfigParser()
        config.read(self.config_file)
        self.pdf_template_path = config.get("SETTINGS", "pdf_template_path", fallback="")
        if self.pdf_template_path:
            self.label_selected_file.setText(self.pdf_template_path)
            self.activate_fields()
            logging.info("PDF şablonu konfigürasyondan yüklendi.")

    def save_pdf_template_path(self):
        config = configparser.ConfigParser()
        config.read(self.config_file)
        config['SETTINGS'] = {'pdf_template_path': self.pdf_template_path}
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)
            logging.info("PDF şablon yolu config'e kaydedildi.")

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

    def pdf_olustur(self):
        if not self.pdf_template_path:
            QMessageBox.critical(self, "Hata", "Lütfen önce bir PDF şablonu seçin!")
            return

        fields = [self.dealer_combo.currentText()] + [entry.text() for entry in self.entries.values()]
        if not all(fields):
            QMessageBox.critical(self, "Hata", "Lütfen tüm alanları doldurun!")
            return

        output_pdf_path = os.path.join(os.path.dirname(self.pdf_template_path), f"{self.entries['Seri No'].text()}.pdf")
        pdf_reader, pdf_writer = PdfReader(self.pdf_template_path), PdfWriter()

        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        self.write_to_pdf(c)
        c.save()

        packet.seek(0)
        page = pdf_reader.pages[0]
        page.merge_page(PdfReader(packet).pages[0])
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

    def send_email(self):
        try:
            msg = MIMEMultipart()
            msg['From'], msg['To'], msg['Subject'] = self.email, self.entries["e-Mail"].text(), f"{self.entries['Seri No'].text()} Seri Numaralı Cihazın Tamir Teklifi"
            msg.attach(MIMEText("Teklif ektedir.\n\nSaygılarımızla.", 'plain'))

            pdf_filename = f"{self.entries['Seri No'].text()}.pdf"
            with open(pdf_filename, "rb") as file:
                part = MIMEApplication(file.read(), Name=os.path.basename(pdf_filename))
                part['Content-Disposition'] = f'attachment; filename="{pdf_filename}"'
                msg.attach(part)

            with smtplib.SMTP_SSL("proxy.uzmanposta.com", 465) as server:
                server.login(self.email, self.password)
                server.send_message(msg)

            QMessageBox.information(self, "Başarılı", "E-posta başarıyla gönderildi.")
            logging.info(f"{self.entries['Seri No'].text()} seri numaralı teklif e-posta ile gönderildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"E-posta gönderimi sırasında bir hata oluştu: {str(e)}")
            logging.error(f"E-posta gönderimi hatası: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_window = LoginWindow()
    login_window.show()
    sys.exit(app.exec_())