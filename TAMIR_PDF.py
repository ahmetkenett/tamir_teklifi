import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
import os
import smtplib
import threading
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import logging
import re
from config import *

# Set up logging
logging.basicConfig(filename='app.log', level=logging.INFO)

class TamirTeklifUygulamasi:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Oluşturucu")
        self.root.geometry("850x520")
        self.pdf_template_path = ""
        self.config_file = "settings.ini"
        self.setup_gui()
        self.load_pdf_template_path()

    def setup_gui(self):
        title_label = tk.Label(self.root, text="FAST TAMIR TEKLIF PROGRAMI", font=("Helvetica", 18, "bold"), bg="#f0f0f0", fg="#000000")
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        tk.Button(self.root, text="PDF Şablonunu Seç", font=("Helvetica", 16, "bold"), command=self.select_pdf_template, bg="#4CAF50", fg="white").grid(row=1, column=0, padx=10, pady=10)
        self.label_selected_file = tk.Label(self.root, text="Seçilen PDF: Henüz seçilmedi", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
        self.label_selected_file.grid(row=1, column=1, sticky="w")

        self.create_form()

        self.button_create_pdf = tk.Button(self.root, text="PDF Oluştur", font=("Helvetica", 14, "bold"), command=self.pdf_olustur, state="disabled", bg="#2196F3", fg="white")
        self.button_create_pdf.grid(row=8, column=1, padx=300, pady=5, sticky="w")

        self.button_send_email = tk.Button(self.root, text="E-posta Gönder", font=("Helvetica", 14, "bold"), command=self.send_email_threaded, state="disabled", bg="#2196F3", fg="white")
        self.button_send_email.grid(row=9, column=1, padx=300, pady=5, sticky="w")

    def create_form(self):
        kime_var = tk.StringVar()
        self.dropdown_kime = ttk.Combobox(self.root, textvariable=kime_var, values=bayiler_listesi, state="disabled", font=("Helvetica", 14), width=50)
        self.dropdown_kime.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        self.dropdown_kime.bind("<<ComboboxSelected>>", self.bayi_secildi)

        tk.Label(self.root, text="Kime:", font=("Arial", 15, "bold"), bg="#f0f0f0", fg="black").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        tk.Label(self.root, text="by Ahmet Erdem Kenet", font=("Arial", 8, "bold"), bg="#f0f0f0", fg="black").grid(row=10, column=0, padx=5, pady=5, sticky="e")

        self.entries = {}
        for label_text in ["Telefon", "Faks", "Adres", "e-Mail", "Model", "Seri No", "Tamir Fiyatı"]:
            self.create_entry(label_text)

    def create_entry(self, label_text):
        entry = tk.Entry(self.root, font=("Helvetica", 14), width=20 if label_text != "Tamir Fiyatı" else 6, state="disabled")
        self.entries[label_text] = entry

        row = len(self.entries) + 2
        tk.Label(self.root, text=f"{label_text}:", font=("Arial", 15, "bold"), bg="#f0f0f0").grid(row=row, column=0, padx=5, pady=8, sticky="e")
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="w")

    def select_pdf_template(self):
        file_path = filedialog.askopenfilename(title="PDF Şablon Dosyasını Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if file_path:
            self.label_selected_file.config(text=file_path, foreground="green")
            self.pdf_template_path = file_path
            self.activate_fields()
            self.save_pdf_template_path()

    def load_pdf_template_path(self):
        config = configparser.ConfigParser()
        if not os.path.exists(self.config_file):
            config['SETTINGS'] = {'pdf_template_path': ''}
            with open(self.config_file, 'w') as configfile:
                config.write(configfile)

        try:
            config.read(self.config_file)
            self.pdf_template_path = config.get("SETTINGS", "pdf_template_path", fallback="")
            if self.pdf_template_path:
                self.label_selected_file.config(text=self.pdf_template_path, foreground="green")
                self.activate_fields()
                logging.info("Loaded PDF template path from config.")
        except configparser.Error as e:
            messagebox.showerror("Hata", f"Config dosyası okunamadı: {str(e)}")

    def save_pdf_template_path(self):
        config = configparser.ConfigParser()
        config.read(self.config_file)
        if 'SETTINGS' not in config:
            config['SETTINGS'] = {}

        config['SETTINGS']['pdf_template_path'] = self.pdf_template_path
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)
            logging.info("Saved PDF template path to config.")

    def activate_fields(self):
        self.dropdown_kime.config(state="readonly")
        for entry in self.entries.values():
            entry.config(state="normal")
        self.button_create_pdf.config(state="normal")

    def bayi_secildi(self, event):
        self.populate_entries(self.dropdown_kime.get())

    def populate_entries(self, bayi: str):
        self.entries["Telefon"].delete(0, tk.END)
        self.entries["Telefon"].insert(0, telefonlar.get(bayi, ""))
        self.entries["Faks"].delete(0, tk.END)
        self.entries["Faks"].insert(0, fakslar.get(bayi, ""))
        self.entries["Adres"].delete(0, tk.END)
        self.entries["Adres"].insert(0, adresler.get(bayi, ""))
        self.entries["e-Mail"].delete(0, tk.END)
        self.entries["e-Mail"].insert(0, email_adresleri.get(bayi, ""))

    def pdf_olustur(self):
        if not self.pdf_template_path:
            messagebox.showerror("Hata", "Lütfen önce bir PDF şablonu seçin!")
            return

        fields = [self.dropdown_kime.get()] + [entry.get() for entry in self.entries.values()]
        if not all(fields):
            messagebox.showerror("Hata", "Lütfen tüm alanları doldurun!")
            return

        output_pdf_path = os.path.join(os.path.dirname(self.pdf_template_path), f"{self.entries['Seri No'].get()}.pdf")
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

        try:
            with open(output_pdf_path, "wb") as output_pdf_file:
                pdf_writer.write(output_pdf_file)
            messagebox.showinfo("Başarılı", f"{self.entries['Seri No'].get()} seri numaralı teklif oluşturuldu!")
            self.button_send_email.config(state="normal")
            logging.info(f"PDF created: {output_pdf_path}")
        except Exception as e:
            messagebox.showerror("Hata", f"PDF oluşturulurken hata: {str(e)}")
            logging.error(f"PDF creation error: {e}")

    def write_to_pdf(self, canvas):
        canvas.drawString(102, 571, self.dropdown_kime.get())
        canvas.drawString(102, 555, self.entries["Telefon"].get())
        canvas.drawString(102, 539, self.entries["Faks"].get())
        canvas.drawString(102, 525, self.entries["Adres"].get())
        canvas.drawString(102, 343, self.entries["Model"].get())
        canvas.drawString(320, 343, self.entries["Seri No"].get())
        canvas.drawString(480, 343, f"{self.entries['Tamir Fiyatı'].get()} TL + KDV")

    def is_valid_email(self, email):
        regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
        valid = re.match(regex, email)
        if not valid:
            logging.warning(f"Email adresi hatalı!: {email}")
        return valid

    def send_email(self):
        secilen_bayi = self.dropdown_kime.get()
        pdf_path = os.path.join(os.path.dirname(self.pdf_template_path), f"{self.entries['Seri No'].get()}.pdf")
        email_address = self.entries["e-Mail"].get() or email_adresleri.get(secilen_bayi)

        if not os.path.exists(pdf_path):
            messagebox.showerror("Hata", "Lütfen önce PDF oluşturun.")
            return

        if not self.is_valid_email(email_address):
            messagebox.showerror("Hata", "Geçersiz e-posta adresi.")
            return

        subject = "TAMİR TEKLİFİ"
        body = "Merhaba,\n\nLütfen ekteki tamir teklifini inceleyin.\n\nSaygılarımızla,\n"

        sender_email = EMAIL["sender_email"]
        sender_password = EMAIL["sender_password"]
        smtp_server = EMAIL["smtp_server"]
        smtp_port = EMAIL["smtp_port"]

        try:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = email_address
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            with open(pdf_path, "rb") as f:
                attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                msg.attach(attach)

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email_address, msg.as_string())

            messagebox.showinfo("Başarılı", "{} için e-posta başarıyla gönderildi!".format(secilen_bayi))
            logging.info(f"Email sent to {email_address}")
        except Exception as e:
            messagebox.showerror("Hata", "E-posta gönderimi sırasında bir hata oluştu: {}".format(str(e)))
            logging.error(f"Email sending error: {e}")

    def send_email_threaded(self):
        threading.Thread(target=self.send_email).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = TamirTeklifUygulamasi(root)
    root.mainloop()