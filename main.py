import smtplib
from tkinter import *
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import schedule
import threading
import time
from datetime import datetime

#7enviar o e-mail
def send_email():
    try:
        sender_email = email_entry.get()
        password = "rlod jswu bmgi uzfv"
        recipient_email = recipient_entry.get()
        subject = subject_entry.get()
        body = body_entry.get("1.0", END)
        file_path = file_path_var.get()

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        if file_path:
            part = MIMEBase('application', 'octet-stream')
            with open(file_path, "rb") as attachment:
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
            msg.attach(part)

        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        
        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar o e-mail: {str(e)}")

# agendar email
def agendar_email():
    date = date_entry.get()  #
    hour = hour_entry.get()  

    if not date or not hour:
        messagebox.showerror("Erro", "Defina uma data e horário para o envio!")
        return

    try:
        data_hora = datetime.strptime(f"{date} {hour}", "%d/%m/%Y %H:%M")  # Convertendo para datetime
    except ValueError:
        messagebox.showerror("Erro", "Formato de data ou hora inválido! Use dd/mm/aaaa HH:MM")
        return

    #verifiçao se chegou a data/hora certa para enviar
    def verificar_agendamento():
        while True:
            if datetime.now() >= data_hora:
                send_email()
                break
            time.sleep(30)  # Verifica a cada 30 segundos

    threading.Thread(target=verificar_agendamento, daemon=True).start()
    messagebox.showinfo("Agendado", f"E-mail será enviado em {date} às {hour}")

#selecionar amultiplos arquivos
def selecionar_arquivo():
    file_path = filedialog.askopenfilename(filetypes=[("Todos os arquivos", "*.*")])
    file_path_var.set(file_path)

# iniciar o GUI
root = Tk()
root.title("Envio de E-mail")

#Layout
Label(root, text="Seu E-mail:").grid(row=0, column=0, padx=10, pady=5)
email_entry = Entry(root, width=40)
email_entry.grid(row=0, column=1, padx=10, pady=5)

Label(root, text="E-mail do Destinatário:").grid(row=2, column=0, padx=10, pady=5)
recipient_entry = Entry(root, width=40)
recipient_entry.grid(row=2, column=1, padx=10, pady=5)

Label(root, text="Assunto:").grid(row=3, column=0, padx=10, pady=5)
subject_entry = Entry(root, width=40)
subject_entry.grid(row=3, column=1, padx=10, pady=5)

Label(root, text="Corpo do E-mail:").grid(row=4, column=0, padx=10, pady=5)
body_entry = Text(root, width=40, height=5)
body_entry.grid(row=4, column=1, padx=10, pady=5)

#botao do arquivo
Label(root, text="Selecionar Arquivo:").grid(row=5, column=0, padx=10, pady=5)
file_path_var = StringVar()
file_path_entry = Entry(root, textvariable=file_path_var, width=40, state="readonly")
file_path_entry.grid(row=5, column=1, padx=10, pady=5)
file_button = Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
file_button.grid(row=5, column=2, padx=10, pady=5)

#data
Label(root, text="Data do envio (dd/mm/aaaa):", font=("Arial", 10)).grid(row=6, column=0, columnspan=2, padx=10, pady=(10, 0))
date_entry = Entry(root, width=20)
date_entry.grid(row=7, column=0, columnspan=2, padx=10, pady=5)


Label(root, text="Hora do envio (HH:MM):", font=("Arial", 10)).grid(row=8, column=0, columnspan=2, padx=10, pady=(10, 0))
hour_entry = Entry(root, width=20)
hour_entry.grid(row=9, column=0, columnspan=2, padx=10, pady=5)


#enviar o e-mail imediatamente
send_button = Button(root, text="Enviar Agora", command=send_email)
send_button.grid(row=12, column=0, columnspan=2, pady=10)

#agendar o envio
schedule_button = Button(root, text="Agendar Envio", command=agendar_email)
schedule_button.grid(row=10, column=0, columnspan=2, pady=10)

# Iniciando a interface gráfica
root.mainloop()
