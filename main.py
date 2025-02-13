import smtplib
from tkinter import *
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import threading
import time
import json
from datetime import datetime
import os


email_senhas = {}

EMAILS_FILE = "emails.json"

if not os.path.exists(EMAILS_FILE):
    with open(EMAILS_FILE, "w") as file:
        json.dump({}, file)  # cria um JSON vazio

# função para selecionar um e-mail salvo e preencher os campos
def selecionar_email(event):
    try:
        selected_index = email_listbox.curselection()[0]
        email = email_listbox.get(selected_index)

        email_entry.delete(0, END)
        email_entry.insert(0, email)

        # preenche automaticamente a senha correspondente
        if email in email_senhas:
            password_entry.delete(0, END)
            password_entry.insert(0, email_senhas[email])
        else:
            password_entry.delete(0, END)

    except IndexError:
        pass


def remover_email():
    try:
        selected_index = email_listbox.curselection()[0]
        email = email_listbox.get(selected_index)

        if email in email_senhas:
            del email_senhas[email]  # remove do dicionário
            email_listbox.delete(selected_index)  # remove da listbox
            salvar_emails()  # atualiza o arquivo JSON
            messagebox.showinfo("Sucesso", f"E-mail {email} removido com sucesso!")
        
        email_entry.delete(0, END)
        password_entry.delete(0, END)

    except IndexError:
        messagebox.showerror("Erro", "Selecione um e-mail para remover!")

# função para enviar e-mail
def send_email():
    try:
        # pega os dados inseridos no tkinter
        sender_email = email_entry.get().strip()
        password = password_entry.get().strip()
        recipient_email = recipient_entry.get().strip()
        subject = subject_entry.get()
        body = body_entry.get("1.0", END)
        
        # verifica se ha senha
        if not password:
            messagebox.showerror("Erro", "Senha não encontrada para este e-mail!")
            return

        # enviar o e-mail
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # anexar os arquivos
        for file_path in file_paths:
            part = MIMEBase('application', 'octet-stream')
            with open(file_path, "rb") as attachment:
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)

        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()

        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar o e-mail: {str(e)}")

# agendar envio de e-mail com data e hora
def agendar_envio_unico():
    data = data_entry.get().strip()
    hora = hora_entry.get().strip()
    try:
        envio_datetime = datetime.strptime(f"{data} {hora}", '%d/%m/%Y %H:%M')
        agora = datetime.now() 
        segundos_ate_envio = (envio_datetime - agora).total_seconds()

        if segundos_ate_envio <= 0:
            messagebox.showerror("Erro", "A data e hora escolhidas já passaram. Escolha um momento futuro.")
            return

        messagebox.showinfo("Agendado", f"E-mail será enviado em {data} às {hora}.")
        threading.Timer(segundos_ate_envio, lambda: (send_email(), data_entry.delete(0, END), hora_entry.delete(0, END))).start()
    except ValueError:
        messagebox.showerror("Erro", "Formato inválido! Use DD/MM/AAAA para data e HH:MM para hora.")

# adicionar um novo e-mail à lista manualmente
def adicionar_email():
    novo_email = email_entry.get().strip()
    nova_senha = password_entry.get().strip()

    if novo_email and nova_senha:
        if novo_email not in email_senhas:
            email_senhas[novo_email] = nova_senha
            email_listbox.insert(END, novo_email)
            salvar_emails()  # salva os dados no arquivo
            messagebox.showinfo("Sucesso", f"E-mail {novo_email} salvo com sucesso!")
        else:
            messagebox.showerror("Erro", "E-mail já está salvo!")
    else:
        messagebox.showerror("Erro", "Preencha o e-mail e a senha!")

# função para carregar e-mails no arquivo JSON
def carregar_emails():
    global email_senhas
    try:
        with open(EMAILS_FILE, "r") as file:
            content = file.read()
            if content.strip():  # verifica se o arquivo não está vazio
                email_senhas = json.loads(content)
            else:
                email_senhas = {}  
         
            email_listbox.delete(0, END)
            for email in email_senhas.keys():
                email_listbox.insert(END, email)
    except FileNotFoundError:
        email_senhas = {}  # Se o arquivo não existir, inicializa com um dicionário vazio
    except json.JSONDecodeError:
        messagebox.showerror("Erro", "O arquivo de e-mails está corrompido. Inicializando com um dicionário vazio.")
        email_senhas = {}  # inicializa tudo do 0 de o json tiver corrompido

# função para salvar e-mails no arquivo JSON
def salvar_emails():
    with open(EMAILS_FILE, "w") as file:
        json.dump(email_senhas, file)

# lista para armazenar os caminhos dos arquivos anexados
file_paths = []

# função para adicionar um arquivo à lista de anexos
def adicionar_arquivo():
    arquivo = filedialog.askopenfilename()
    if arquivo:
        file_paths.append(arquivo)
        atualizar_lista_anexos()

# função para remover um arquivo da lista de anexos
def remover_arquivo():
    try:
        selected_index = attachment_listbox.curselection()[0]
        file_paths.pop(selected_index)  # Remove o arquivo da lista
        atualizar_lista_anexos()  # Atualiza a listbox
    except IndexError:
        messagebox.showerror("Erro", "Selecione um arquivo para remover!")

# função para atualizar a lista de anexos na interface
def atualizar_lista_anexos():
    attachment_listbox.delete(0, END)
    for file_path in file_paths:
        attachment_listbox.insert(END, os.path.basename(file_path))  # Exibe apenas o nome do arquivo

# criar a interface gráfica
root = Tk()
root.title("Envio de E-mail")

# lista de e-mails salvos
Label(root, text="E-mails Salvos:").grid(row=0, column=0, padx=10, pady=5)
email_listbox = Listbox(root, height=5)
email_listbox.grid(row=0, column=1, padx=10, pady=5)
for email in email_senhas.keys():
    email_listbox.insert(END, email)

email_listbox.bind("<<ListboxSelect>>", selecionar_email)

# entrada para novo e-mail
Label(root, text="Novo E-mail:").grid(row=1, column=0, padx=10, pady=5)
email_entry = Entry(root, width=40)
email_entry.grid(row=1, column=1, padx=10, pady=5)

# entrada para senha
Label(root, text="Senha:").grid(row=2, column=0, padx=10, pady=5)
password_entry = Entry(root, width=40, show="*")  # Esconde a senha com '*'
password_entry.grid(row=2, column=1, padx=10, pady=5)


add_button = Button(root, text="Salvar E-mail", command=adicionar_email)
add_button.grid(row=3, column=0, pady=10)

remove_button = Button(root, text="Remover E-mail", command=remover_email)
remove_button.grid(row=3, column=1, pady=10)


Label(root, text="E-mail do Destinatário:").grid(row=4, column=0, padx=10, pady=5)
recipient_entry = Entry(root, width=40)
recipient_entry.grid(row=4, column=1, padx=10, pady=5)


Label(root, text="Assunto:").grid(row=5, column=0, padx=10, pady=5)
subject_entry = Entry(root, width=40)
subject_entry.grid(row=5, column=1, padx=10, pady=5)


Label(root, text="Corpo do E-mail:").grid(row=6, column=0, padx=10, pady=5)
body_entry = Text(root, width=40, height=5)
body_entry.grid(row=6, column=1, padx=10, pady=5)


Label(root, text="Selecionar Arquivo:").grid(row=7, column=0, padx=10, pady=5)
file_button = Button(root, text="Selecionar Arquivo", command=adicionar_arquivo)
file_button.grid(row=7, column=2, padx=10, pady=5)


Label(root, text="Arquivos Anexados:").grid(row=8, column=0, padx=10, pady=5)
attachment_listbox = Listbox(root, height=5)
attachment_listbox.grid(row=8, column=1, padx=10, pady=5)


remove_attachment_button = Button(root, text="Remover Anexo", command=remover_arquivo)
remove_attachment_button.grid(row=8, column=2, padx=10, pady=5)


send_button = Button(root, text="Enviar Agora", command=send_email)
send_button.grid(row=9, column=0, columnspan=2, pady=10)


Label(root, text="Data de Envio (DD/MM/AAAA):").grid(row=10, column=0, padx=10, pady=5)
data_entry = Entry(root, width=20)
data_entry.grid(row=10, column=1, padx=10, pady=5)

Label(root, text="Hora de Envio (HH:MM):").grid(row=11, column=0, padx=10, pady=5)
hora_entry = Entry(root, width=20)
hora_entry.grid(row=11, column=1, padx=10, pady=5)


schedule_button = Button(root, text="Agendar Envio", command=agendar_envio_unico)
schedule_button.grid(row=12, column=0, columnspan=2, pady=10)

# carregar e-mails salvos ao iniciar
carregar_emails()


root.mainloop()
