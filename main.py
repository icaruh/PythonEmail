import smtplib
from tkinter import *
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import threading
import time
from datetime import datetime

# Lista de e-mails e senhas salvas manualmente
email_senhas = {
    
}

# Função para selecionar um e-mail salvo e preencher os campos
def selecionar_email(event):
    try:
        selected_index = email_listbox.curselection()[0]
        email = email_listbox.get(selected_index)

        email_entry.delete(0, END)
        email_entry.insert(0, email)

        # Preenche automaticamente a senha correspondente
        if email in email_senhas:
            password_entry.delete(0, END)
            password_entry.insert(0, email_senhas[email])
        else:
            password_entry.delete(0, END)

    except IndexError:
        pass

# Função para remover um e-mail salvo
def remover_email():
    try:
        selected_index = email_listbox.curselection()[0]
        email = email_listbox.get(selected_index)

        if email in email_senhas:
            del email_senhas[email]  # Remove do dicionário
            email_listbox.delete(selected_index)  # Remove da listbox
            messagebox.showinfo("Sucesso", f"E-mail {email} removido com sucesso!")
        
        email_entry.delete(0, END)
        password_entry.delete(0, END)

    except IndexError:
        messagebox.showerror("Erro", "Selecione um e-mail para remover!")

# Enviar e-mail
def send_email():
    try:
        sender_email = email_entry.get().strip()
        password = password_entry.get().strip()
        recipient_email = recipient_entry.get().strip()
        subject = subject_entry.get()
        body = body_entry.get("1.0", END)
        file_path = file_path_var.get()

        if not password:
            messagebox.showerror("Erro", "Senha não encontrada para este e-mail!")
            return

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
            part.add_header('Content-Disposition', f'attachment; filename="{file_path.split("/")[-1]}"')
            msg.attach(part)

        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()

        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao enviar o e-mail: {str(e)}")

# Função para adicionar um novo e-mail à lista manualmente
def adicionar_email():
    novo_email = email_entry.get().strip()
    nova_senha = password_entry.get().strip()

    if novo_email and nova_senha:
        if novo_email not in email_senhas:
            email_senhas[novo_email] = nova_senha
            email_listbox.insert(END, novo_email)
            messagebox.showinfo("Sucesso", f"E-mail {novo_email} salvo com sucesso!")
        else:
            messagebox.showerror("Erro", "E-mail já está salvo!")
    else:
        messagebox.showerror("Erro", "Preencha o e-mail e a senha!")

# Criar interface gráfica
root = Tk()
root.title("Envio de E-mail")

# Lista de e-mails salvos
Label(root, text="E-mails Salvos:").grid(row=0, column=0, padx=10, pady=5)
email_listbox = Listbox(root, height=5)
email_listbox.grid(row=0, column=1, padx=10, pady=5)
for email in email_senhas.keys():
    email_listbox.insert(END, email)

email_listbox.bind("<<ListboxSelect>>", selecionar_email)

# Entrada para novo e-mail
Label(root, text="Novo E-mail:").grid(row=1, column=0, padx=10, pady=5)
email_entry = Entry(root, width=40)
email_entry.grid(row=1, column=1, padx=10, pady=5)

# Entrada para senha
Label(root, text="Senha:").grid(row=2, column=0, padx=10, pady=5)
password_entry = Entry(root, width=40, show="*")  # Esconde a senha com '*'
password_entry.grid(row=2, column=1, padx=10, pady=5)

# Botões para adicionar e remover e-mail
add_button = Button(root, text="Salvar E-mail", command=adicionar_email)
add_button.grid(row=3, column=0, pady=10)

remove_button = Button(root, text="Remover E-mail", command=remover_email)
remove_button.grid(row=3, column=1, pady=10)

# Entrada para destinatário
Label(root, text="E-mail do Destinatário:").grid(row=4, column=0, padx=10, pady=5)
recipient_entry = Entry(root, width=40)
recipient_entry.grid(row=4, column=1, padx=10, pady=5)

# Entrada para assunto
Label(root, text="Assunto:").grid(row=5, column=0, padx=10, pady=5)
subject_entry = Entry(root, width=40)
subject_entry.grid(row=5, column=1, padx=10, pady=5)

# Corpo do e-mail
Label(root, text="Corpo do E-mail:").grid(row=6, column=0, padx=10, pady=5)
body_entry = Text(root, width=40, height=5)
body_entry.grid(row=6, column=1, padx=10, pady=5)

# Botão para anexar arquivos
Label(root, text="Selecionar Arquivo:").grid(row=7, column=0, padx=10, pady=5)
file_path_var = StringVar()
file_path_entry = Entry(root, textvariable=file_path_var, width=40, state="readonly")
file_path_entry.grid(row=7, column=1, padx=10, pady=5)
file_button = Button(root, text="Selecionar Arquivo", command=lambda: file_path_var.set(filedialog.askopenfilename()))
file_button.grid(row=7, column=2, padx=10, pady=5)

# Botão para enviar e-mail
send_button = Button(root, text="Enviar Agora", command=send_email)
send_button.grid(row=8, column=0, columnspan=2, pady=10)

# Iniciar a interface gráfica
root.mainloop()
