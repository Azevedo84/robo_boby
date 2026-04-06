from core.erros import trata_excecao
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

def dados_email():
    try:
        email_user = 'ti.ahcmaq@gmail.com'
        password = "fsno grka ifsq jmzm"

        current_time = (datetime.now())
        horario = current_time.strftime('%H')
        hora_int = int(horario)
        saudacao = ""
        if 4 < hora_int < 13:
            saudacao = "Bom Dia!"
        elif 12 < hora_int < 19:
            saudacao = "Boa Tarde!"
        elif hora_int > 18:
            saudacao = "Boa Noite!"
        elif hora_int < 5:
            saudacao = "Boa Noite!"

        msg_final = ""

        msg_final += f"Att,\n"
        msg_final += f"Suzuki Máquinas Ltda\n"
        msg_final += f"Fone (51) 3561.2583/(51) 3170.0965\n\n"
        msg_final += f"🟦 Mensagem gerada automaticamente pelo sistema de planejamento e Controle da Produção (PCP) do ERP Suzuki.\n"
        msg_final += "🔸Por favor, não responda este e-mail diretamente."

        return saudacao, msg_final, email_user, password

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_desenho_duplicado(destinatario, texto, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA PI - DESENHO DUPLICADO {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO desenho {desenho} está duplicado!\n\n"

        for i in texto:
            body += f"{i}\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, destinatario, text)
        server.quit()

        print("email enviado DUPLICADO",desenho)

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_desenho_sem_vinculo(destinatario, texto, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA PI - DESENHO SEM VÍNCULO {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO desenho {desenho} tem problemas com vínculos dos arquivos!\n\n"

        for i in texto:
            body += f"{i}\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, destinatario, text)
        server.quit()

        print("email enviado SEM VINCULO")

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_sem_idw(destinatario, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA PI - DESENHO {desenho} SEM IDW'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO desenho {desenho} está sem IDW!\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, destinatario, text)
        server.quit()

        print("email enviado SEM IDW")

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_arquivo_nao_encontrado(destinatario, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA PI - ARQUIVO NÃO FOI ENCONTRADO!'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO Arquivo {desenho} não foi encontrado!\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, destinatario, text)
        server.quit()

        print("email enviado ARQUIVO NÃO ENCONTRADO")

    except Exception as e:
        trata_excecao(e)
        raise