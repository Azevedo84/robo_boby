

def envia_email_sem_ncm(self, caminho, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - SEM NCM NO DESENHO {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f'{saudacao}\n\nO desenho {desenho} não possui NCM cadastrada!\n\n'

        body += f"'{caminho}'\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado SEM NCM")

    except Exception as e:
        trata_excecao(e)
        raise

def enviar_email_atualiza_cadastro_referencia_erp(self, dados_duplicados, desenho, codigo, referencia):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - ATUALIZAR REFERÊNCIA NO CADASTRO DO PRODUTO {codigo}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f'{saudacao}\n\nAtualizar o cadastro do produto {codigo}:\n\n'

        body += f"A Referência atual é: {referencia}. Alterar para D {desenho}\n\n"

        if dados_duplicados:
            for i in dados_duplicados:
                id_arquivo, caminho, num_desenho, descr = i

                body += f"{descr} - {num_desenho}: '{caminho}'\n\n"

        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado ATUALIZAR REFERENCIA CADASTRO ERP")

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_desenho_duplicado(self, caminhos, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - DESENHO DUPLICADO {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO desenho {desenho} está duplicado!\n\n"

        for i in caminhos:
            nome_arquivo, nome_base, tipo, classifica, caminho = i

            body += f"{nome_base} - {tipo} - {classifica}\n\n"
            body += f"{caminho}\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado DUPLICADO")

    except Exception as e:
        print(e)

def envia_email_sem_descricao_produto(self, caminho, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - SEM DESCRIÇÃO NO DESENHO {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f'{saudacao}\n\nO desenho {desenho} não possui descrição do produto!\n\n'

        body += f"'{caminho}'\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado SEM DESCRIÇÃO PRODUTO")

    except Exception as e:
        trata_excecao(e)
        raise

def envia_email_sem_idw(self, caminho, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - SEM IDW {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO desenho {desenho} não tem a versão IDW!\n\n"
        body += f"{caminho}\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado SEM IDW")

    except Exception as e:
        print(e)

def enviar_email_criar_estrutura_nova(self, dados_estrutura):
    try:
        cod_prod, descricao, nome_base, caminho, dados_op = dados_estrutura

        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - CRIAR ESTRUTURA {nome_base}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f"{saudacao}\n\nO produto Código: {cod_prod} - {descricao} precisar criar um versão da Estrutura\n\n"

        body += f"{caminho}\n\n"

        if dados_op:
            body += f"Existe Ordens de Produção abertas para este produto:\n\n"
            for i in dados_op:
                body += f"{i}\n\n"

        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado criar estrutura!")

    except Exception as e:
        print(e)

def envia_email_sem_medida_corte(self, caminho, desenho):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - MEDIDA DE CORTE DIVERGENTE {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f'{saudacao}\n\nO desenho {desenho} IDW não possui medida de corte conforme propriedade "Comprimento"!\n\n'

        body += f"'{caminho}'\n\n"
        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado SEM MEDIDA DE CORTE")

    except Exception as e:
        trata_excecao(e)
        raise

def enviar_email_muitos_caracteres(self, nome_campo, valor_campo, caminho, desenho, qtde):
    try:
        saudacao, msg_final, email_user, password = dados_email()

        subject = f'ENGENHARIA/ERP - MUITOS CARACTERES NO CAMPO {nome_campo} - {desenho}'

        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['Subject'] = subject

        body = f'{saudacao}\n\nO campo {nome_campo} tem muitos caracteres e não pode ser cadastrado no desenho {desenho}.\n\n'

        body += f"'{nome_campo}: {valor_campo} - {qtde} caracteres'\n\n"

        body += f"'{caminho}'\n\n"

        body += f"\n{msg_final}"

        msg.attach(MIMEText(body, 'plain'))

        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, password)

        server.sendmail(email_user, self.destinatario, text)
        server.quit()

        print("email enviado MUITOS CARACTERES NO CAMPO")

    except Exception as e:
        trata_excecao(e)
        raise