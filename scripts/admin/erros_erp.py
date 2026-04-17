from core.banco import conecta, conecta_robo
from core.erros import trata_excecao
from core.email_service import dados_email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class EnviaErrosERP:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>']

    def manipula_comeco(self):
        try:
            tabela_final = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ID, NOME_PC, ARQUIVO, FUNCAO, MENSAGEM, ENTREGUE, CRIACAO "
                           f"FROM ZZZ_ERROS "
                           f"WHERE (entregue IS NULL OR entregue = '');")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao = i

                    cursor = conecta_robo.cursor()
                    cursor.execute(f"SELECT ID, ID_ERRO FROM ENVIA_ERROS_ERP "
                                   f"where ID_ERRO = {id_erro};")
                    dados_criados = cursor.fetchall()

                    if not dados_criados:
                        cursor = conecta_robo.cursor()
                        cursor.execute(f"Insert into ENVIA_ERROS_ERP (ID, ID_ERRO) "
                                       f"values (GEN_ID(GEN_ENVIA_ERROS_ERP_ID,1), {id_erro});")

                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE ZZZ_ERROS SET ENTREGUE = 'S' "
                                       f"WHERE id= {id_erro};")

                        conecta_robo.commit()

                        dados = (id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao)
                        tabela_final.append(dados)

                if tabela_final:
                    self.envia_email(tabela_final)

        except Exception as e:
            trata_excecao(e)
            raise

    def envia_email(self, dados):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            subject = f'ERROS - PROBLEMAS COM O ERP - SUZUKI!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\nAbaixo a lista de erros do ERP - Suzuki:\n\n"

            for i in dados:
                id_erro, nome_pc, arquivo, funcao, mensagem, entregue, criacao = i

                data_criacao = criacao.strftime('%d/%m/%Y %H:%M:%S')

                body += f"- Nome Computador: {nome_pc}\n" \
                        f"- Data Criação: {data_criacao}\n" \
                        f"- Nome do Arquivo: {arquivo}\n" \
                        f"- Nome da Função: {funcao}\n" \
                        f"- Mensagem de Erro: {mensagem}\n\n"

            body += f"\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, self.destinatario, text)
            server.quit()

            print("email enviado")

        except Exception as e:
            trata_excecao(e)
            raise


chama_classe = EnviaErrosERP()
chama_classe.manipula_comeco()
