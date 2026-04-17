import os
from pathlib import Path
import sys

os.chdir(r"C:\Users\Anderson\PycharmProjects\robo_boby")

BASE_DIR = Path(__file__).resolve().parents[2]

if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from core.banco import conecta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from core.erros import trata_excecao
from core.email_service import dados_email


class EnviaCadastroProduto:
    def __init__(self):
        self.destinatario = ['<maquinas@unisold.com.br>', '<ahcmaquinas@gmail.com>']
        self.manipula_dados_prod()

    def envia_email(self, num_reg, lista):
        try:
            saudacao, msg_final, email_user, password = dados_email()

            msg_final_html = msg_final.replace("\n", "<br>")

            obs_curto = num_reg[:20]
            if len(num_reg) > 20:
                obs_curto += "..."

            subject = f'PRO - Cadastro Registro {obs_curto}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            linhas_html = ""
            for item in lista:
                linhas_html += f"""
                <tr>
                    <td>{item['codigo']}</td>
                    <td>{item['descricao']}</td>
                    <td>{item['um']}</td>
                </tr>
                """

            html = f"""
            <p>{saudacao}</p>

            <p>Produtos cadastrados:</p>

            <table border="1" cellpadding="5" cellspacing="0">
                <tr>
                    <th>Código</th>
                    <th>Descrição</th>
                    <th>UM</th>
                </tr>
                {linhas_html}
            </table>

            <p>{msg_final_html}</p>
            """

            msg.attach(MIMEText(html, 'html'))

            text = msg.as_string()

            # ✅ SMTP seguro
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(email_user, password)
                server.sendmail(email_user, self.destinatario, text)

            print(f"Email enviado para OBS: {num_reg}")
            return True

        except Exception as e:
            trata_excecao(e)
            return False

    def manipula_dados_prod(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("""
                SELECT id, obs, descricao, descr_compl, referencia, um, ncm,
                       kg_mt, fornecedor, data_criacao, codigo
                FROM PRODUTOPRELIMINAR
                WHERE codigo IS NOT NULL
            """)
            dados_banco = cursor.fetchall()

            dados_dict = {}

            if dados_banco:
                for i in dados_banco:
                    id_pre, obs, descr, compl, ref, um, ncm, kg_mt, forn, emissao, codigo = i

                    obs = obs.strip()

                    if obs not in dados_dict:
                        dados_dict[obs] = []

                    dados_dict[obs].append({
                        "id_pre": id_pre,
                        "descricao": descr,
                        "compl": compl,
                        "referencia": ref,
                        "um": um,
                        "ncm": ncm,
                        "kg_mt": kg_mt,
                        "fornecedor": forn,
                        "data": emissao,
                        "codigo": codigo
                    })

            # 🔥 envio por grupo (obs)
            for obs, lista in dados_dict.items():

                ids_para_excluir = []

                # (opcional) ordenar por código
                lista.sort(key=lambda x: x["codigo"])

                for item in lista:
                    ids_para_excluir.append(item["id_pre"])

                enviado = self.envia_email(obs, lista)

                if enviado:
                    try:
                        for id_pre in ids_para_excluir:
                            self.excluir_pre_cadastro(id_pre)

                        conecta.commit()  # ✅ commit único

                    except Exception as e:
                        conecta.rollback()  # 🔥 segurança total
                        trata_excecao(e)
                        raise

        except Exception as e:
            trata_excecao(e)
            raise

    def excluir_pre_cadastro(self, num_registro):
        cursor = conecta.cursor()
        cursor.execute(f"DELETE FROM PRODUTOPRELIMINAR WHERE ID = {num_registro};")
        print(f"ID {num_registro} excluído")


# 🚀 execução
EnviaCadastroProduto()