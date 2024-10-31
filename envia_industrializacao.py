import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from datetime import datetime
import traceback
import pandas as pd
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook


class EnviaIndustrializacao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)
        
    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

    def lanca_dados_coluna(self, bloco_book, celula, informacao, fonte_alinhamento):
        try:
            nome_fonte, tamanho_fonte, e_negrito, alin_hor, alin_ver = fonte_alinhamento

            celula_sup_esq = bloco_book[celula]
            cel = bloco_book[celula]
            cel.alignment = Alignment(horizontal=alin_hor, vertical=alin_ver, text_rotation=0,
                                      wrap_text=False, shrink_to_fit=False, indent=0)
            cel.font = Font(name=nome_fonte, size=tamanho_fonte, bold=e_negrito)
            celula_sup_esq.value = informacao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_mesclado(self, bloco_book, mesclado, celula, informacao, fonte_alinhamento):
        try:
            nome_fonte, tamanho_fonte, e_negrito, alin_hor, alin_ver = fonte_alinhamento

            bloco_book.merge_cells(mesclado)
            celula_sup_esq = bloco_book[celula]
            cel = bloco_book[celula]
            cel.alignment = Alignment(horizontal=alin_hor, vertical=alin_ver, text_rotation=0,
                                      wrap_text=False, shrink_to_fit=False, indent=0)
            cel.font = Font(name=nome_fonte, size=tamanho_fonte, bold=e_negrito)
            celula_sup_esq.value = informacao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def edita_alinhamento(self, cell, ali_horizontal='center', ali_vertical='center', rotacao=0, quebra_linha=False,
                          encolher=False, recuar=0):
        try:
            cell.alignment = Alignment(horizontal=ali_horizontal,
                                       vertical=ali_vertical,
                                       text_rotation=rotacao,
                                       wrap_text=quebra_linha,
                                       shrink_to_fit=encolher,
                                       indent=recuar)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def edita_bordas(self, cell):
        try:
            cell.border = Border(left=Side(border_style='thin', color='00000000'),
                                 right=Side(border_style='thin', color='00000000'),
                                 top=Side(border_style='thin', color='00000000'),
                                 bottom=Side(border_style='thin', color='00000000'),
                                 diagonal=Side(border_style='thick', color='00000000'),
                                 diagonal_direction=0,
                                 outline=Side(border_style='thin', color='00000000'),
                                 vertical=Side(border_style='thin', color='00000000'),
                                 horizontal=Side(border_style='thin', color='00000000'))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def linhas_colunas_p_edicao(self, sheet, min_linha, max_linha, min_coluna, max_coluna):
        try:
            for row in sheet.iter_rows(min_row=min_linha,
                                       max_row=max_linha,
                                       min_col=min_coluna,
                                       max_col=max_coluna):
                for cell in row:
                    yield cell

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def edita_preenchimento(self, cell):
        try:
            cell.fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def edita_fonte(self, cell, nome_fonte='Calibri', tamanho=10, negrito=False):
        try:
            cell.font = Font(name=nome_fonte, size=tamanho, bold=negrito)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mesclar_descricao(self, ws, desc_produtos, ini_prod):
        try:
            startcol_descricao = 5
            endcol_descricao = 10

            for idx, desc in enumerate(desc_produtos):
                row_num = ini_prod + idx

                ws.cell(row=row_num, column=startcol_descricao).value = desc
                ws.cell(row=row_num, column=startcol_descricao).alignment = Alignment(horizontal='center',
                                                                                      vertical='bottom', wrap_text=True)

                if len(desc_produtos) > 1:
                    ws.merge_cells(start_row=row_num, start_column=startcol_descricao, end_row=row_num,
                                   end_column=endcol_descricao)

            for col_num in range(startcol_descricao + 1, endcol_descricao + 1):
                for idx in range(len(desc_produtos)):
                    cell = ws.cell(row=ini_prod + idx, column=col_num)
                    cell.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_do_cod_e_qtde(self, writer, inicio_produtos, dados_tab):
        try:
            df_dados_um = pd.DataFrame(dados_tab, columns=['*COD. PROD.', 'QTDE'])
            codigo_int = {'*COD. PROD.': int}
            df_dados_um = df_dados_um.astype(codigo_int)
            qtde_float = {'QTDE': float}
            df_dados_um = df_dados_um.astype(qtde_float)
            df_dados_um.to_excel(writer, 'ORIGINAL', startrow=inicio_produtos - 1, startcol=2, header=False,
                                 index=False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_descricao(self, ws, writer, inicio_produtos, dados_tab):
        try:
            desc_produtos = [tabi[2] for tabi in dados_tab]
            df_descricao = pd.DataFrame({'DESCRIÇÃO DO PRODUTO': desc_produtos})
            df_descricao.to_excel(writer, 'ORIGINAL', startrow=inicio_produtos - 1, startcol=4, header=False,
                                  index=False)

            ini_coluna_descr = 5
            fim_coluna_descr = 10

            for index, desc in enumerate(desc_produtos):
                num_linha = inicio_produtos + index

                ws.cell(row=num_linha, column=ini_coluna_descr).value = desc
                ws.cell(row=num_linha, column=ini_coluna_descr).alignment = Alignment(horizontal='center',
                                                                                      vertical='bottom',
                                                                                      wrap_text=True)

                if len(desc_produtos) > 1:
                    ws.merge_cells(start_row=num_linha,
                                   start_column=ini_coluna_descr,
                                   end_row=num_linha,
                                   end_column=fim_coluna_descr)

            for col_num in range(ini_coluna_descr + 1, fim_coluna_descr + 1):
                for idx in range(len(desc_produtos)):
                    cell = ws.cell(row=inicio_produtos + idx, column=col_num)
                    cell.alignment = Alignment(horizontal='center',
                                               vertical='bottom',
                                               wrap_text=True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_da_um_ate_unid(self, writer, inicio_produtos, dados_tab):
        try:
            df = pd.DataFrame(dados_tab, columns=['UN', 'IND.', 'VLR UNIT.'])
            df.to_excel(writer, 'ORIGINAL', startrow=inicio_produtos - 1, startcol=10, header=False, index=False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_total(self, writer, inicio_produtos, dados_tab):
        try:
            df_dados_dois = pd.DataFrame(dados_tab, columns=['VLR TOTAL'])
            df_dados_dois.to_excel(writer, 'ORIGINAL', startrow=inicio_produtos - 1, startcol=13, header=False,
                                   index=False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gera_excel(self, dados_fornc, list_produtos, caminho_remessa):
        try:
            tem_s = 0
            tem_n = 0

            id_fornc = dados_fornc[0]
            nome_fornc = dados_fornc[1]

            texto_fornc = f"{id_fornc} - {nome_fornc}"

            dados_tabela1 = []

            for index, iii in enumerate(list_produtos):
                seq = index + 1
                cod, descr, ref, um, qtde, conj, arq_org, camin = iii

                didi = (seq, cod, descr, um, qtde, "", conj, "")
                dados_tabela1.append(didi)

            dados_p_descricao = []

            caminho_arquivo = "modelo_remessa.xlsx"

            book = load_workbook(caminho_arquivo)

            writer = pd.ExcelWriter(caminho_remessa, engine='openpyxl')

            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

            ws = book.active

            dados_um = []
            dados_dois = []
            dados_tres = []

            texto_ocs = ""

            for tabi in dados_tabela1:
                dados_p_descricao.append(tabi)
                seq, cod, desc, um, qtde, unit, conj, total = tabi

                if conj == "Produtos Acabados":
                    indust = "SIM"
                else:
                    indust = "NÃO"

                qtdezinha_float = float(qtde)

                if indust == "SIM":
                    tem_s += 1
                if indust == "NÃO":
                    tem_n += 1

                dados1 = (cod, qtdezinha_float)
                dados2 = (um, indust, unit)

                dados_um.append(dados1)
                dados_dois.append(dados2)
                dados_tres.append(total)

            inicio_produtos = 21

            linhas_produtos = len(dados_tabela1)
            if linhas_produtos < 10:
                dif = 10 - linhas_produtos
                totalzao = dif + linhas_produtos
                dedos = ['', '', '', '', '', '', '', '']
                for repite in range(dif):
                    dados_p_descricao.append(dedos)
            else:
                totalzao = linhas_produtos

            num_seq = [tabi[0] for tabi in dados_tabela1]
            df_ordem = pd.DataFrame({'ITEM': num_seq})
            seq_int = {'ITEM': int}
            df_sequencia = df_ordem.astype(seq_int)
            df_sequencia.to_excel(writer, 'ORIGINAL', startrow=inicio_produtos - 1, startcol=1, header=False,
                                  index=False)

            self.manipula_do_cod_e_qtde(writer, inicio_produtos, dados_um)
            self.manipula_da_um_ate_unid(writer, inicio_produtos, dados_dois)
            self.manipula_total(writer, inicio_produtos, dados_tres)
            self.manipula_descricao(ws, writer, inicio_produtos, dados_p_descricao)

            for cell in self.linhas_colunas_p_edicao(ws, inicio_produtos, inicio_produtos + totalzao - 1, 2, 14):
                self.edita_bordas(cell)
                self.edita_alinhamento(cell, ali_vertical='bottom')
                self.edita_fonte(cell)

            for cell in self.linhas_colunas_p_edicao(ws, inicio_produtos, inicio_produtos + totalzao - 1, 1, 1):
                self.edita_alinhamento(cell, ali_vertical='bottom')
                self.edita_fonte(cell)

                cell.border = Border(right=Side(border_style='medium', color='00000000'))

            for cell in self.linhas_colunas_p_edicao(ws, inicio_produtos, inicio_produtos + totalzao - 1, 14, 14):
                self.edita_preenchimento(cell)
                cell.border = Border(left=Side(border_style='thin', color='00000000'),
                                     right=Side(border_style='medium', color='00000000'),
                                     top=Side(border_style='thin', color='00000000'),
                                     bottom=Side(border_style='thin', color='00000000'),
                                     diagonal=Side(border_style='thick', color='00000000'),
                                     diagonal_direction=0,
                                     outline=Side(border_style='thin', color='00000000'),
                                     vertical=Side(border_style='thin', color='00000000'),
                                     horizontal=Side(border_style='thin', color='00000000'))

            altura_celula = 24.75
            for linha in range(inicio_produtos, inicio_produtos + totalzao):
                ws.row_dimensions[linha].height = altura_celula

            personalizacao = ['Times New Roman', 10, False, 'left', 'bottom']
            self.lanca_dados_mesclado(ws, 'C11:N11', 'C11', texto_fornc, personalizacao)

            texto_operacao = ""

            texto_operacao += "5.901 - REM. P/INDUSTR. POR ENCOM."

            if tem_s:
                personalizacao = ['Times New Roman', 10, True, 'center', 'center']
                informacao = "X"
                self.lanca_dados_coluna(ws, "E5", informacao, personalizacao)

            if tem_n:
                personalizacao = ['Times New Roman', 10, True, 'center', 'center']
                informacao = "X"
                self.lanca_dados_coluna(ws, "E7", informacao, personalizacao)

            personalizacao = ['Times New Roman', 10, True, 'left', 'bottom']
            self.lanca_dados_mesclado(ws, 'E16:N16', 'E16', texto_operacao, personalizacao)

            personalizacao = ['Times New Roman', 10, False, 'left', 'bottom']
            self.lanca_dados_mesclado(ws, 'C13:N13', 'C13', texto_ocs, personalizacao)

            linha_vazia = inicio_produtos + totalzao
            for linha in range(linha_vazia, linha_vazia + 1):
                ws.row_dimensions[linha].height = altura_celula

                ws.merge_cells(f'C{linha}:M{linha}')
                celula_sup_esq = ws[f'C{linha}']
                celula_sup_esq.value = ""
                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 3, 13):
                    cell.border = Border(left=Side(border_style='thin', color='00000000'),
                                         right=Side(border_style='thin', color='00000000'),
                                         top=Side(border_style='medium', color='00000000'),
                                         bottom=Side(border_style='thin', color='00000000'),
                                         diagonal=Side(border_style='thick', color='00000000'),
                                         diagonal_direction=0,
                                         outline=Side(border_style='thin', color='00000000'),
                                         vertical=Side(border_style='thin', color='00000000'),
                                         horizontal=Side(border_style='thin', color='00000000'))

                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 2, 2):
                    cell.border = Border(top=Side(border_style='medium', color='00000000'))

                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 14, 14):
                    cell.border = Border(top=Side(border_style='medium', color='00000000'))

            segunda_linha_vazia = inicio_produtos + totalzao + 1
            for linha in range(segunda_linha_vazia, segunda_linha_vazia + 1):
                ws.row_dimensions[linha].height = altura_celula

                ws.merge_cells(f'C{linha}:M{linha}')
                celula_sup_esq = ws[f'C{linha}']
                celula_sup_esq.value = ""
                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 3, 13):
                    self.edita_bordas(cell)

            segunda_linha_vazia = inicio_produtos + totalzao + 2
            for linha in range(segunda_linha_vazia, segunda_linha_vazia + 1):
                ws.row_dimensions[linha].height = altura_celula

                ws.merge_cells(f'C{linha}:H{linha}')
                celula_sup_esq = ws[f'C{linha}']
                celula_sup_esq.value = "* Produtos com * cod informados se possível."
                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 3, 8):
                    self.edita_alinhamento(cell, ali_vertical='bottom')
                    self.edita_fonte(cell)
                    self.edita_bordas(cell)

            segunda_linha_vazia = inicio_produtos + totalzao + 3
            for linha in range(segunda_linha_vazia, segunda_linha_vazia + 1):
                ws.row_dimensions[linha].height = altura_celula

                ws.merge_cells(f'C{linha}:H{linha}')
                celula_sup_esq = ws[f'C{linha}']
                celula_sup_esq.value = "* Cada nota possui somente 18 linhas disponíveis."
                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 3, 8):
                    self.edita_alinhamento(cell, ali_vertical='bottom')
                    self.edita_fonte(cell)
                    self.edita_bordas(cell)

            segunda_linha_vazia = inicio_produtos + totalzao + 4
            for linha in range(segunda_linha_vazia, segunda_linha_vazia + 1):
                ws.row_dimensions[linha].height = altura_celula

                ws.merge_cells(f'C{linha}:H{linha}')
                celula_sup_esq = ws[f'C{linha}']
                celula_sup_esq.value = "* Favor digitar volume e peso no final de cada nota."
                for cell in self.linhas_colunas_p_edicao(ws, linha, linha, 3, 8):
                    self.edita_alinhamento(cell, ali_vertical='bottom')
                    self.edita_fonte(cell)
                    self.edita_bordas(cell)

            linha_bruto = inicio_produtos + totalzao + 2
            person = ['Times New Roman', 10, False, 'center', 'bottom']
            self.lanca_dados_mesclado(ws, f'K{linha_bruto}:M{linha_bruto}', f'K{linha_bruto}', "", person)
            for cell in self.linhas_colunas_p_edicao(ws, linha_bruto, linha_bruto, 11, 13):
                self.edita_bordas(cell)

            personalizacao = ['Times New Roman', 10, True, 'center', 'bottom']
            informacao = "PESO BRUTO"
            self.lanca_dados_coluna(ws, f'J{linha_bruto}', informacao, personalizacao)
            for cell in self.linhas_colunas_p_edicao(ws, linha_bruto, linha_bruto, 10, 10):
                self.edita_bordas(cell)

            linha_liq = inicio_produtos + totalzao + 3
            person = ['Times New Roman', 10, False, 'center', 'bottom']
            self.lanca_dados_mesclado(ws, f'K{linha_liq}:M{linha_liq}', f'K{linha_liq}', "", person)
            for cell in self.linhas_colunas_p_edicao(ws, linha_liq, linha_liq, 11, 13):
                self.edita_bordas(cell)

            personalizacao = ['Times New Roman', 10, True, 'center', 'bottom']
            informacao = "PESO LÍQUIDO"
            self.lanca_dados_coluna(ws, f'J{linha_liq}', informacao, personalizacao)
            for cell in self.linhas_colunas_p_edicao(ws, linha_liq, linha_liq, 10, 10):
                self.edita_bordas(cell)

            linha_vol = inicio_produtos + totalzao + 4
            person = ['Times New Roman', 10, False, 'center', 'bottom']
            self.lanca_dados_mesclado(ws, f'K{linha_vol}:M{linha_vol}', f'K{linha_vol}', "", person)
            for cell in self.linhas_colunas_p_edicao(ws, linha_vol, linha_vol, 11, 13):
                self.edita_bordas(cell)

            personalizacao = ['Times New Roman', 10, True, 'center', 'bottom']
            informacao = "VOLUME"
            self.lanca_dados_coluna(ws, f'J{linha_vol}', informacao, personalizacao)
            for cell in self.linhas_colunas_p_edicao(ws, linha_vol, linha_vol, 10, 10):
                self.edita_bordas(cell)

            writer.save()

            print(f'excel gerado com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_email(self):
        try:
            to = ['<maquinas@unisold.com.br>']

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

            msg_final = f"Att,\n" \
                        f"Suzuki Máquinas Ltda\n" \
                        f"Fone (51) 3561.2583/(51) 3170.0965\n\n" \
                        f"Mensagem enviada automaticamente, por favor não responda.\n\n" \
                        f"Se houver algum problema com o recebimento de emails ou conflitos com o arquivo excel, " \
                        f"favor entrar em contato pelo email maquinas@unisold.com.br.\n\n"

            email_user = 'ti.ahcmaq@gmail.com'
            password = 'poswxhqkeaacblku'

            return saudacao, msg_final, email_user, to, password

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_arquivo(self, caminho_arquivo):
        try:
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
            else:
                print("O arquivo não existe no caminho especificado.")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_arquivos(self):
        try:
            caminho = r'\\publico\C\OP\Terceiros/'
            extensao = ".pdf"

            arquivos_pdfs = []

            for arquivo in os.listdir(caminho):
                if arquivo.endswith(extensao):
                    caminho_arquivo = os.path.join(caminho, arquivo)
                    arquivos_pdfs.append(caminho_arquivo)

            return arquivos_pdfs

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self, nome_fornc, arquivos_set, caminho_remessa):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            to = ['<maquinas@unisold.com.br>']

            subject = f'IND - Enviar Material para Industrialização {nome_fornc}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Separar material para ser enviado para o Fornecedor {nome_fornc} como remessa de " \
                   f"industrialização.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            arquivos_list = list(arquivos_set)
            if arquivos_list:
                for ii in arquivos_list:
                    nome_arq, caminho_arq = ii

                    attachment1 = open(caminho_arq, 'rb')
                    part1 = MIMEBase('application', "octet-stream")
                    part1.set_payload(attachment1.read())
                    encoders.encode_base64(part1)
                    part1.add_header('Content-Disposition', 'attachment', filename=Header(nome_arq, 'utf-8').encode())
                    msg.attach(part1)
                    attachment1.close()

            attachment2 = open(caminho_remessa, 'rb')
            part2 = MIMEBase('application', "octet-stream")
            part2.set_payload(attachment2.read())
            encoders.encode_base64(part2)
            part2.add_header('Content-Disposition', 'attachment', filename=Header(caminho_remessa, 'utf-8').encode())
            msg.attach(part2)
            attachment2.close()

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)

            server.quit()

            print(f'Industrialização {nome_fornc} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_nao_acha_desenho(self, num_desenho, arq_original, caminho_original):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            subject = f'OP - Não foi encontrado o desenho {num_desenho}'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"O desenho {num_desenho} não foi encontrado no cadastro dos produtos.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=Header(arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'Desenho {num_desenho} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_tipo_nao_cadastrado(self, cod_prod, arq_original, caminho_original):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            subject = f'OP - O produto {cod_prod} não tem o "Tipo de Material" definido no cadsatro'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Favor definir o Tipo de Material no cadastro do produto: {cod_prod}.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'produto sem tipo {cod_prod} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_tipo_nao_industrializado(self, cod_prod, arq_original, caminho_original):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            subject = f'OP - O produto {cod_prod} não tem o "Tipo de Material" definido como Industrializado'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Favor definir o Tipo de Material como Industrializado no cadastro do produto: {cod_prod}.\n\n" \
                   f"{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'produto sem tipo Industrializado {cod_prod} enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email_desenho_duplicado(self, lista_produtos, arq_original, caminho_original):
        try:
            saudacao, msg_final, email_user, to, password = self.dados_email()

            numero = lista_produtos[0][0]

            subject = f'OP - Foram encontrados produtos com desenho duplicado no cadastro - {numero}!'

            msg = MIMEMultipart()
            msg['From'] = email_user
            msg['Subject'] = subject

            body = f"{saudacao}\n\n" \
                   f"Segue lista dos itens encontrados:\n\n"

            for i in lista_produtos:
                body += f"{i}.\n" \

            body += f"\n\n{msg_final}"

            msg.attach(MIMEText(body, 'plain'))

            attachment = open(caminho_original, 'rb')

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',
                            filename=Header(arq_original, 'utf-8').encode())
            msg.attach(part)

            text = msg.as_string()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email_user, password)

            server.sendmail(email_user, to, text)
            attachment.close()

            server.quit()

            print(f'produto desenho duplicado enviada com sucesso!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_produto(self, num_desenho, qtde, arq_original, caminho):
        try:
            nova_lista = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao, TERCEIRIZADOOBS FROM produto where obs = '{num_desenho}';")
            select_prod = cursor.fetchall()

            idez, cod_pai, id_estrut, servico = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"conj.conjunto, prod.unidade, prod.localizacao, prod.quantidade "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where estprod.id_estrutura = {id_estrut} "
                           f"order by conj.conjunto DESC, prod.descricao ASC;")
            tabela_estrutura = cursor.fetchall()

            if tabela_estrutura:
                itens_na_estrut = len(tabela_estrutura)

                if itens_na_estrut == 1:
                    cod, descr, ref, conj, um, local, saldo = tabela_estrutura[0]
                    if float(saldo) >= qtde:
                        dados = (cod_pai, cod, descr, ref, um, qtde, conj, arq_original, caminho, servico)
                        nova_lista.append(dados)
                    else:
                        print(f"O saldo do {cod} é insuficiente para a quantidade solicitada: {saldo}")
                else:
                    print("ESTRUTURA DO PRODUTO TEM MAIS ITENS")

            return nova_lista

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_produto_pelo_desenho(self, num_des_com_d, arq_original, caminho):
        try:
            verifica_produto = False

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT DISTINCT codigo, descricao, COALESCE(obs, ''), unidade, COALESCE(tipomaterial, ''), "
                f"COALESCE(localizacao, '') "
                f"FROM produto "
                f"WHERE obs = '{num_des_com_d}';")
            detalhes_produto = cursor.fetchall()

            if detalhes_produto:
                qtde_itens = len(detalhes_produto)
                if qtde_itens == 1:
                    for i in detalhes_produto:
                        cod_prod, descr_prod, ref_prod, um, tipo, local = i

                        if tipo:
                            if tipo == "119":
                                verifica_produto = True
                            else:
                                self.envia_email_tipo_nao_industrializado(cod_prod, arq_original, caminho)
                        else:
                            self.envia_email_tipo_nao_cadastrado(cod_prod, arq_original, caminho)
                else:
                    self.envia_email_desenho_duplicado(detalhes_produto, arq_original, caminho)
            else:
                self.envia_email_nao_acha_desenho(num_des_com_d, arq_original, caminho)

            return verifica_produto

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self):
        try:
            lista_indust = []

            arquivos_pdfs = self.procura_arquivos()

            for arqsv in arquivos_pdfs:
                caminho = arqsv

                arq_original = caminho[25:]

                if " - " in arq_original:
                    inicio = arq_original.find(" - ")
                    dadinhos1 = arq_original[inicio + 3:]
                elif "-" in arq_original:
                    inicio = arq_original.find("-")
                    dadinhos1 = arq_original[inicio + 1:]
                else:
                    inicio = ""
                    dadinhos1 = ""

                if inicio and dadinhos1:
                    qtde = arq_original[:inicio]
                    try:
                        qtde_produto = int(qtde)

                        num_desenho_arq = dadinhos1[:-4]
                        num_des_com_d = f"D {num_desenho_arq}"

                        print("Arquivo Original:", arq_original)

                        verifica_prod = self.procura_produto_pelo_desenho(num_des_com_d, arq_original, caminho)

                        if verifica_prod:
                            try:
                                numero_maquina = num_desenho_arq[:2]
                                num_maq_int = int(numero_maquina)
                                print(num_maq_int, num_desenho_arq, qtde_produto)
                            except ValueError:
                                print("NUMERO DE DESENHO FORA DE PADRÃO!")

                    except ValueError:
                        print("A QUANTIDADE NÃO ESTÁ CERTA")

                else:
                    print('O ARQUIVO NÃO ESTÁ NO PADRÃO CERTO: "QTDE - NUM. DESENHO"')

            if lista_indust:
                for i in lista_indust:
                    print("lista_indust", i)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


chama_classe = EnviaIndustrializacao()
chama_classe.manipula_comeco()
