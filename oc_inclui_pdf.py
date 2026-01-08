import sys
from conexao import conecta
from forms.tela_oc_inclui_pdf import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import PyPDF2
import numpy as np
import os
from datetime import datetime
from tkinter import messagebox


class TelaOrdem(QMainWindow, Ui_Lanca_Ordem):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.setWindowIcon(QtGui.QIcon(r'C:\Python\menuzinho\icones\compras.png'))

        self.codigo_item = 0
        self.last = 0
        self.number_of_pages = 0
        self.escolha_total = ""

        self.lista_completa_req = []
        self.index_arquivo_para_excluir = 0
        self.arquivo_para_excluir = []

        self.total_ipi = 0
        self.total_mercadorias = 0
        self.total_itens = 0

        self.arquivos_pdf = []
        self.pasta_para_excluir = r'C:/ordens/'

        self.amarelo = "#eaff00"
        self.branco = "#ffffff"

        self.encontra_arquivos()
        self.ajusta_layout()

    def ajusta_layout(self):
        self.btn_LerPDF.clicked.connect(self.imprimir_itens)
        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Cad_Produto.clicked.connect(self.cadastra_produto)
        self.btn_Cad_Fornc.clicked.connect(self.cadastra_fornecedor)

        self.table_Itens_OC.setColumnWidth(0, 35)
        self.table_Itens_OC.setColumnWidth(1, 170)
        self.table_Itens_OC.setColumnWidth(2, 95)
        self.table_Itens_OC.setColumnWidth(3, 15)
        self.table_Itens_OC.setColumnWidth(4, 40)
        self.table_Itens_OC.setColumnWidth(5, 60)
        self.table_Itens_OC.setColumnWidth(6, 60)
        self.table_Itens_OC.setColumnWidth(7, 15)
        self.table_Itens_OC.setColumnWidth(8, 60)
        self.table_Itens_OC.setColumnWidth(9, 35)
        self.table_Itens_OC.setColumnWidth(10, 45)
        self.table_Itens_OC.horizontalHeader().setStyleSheet("QHeaderView::section { background-color:#bfbfbf }")

        self.table_Itens_Req.setColumnWidth(0, 35)
        self.table_Itens_Req.setColumnWidth(1, 45)
        self.table_Itens_Req.setColumnWidth(2, 35)
        self.table_Itens_Req.setColumnWidth(3, 170)
        self.table_Itens_Req.setColumnWidth(4, 95)
        self.table_Itens_Req.setColumnWidth(5, 15)
        self.table_Itens_Req.setColumnWidth(6, 40)
        self.table_Itens_Req.horizontalHeader().setStyleSheet("QHeaderView::section { background-color:#bfbfbf }")

        self.label_Sem_Cad_Forn.setHidden(True)
        self.btn_Cad_Fornc.setHidden(True)

        self.label_Sem_Cad_Prod.setHidden(True)
        self.btn_Cad_Produto.setHidden(True)

    def encontra_arquivos(self):
        pasta = r'C:\ordens'
        for diretorio, subpastas, arquivos in os.walk(pasta):
            if not arquivos:
                messagebox.showinfo(title="Aviso!",
                                    message=f'Não foi encontrado Ordens de compra!\n'
                                            f'Salve as ordens em formato ".pdf" no caminho "C:\ordens"')
                self.btn_LerPDF.setDisabled(True)
            for arquivo in arquivos:
                self.arquivos_pdf.append(str(os.path.join(arquivo)))

        self.combo_arquivos.addItems(self.arquivos_pdf)

    def ler_pdf_excluir(self):
        try:
            self.arquivo_para_excluir = []
            arquivo = self.combo_arquivos.currentText()
            self.index_arquivo_para_excluir = self.arquivos_pdf.index(arquivo)
            caminho = self.pasta_para_excluir + arquivo
            self.arquivo_para_excluir.append(caminho)
            self.itens_para_excluir()
        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "ler_pdf_excluir".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def ler_pdf(self):
        try:
            pasta = r'C:/ordens/'
            arquivo = self.combo_arquivos.currentText()
            caminho = pasta + arquivo
            pdf_file = open(caminho, 'rb')
            read_pdf = PyPDF2.PdfFileReader(pdf_file)
            self.number_of_pages = read_pdf.getNumPages()
            if self.number_of_pages == 1:
                page = read_pdf.getPage(0)
                page_content = page.extractText()
                parsed = ''.join(page_content)
            else:
                parsed = ""
                for qtde_paginas in range(self.number_of_pages):
                    page = read_pdf.getPage(qtde_paginas)
                    page_content = page.extractText()
                    parsed1 = ''.join(page_content)
                    parsed = parsed + parsed1
            pdf_file.close()
            return parsed

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "ler_pdf".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def itens_para_excluir(self):
        try:
            arquivos = self.arquivo_para_excluir
            for arquivo in arquivos:
                try:
                    os.remove(arquivo)
                except OSError as e:
                    messagebox.showinfo(title="Aviso!", message=f"Error:{e.strerror}")

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "itens_para_excluir".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def dados_cabecalho(self):
        try:
            parsed = self.ler_pdf()
            posicao_oc = parsed.find("O.C.:")
            inicio_oc = posicao_oc + 7
            fim_oc = inicio_oc + 5
            escolha_oc = parsed[inicio_oc:fim_oc]
            brancos_oc = escolha_oc.replace(" ", "")
            numero_oc = brancos_oc.replace("\n", "")

            posicao_emissao = parsed.find("Emissão:  ")
            inicio_emissao = posicao_emissao + 10
            fim_emissao = inicio_emissao + 10
            escolha_emissao = parsed[inicio_emissao:fim_emissao]
            brancos_emissao = escolha_emissao.replace(" ", "")
            data_emissao = brancos_emissao.replace("\n", "")

            posicao_forn = parsed.find("Código: ")
            inicio_forn = posicao_forn + 8
            fim_forn = inicio_forn + 5
            escolha_forn = parsed[inicio_forn:fim_forn]
            brancos_forn = escolha_forn.replace(" ", "")
            cod_fornecedor = brancos_forn.replace("\n", "")

            posicao_nome_forn = parsed.find("Fornecedor: ")
            inicio_nome_forn = posicao_nome_forn + 12
            fim_nome_forn = posicao_forn - 1
            escolha_nome_forn = parsed[inicio_nome_forn:fim_nome_forn]
            nome_fornecedor = escolha_nome_forn.replace("\n", "")

            inicio_frete = parsed.find("frete: ")
            fim_frete = parsed.find("Outras despesas:")
            frete_primeiro = inicio_frete + 7
            escolha_frete = parsed[frete_primeiro:fim_frete]
            brancos_frete = escolha_frete.replace(" ", "")
            valor_frete = brancos_frete.replace("\n", "")
            frete_ponto = valor_frete.replace(",", ".")
            frete_float = float(frete_ponto)

            inicio_desconto = parsed.find("Descontos: ")
            fim_desconto = parsed.find("Acrésc. Financeiro:")
            desconto_primeiro = inicio_desconto + 11
            escolha_desconto = parsed[desconto_primeiro:fim_desconto]
            brancos_desconto = escolha_desconto.replace(" ", "")
            valor_desconto = brancos_desconto.replace("\n", "")
            desconto_ponto = valor_desconto.replace(",", ".")
            desconto_float = float(desconto_ponto)
            self.dados_absolutos_itens()

            return numero_oc, data_emissao, cod_fornecedor, nome_fornecedor, frete_float, desconto_float

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "dados_cabecalho".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def dados_absolutos_itens(self):
        try:
            parsed = self.ler_pdf()

            if self.number_of_pages == 1:
                inicio_lista = parsed.find("ENTREGA\n")
                fim_lista = parsed.find("\nOBSERVAÇÕES:")
                lista_primeiro = inicio_lista + 8
                self.escolha_total = parsed[lista_primeiro:fim_lista]

                quantidade_de_itens = 0
                teste1 = self.escolha_total.split("\n")

                for cucu in teste1:
                    if len(cucu) > 22:
                        quantidade_de_itens = quantidade_de_itens + 1

                lista_original = [int(temp) for temp in self.escolha_total.split() if temp.isdigit()]

                lst = np.array(lista_original)

                lista_de_itens = [1]

                x = range(1, 28)
                for n in x:
                    result = np.where(lst == n)
                    if not result:
                        pass
                    else:
                        result1 = result[0]
                        list1 = result1.tolist()
                        if not list1:
                            pass
                        else:
                            if n == (len(lista_de_itens) + 1):
                                lista_de_itens.append(n)

                self.last = quantidade_de_itens

            else:
                inicio_lista = parsed.find("ENTREGA\n")
                fim_lista = parsed.find("OBSERVAÇÕES:")
                lista_primeiro = inicio_lista + 8
                escolha_primeiro = parsed[lista_primeiro:fim_lista]

                fim_escolha = escolha_primeiro.find(" SUZUKI")
                escolha_primeiro1 = escolha_primeiro[0:fim_escolha]

                comeco_escolha = escolha_primeiro.find("NOTA FISCAL\n")
                comeco_escolha1 = comeco_escolha + 12
                escolha_primeiro2 = escolha_primeiro[comeco_escolha1:-1]

                self.escolha_total = escolha_primeiro1 + '\n' + escolha_primeiro2

                lista_original = [int(temp) for temp in self.escolha_total.split() if temp.isdigit()]

                lst = np.array(lista_original)

                lista_de_itens = [1]

                x = range(1, 60)

                for n in x:
                    result = np.where(lst == n)
                    if not result:
                        pass
                    else:
                        result1 = result[0]
                        list1 = result1.tolist()
                        if not list1:
                            pass
                        else:
                            if n == (len(lista_de_itens) + 1):
                                lista_de_itens.append(n)

                self.last = lista_de_itens.pop()

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "dados_absolutos_itens".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def itens(self, num_item_comeco, num_item_final):
        try:
            posicaoinicio_umitem = self.escolha_total.find(num_item_comeco)
            posicaofim_umitem = self.escolha_total.find(num_item_final)

            inicio_umitem = posicaoinicio_umitem + (len(num_item_comeco))
            if num_item_final == "ultimo":
                escolha_umitem = self.escolha_total[inicio_umitem:]
            else:
                escolha_umitem = self.escolha_total[inicio_umitem:posicaofim_umitem]

            umitem = escolha_umitem.replace("\n", " ")

            numbers_umitem = [int(temp) for temp in umitem.split() if temp.isdigit()]

            quantidade_numbers_umitem = len(numbers_umitem)

            if quantidade_numbers_umitem == 1:
                self.codigo_item = numbers_umitem[0]
            else:
                for codigos in numbers_umitem:
                    cur = conecta.cursor()
                    cur.execute(f"SELECT codigo, descricao FROM produto where codigo = {codigos};")
                    extrair_dados = cur.fetchall()

                    if not extrair_dados:
                        self.codigo_item = numbers_umitem[0]
                    else:
                        codigo, descricao = extrair_dados[0]
                        output = descricao.split(' ')
                        output1 = output[0]
                        contar = umitem.count(output1)
                        if contar != 0:
                            self.codigo_item = int(codigo)

            posicao_inicio_codigo = umitem.find(f"{self.codigo_item}")
            posicao_fim_codigo = len(str(self.codigo_item)) + posicao_inicio_codigo

            carac = ' '
            lista_espacos = []
            for pos, char in enumerate(umitem):
                if char == carac:
                    lista_espacos.append(pos)
            qtdes_espacos = len(lista_espacos)

            item_13_espaco = qtdes_espacos - 1
            ultimo_espaco = lista_espacos[item_13_espaco]

            item_12_espaco = item_13_espaco - 1
            penultimo_espaco = lista_espacos[item_12_espaco]

            item_11_espaco = item_12_espaco - 1
            posicao_11_espaco = lista_espacos[item_11_espaco]

            item_10_espaco = item_11_espaco - 1
            posicao_10_espaco = lista_espacos[item_10_espaco]

            item_09_espaco = item_10_espaco - 1
            posicao_09_espaco = lista_espacos[item_09_espaco]

            item_08_espaco = item_09_espaco - 1
            posicao_08_espaco = lista_espacos[item_08_espaco]

            qtde_caracteres = len(umitem)
            inicio_data = ultimo_espaco + 1
            data_entrega = umitem[inicio_data:qtde_caracteres]

            inicio_ipi = penultimo_espaco + 1
            ipi = umitem[inicio_ipi:ultimo_espaco]

            inicio_valor_total = posicao_11_espaco + 1
            valor_total = umitem[inicio_valor_total:penultimo_espaco]

            inicio_valor_unit = posicao_10_espaco + 1
            valor_unit = umitem[inicio_valor_unit:posicao_11_espaco]

            inicio_unidade = posicao_09_espaco + 1
            um = umitem[inicio_unidade:posicao_10_espaco]

            inicio_qtde_item = posicao_08_espaco + 1
            qtde_item = umitem[inicio_qtde_item:posicao_09_espaco]

            inicio_descricao_item = posicao_fim_codigo + 1
            descricao_item = umitem[inicio_descricao_item:posicao_08_espaco]

            referencia_item = umitem[0:posicao_inicio_codigo]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.id, 'X'), COALESCE(prodreq.numero, 'X'), "
                           f"prod.codigo, prod.descricao as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then prodreq.referencia "
                           f"else prod.obs end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prodreq.status = 'A' AND prod.codigo = {self.codigo_item};")
            extrair_req = cursor.fetchall()
            if not extrair_req:
                id_req = "X"

            elif len(extrair_req) > 1:
                id_req = "X"
            else:
                id_req, num_req, codigo_item_req, descricao_req, referencia_req, um_req, qtde_req = extrair_req[0]

            lista_de_oc_um_item = [self.codigo_item, descricao_item, referencia_item, um, qtde_item, valor_unit,
                                   valor_total, ipi, data_entrega, id_req]
            return lista_de_oc_um_item

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "itens".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def detalhar_itens(self):
        try:
            tabela = []
            qtde_itens = range(1, self.last + 1)
            for item in qtde_itens:
                primeiro = item
                segundo = item + 1
                if item == self.last:
                    item_1 = self.itens(f"\n{primeiro} ", f"ultimo")
                else:
                    item_1 = self.itens(f"\n{primeiro} ", f"\n{segundo} ")
                tabela.append(item_1)

            self.lanca_tabela_oc(tabela)

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "detalhar_itens".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def imprimir_itens(self):
        try:
            self.total_ipi = 0
            self.total_mercadorias = 0
            self.total_itens = 0
            self.number_of_pages = 0
            self.last = 0
            self.escolha_total = ""
            self.lista_completa_req = []

            self.label_Sem_Cad_Forn.setHidden(True)
            self.btn_Cad_Fornc.setHidden(True)
            self.label_Sem_Cad_Prod.setHidden(True)
            self.btn_Cad_Produto.setHidden(True)
            self.Line_Nome_Fornc.setStyleSheet(f"background-color: {self.branco}")
            self.Line_Cod_Fornc.setStyleSheet(f"background-color: {self.branco}")

            self.table_Itens_OC.clearContents()
            numero_oc, data_emissao, cod_fornecedor, nome_fornecedor, frete, desconto = self.dados_cabecalho()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT * FROM ordemcompra where entradasaida = 'E' AND NUMERO = {numero_oc};")
            dados_oc = cursor.fetchall()
            if not dados_oc:
                self.Line_Numero_OC.setText(numero_oc)
                self.Line_Emissao_OC.setText(data_emissao)
                self.Line_Cod_Fornc.setText(cod_fornecedor)
                self.Line_Nome_Fornc.setText(nome_fornecedor)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
                dados_fornecedor = cursor.fetchall()
                if not dados_fornecedor:
                    self.label_Sem_Cad_Forn.setHidden(False)
                    self.btn_Cad_Fornc.setHidden(False)
                    self.Line_Nome_Fornc.setStyleSheet(f"background-color: {self.amarelo}")
                    self.Line_Cod_Fornc.setStyleSheet(f"background-color: {self.amarelo}")

                self.detalhar_itens()
                self.procura_requisicoes()
                self.lanca_tabela_req()

                total_ipi_string = str("%.2f" % self.total_ipi)
                total_mercadoria_string = str("%.2f" % self.total_mercadorias)
                total_geral = self.total_itens + frete - desconto
                total_geral_string = str("%.2f" % total_geral)
                frete_string = str("%.2f" % frete)
                desconto_string = str("%.2f" % desconto)

                self.Line_Frete.setText(frete_string)
                self.Line_Desconto.setText(desconto_string)
                self.Line_IPI.setText(total_ipi_string)
                self.Line_Mercadorias.setText(total_mercadoria_string)
                self.Line_Geral.setText(total_geral_string)

            else:
                messagebox.showinfo(title="Aviso!",
                                    message=f'A Ordem de Compra Nº {numero_oc} já foi lançada no Horus!')
                self.ler_pdf_excluir()
                self.encontra_arquivos()
                self.atualiza_combobox()

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "imprimir_itens".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def lanca_tabela_oc(self, estrutura):
        try:
            linhas = (len(estrutura))
            colunas = (len(estrutura[0]))
            self.table_Itens_OC.setRowCount(linhas)
            self.table_Itens_OC.setColumnCount(colunas)

            for i in range(0, linhas):
                self.table_Itens_OC.setRowHeight(i, 7)
                for j in range(0, colunas):
                    delegate = AlignDelegate(self.table_Itens_OC)
                    self.table_Itens_OC.setItemDelegateForColumn(j, delegate)
                    self.table_Itens_OC.setItem(i, j, QtWidgets.QTableWidgetItem(str(estrutura[i][j])))

            self.pintar_tabela_oc(estrutura)

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "lanca_tabela_oc".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def pintar_tabela_oc(self, estrutura):
        try:
            for index, itens in enumerate(estrutura):
                cod, descr, ref, um, qtde, unit, total, ipi, entrega, id_req = itens
                
                qtde_item1 = qtde.replace(',', '.')
                qtde_item_float = float(qtde_item1)

                valor_unit1 = unit.replace(',', '.')
                valor_unit_float = float(valor_unit1)

                ipi_item = ipi.replace(',', '.')
                ipi_item_float = float(ipi_item)

                total_sem_ipi = qtde_item_float * valor_unit_float
                valor_ipi = total_sem_ipi * (ipi_item_float / 100)
                total_com_ipi = total_sem_ipi + valor_ipi

                self.total_ipi = self.total_ipi + valor_ipi
                self.total_mercadorias = self.total_mercadorias + total_sem_ipi
                self.total_itens = self.total_itens + total_com_ipi

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao FROM produto where codigo = {cod};")
                dados_produto = cursor.fetchall()
                if not dados_produto:
                    self.label_Sem_Cad_Prod.setHidden(False)
                    self.btn_Cad_Produto.setHidden(False)
                    self.table_Itens_OC.item(index, 0).setBackground(QColor(self.amarelo))
                    self.table_Itens_OC.item(index, 1).setBackground(QColor(self.amarelo))
                    self.table_Itens_OC.item(index, 2).setBackground(QColor(self.amarelo))
                    self.table_Itens_OC.item(index, 3).setBackground(QColor(self.amarelo))
                if id_req == "X":
                    self.table_Itens_OC.item(index, 0).setBackground(QColor(self.amarelo))
                    self.table_Itens_OC.item(index, 1).setBackground(QColor(self.amarelo))
                    self.table_Itens_OC.item(index, 9).setBackground(QColor(self.amarelo))

        except Exception as e:
            print("pintar_tabela_oc", e)
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "pintar_tabela_oc".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            self.close()

    def procura_requisicoes(self):
        try:
            estrutura_oc = self.extrair_dados_tabela_oc()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.id, ''), COALESCE(prodreq.numero, ''), "
                           f"prod.codigo, prod.descricao as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                           f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prodreq.status = 'A' ORDER BY DESCRICAO;")
            extrair_req = cursor.fetchall()

            for itens in extrair_req:
                id_req, num_req, cod, descricao, ref, um, qtde = itens
                cod_int = int(cod)

                qtde_item_repetido = sum(map(lambda lista_n: lista_n.count(cod), extrair_req))

                item_encontrado = [s for s in estrutura_oc if cod_int in s]

                if qtde_item_repetido > 1 or not item_encontrado:
                    self.lista_completa_req.append(itens)

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "procura_requisicoes".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def lanca_tabela_req(self):
        try:
            estrutura = self.lista_completa_req
            if not estrutura:
                pass
            else:
                linhas = (len(estrutura))
                colunas = (len(estrutura[0]))
                self.table_Itens_Req.setRowCount(linhas)
                self.table_Itens_Req.setColumnCount(colunas)

                for i in range(0, linhas):
                    self.table_Itens_Req.setRowHeight(i, 7)
                    for j in range(0, colunas):
                        delegate = AlignDelegate(self.table_Itens_Req)
                        self.table_Itens_Req.setItemDelegateForColumn(j, delegate)
                        self.table_Itens_Req.setItem(i, j, QtWidgets.QTableWidgetItem(str(estrutura[i][j])))

                estrutura_oc = self.extrair_dados_tabela_oc()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT COALESCE(prodreq.id, ''), COALESCE(prodreq.numero, ''), "
                               f"prod.codigo, prod.descricao as DESCRICAO, "
                               f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                               f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                               f"prod.unidade, prodreq.quantidade "
                               f"FROM produtoordemrequisicao as prodreq "
                               f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                               f"WHERE prodreq.status = 'A' ORDER BY DESCRICAO;")
                extrair_req = cursor.fetchall()
                testinho = 0
                num_ativa_msg = 0
                for itens in extrair_req:
                    id_req, num_req, cod, descricao, ref, um, qtde = itens
                    cod_int = int(cod)

                    qtde_item_repetido = sum(map(lambda lista_n: lista_n.count(cod), extrair_req))

                    item_encontrado = [s for s in estrutura_oc if cod_int in s]

                    if not item_encontrado:
                        testinho = testinho + 1
                    else:
                        if qtde_item_repetido > 1:
                            testinho = testinho + 1
                            testinho2 = testinho - 1
                            num_ativa_msg = num_ativa_msg + 1

                            self.table_Itens_Req.item(testinho2, 0).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 1).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 2).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 3).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 4).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 5).setBackground(QColor(self.amarelo))
                            self.table_Itens_Req.item(testinho2, 6).setBackground(QColor(self.amarelo))

                if num_ativa_msg > 0:
                    messagebox.showinfo(title="Aviso!", message=f'A Ordem de Compra possui um ou mais itens '
                                                                f'com requisições duplicadas!')

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "lanca_tabela_req".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            telaordem.close()

    def atualiza_combobox(self):
        self.combo_arquivos.clear()
        index1 = self.index_arquivo_para_excluir
        self.arquivos_pdf.pop(index1)
        self.combo_arquivos.addItems(self.arquivos_pdf)

    def extrair_dados_tabela_oc(self):
        row_count = self.table_Itens_OC.rowCount()
        column_count = self.table_Itens_OC.columnCount()
        lista_final_itens = []
        linha = []
        for row in range(row_count):
            for column in range(column_count):
                widget_item = self.table_Itens_OC.item(row, column)
                lista_item = widget_item.text()
                linha.append(lista_item)
                if len(linha) == column_count:
                    lista_final_itens.append(linha)
                    linha = []
        return lista_final_itens

    def verifica_line_idreq(self, id_req_prod):
        try:
            if len(id_req_prod) == 0:
                vai_naovai = False

                messagebox.showinfo(title="Aviso!", message='O campo "ID Requis.:" não pode estar vazio')

            elif int(id_req_prod) == 0:
                vai_naovai = False

                messagebox.showinfo(title="Aviso!", message='O campo "ID Requis.:" não pode ser "0"')

            else:
                vai_naovai = self.verifica_sql_idreq(id_req_prod)

            return vai_naovai

        except Exception as e:
            print("verifica_line_idreq", e)
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "verifica_line_idreq".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n{e}')
            self.close()

    def verifica_sql_idreq(self, id_req_prod):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.id, 'X'), COALESCE(prodreq.numero, 'X'), "
                           f"prod.codigo, prod.descricao as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then prodreq.referencia "
                           f"else prod.obs end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prodreq.status = 'A' AND prodreq.id = {id_req_prod};")
            extrair_req = cursor.fetchall()
            if not extrair_req:
                vai_naovai = False

                messagebox.showinfo(title="Aviso!", message='Este "ID" do Produto da Requisição não existe!')
            else:
                vai_naovai = True

            return vai_naovai
        except Exception as e:
            print("verifica_sql_idreq", e)
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "verifica_sql_idreq".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n{e}')
            self.close()

    def verifica_salvamento(self):
        try:
            cod_fornecedor = self.Line_Cod_Fornc.text()
            nome_fornecedor = self.Line_Nome_Fornc.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()

            if not dados_fornecedor:
                messagebox.showinfo(title="Aviso!", message=f'O Fornecedor {nome_fornecedor} não está cadastrado!')
            else:
                testar_erros = 0
                dados_alterados = self.extrair_dados_tabela_oc()
                for itens in dados_alterados:
                    codigos_do_item, descricao_item, referencia_item, um, qtde_item, valor_unit, valor_total, \
                        ipi, data_entrega, id_req = itens

                    if id_req == "X":
                        messagebox.showinfo(title="Aviso!", message=f'Não pode haver produtos sem o "ID" da requisição '
                                                                    f'vinculado a ordem de compra!')
                        testar_erros += 1
                        break
                    else:
                        vai_naovai = self.verifica_line_idreq(id_req)

                        if vai_naovai:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigos_do_item};")
                            dados_produto = cursor.fetchall()
                            if not dados_produto:
                                messagebox.showinfo(title="Aviso!", message=f'O produto {descricao_item} não está '
                                                                            f'cadastrado')
                                testar_erros += 1
                                break

                            posicao_req = id_req.find(",")
                            if posicao_req > 0:
                                messagebox.showinfo(title="Aviso!", message=f'A Requisição do produto {descricao_item} '
                                                                            f'está com o número duplicado!')
                                testar_erros += 1
                                break
                        else:
                            testar_erros += 1
                    if testar_erros == 0:
                        self.salvar_ordem()

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "verifica_salvamento".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            self.close()

    def salvar_ordem(self):
        try:
            cod_fornecedor = self.Line_Cod_Fornc.text()

            numero_oc = self.Line_Numero_OC.text()
            numero_oc_int = int(numero_oc)

            emissao_oc = self.Line_Emissao_OC.text()
            emissao_oc1 = datetime.strptime(emissao_oc, '%d/%m/%Y').date()
            emissao_oc2 = str(emissao_oc1)
            emissao_oc_certo = "'" + emissao_oc2 + "'"

            frete_oc = self.Line_Frete.text()
            frete_oc_float = float(frete_oc)

            desconto_oc = self.Line_Desconto.text()
            desconto_oc_float = float(desconto_oc)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()
            id_fornecedor, razao = dados_fornecedor[0]

            cursor = conecta.cursor()
            cursor.execute("select GEN_ID(GEN_ORDEMCOMPRA_ID,0) from rdb$database;")
            ultimo_oc0 = cursor.fetchall()
            ultimo_oc1 = ultimo_oc0[0]
            ultimo_oc = int(ultimo_oc1[0]) + 1

            cursor = conecta.cursor()
            cursor.execute(f"Insert into ordemcompra "
                           f"(ID, ENTRADASAIDA, NUMERO, DATA, STATUS, FORNECEDOR, LOCALESTOQUE, FRETE, DESCONTOS) "
                           f"values (GEN_ID(GEN_ORDEMCOMPRA_ID,1), "
                           f"'E', {numero_oc_int}, {emissao_oc_certo}, 'A', {id_fornecedor}, '1', {frete_oc_float}, "
                           f"{desconto_oc_float});")

            dados_alterados = self.extrair_dados_tabela_oc()

            for itens in dados_alterados:
                codigo, descricao, referencia, um, qtde, rs_unit, rs, ipi, data_entregas, id_req = itens

                codigo_int = int(codigo)

                entrega_produto1 = datetime.strptime(data_entregas, '%d/%m/%Y').date()
                entrega_produto2 = str(entrega_produto1)
                entrega_produto = "'" + entrega_produto2 + "'"

                qtde_item = qtde.replace(',', '.')
                qtde_item_float = float(qtde_item)

                valor_unit = rs_unit.replace(',', '.')
                valor_unit_float = float(valor_unit)

                ipi_item = ipi.replace(',', '.')
                ipi_item_float = float(ipi_item)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigo_int};")
                dados_produto = cursor.fetchall()

                id_produto, descricao = dados_produto[0]

                id_req_int = int(id_req)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, mestre FROM produtoordemrequisicao "
                               f"WHERE id = {id_req_int};")
                update_req = cursor.fetchall()

                id_req, num_req_mestre = update_req[0]

                cursor = conecta.cursor()
                cursor.execute(f"Insert into produtoordemcompra (ID, MESTRE, PRODUTO, QUANTIDADE, UNITARIO, "
                               f"IPI, DATAENTREGA, NUMERO, CODIGO, PRODUZIDO, ID_PROD_REQ) "
                               f"values (GEN_ID(GEN_PRODUTOORDEMCOMPRA_ID,1), {ultimo_oc}, "
                               f"{id_produto}, {qtde_item_float}, {valor_unit_float}, {ipi_item_float}, "
                               f"{entrega_produto}, {numero_oc_int}, '{codigo}', 0.0, {id_req});")

                cursor = conecta.cursor()
                cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'B', "
                               f"PRODUZIDO = {qtde_item_float} WHERE id = {id_req};")

            conecta.commit()

            messagebox.showinfo(title="Aviso!", message=f'Ordem de Compra foi lançada com sucesso!')

            self.ler_pdf_excluir()
            self.atualiza_combobox()
            self.limpar_tudo()

        except Exception as e:
            messagebox.showinfo(title="Aviso!", message=f'Houve um problema com a função "salvar_ordem".\n'
                                                        f'Comunique o desenvolvedor sobre o erro abaixo:\n'
                                                        f'{e}')
            self.close()

    def cadastra_fornecedor(self):
        cod_fornecedor = self.Line_Cod_Fornc.text()
        nome_fornecedor = self.Line_Nome_Fornc.text()

        cursor = conecta.cursor()
        cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
        dados_fornecedor = cursor.fetchall()
        if not dados_fornecedor:
            cursor = conecta.cursor()
            cursor.execute(f"Insert into fornecedores (ID, RAZAO, REGISTRO) "
                           f"values (GEN_ID(GEN_FORNECEDORES_ID,1), '{nome_fornecedor}', '{cod_fornecedor}');")
            conecta.commit()
            messagebox.showinfo(title="Aviso!",
                                message=f'O Fornecedor {nome_fornecedor} foi cadastrado com sucesso!')

        cursor = conecta.cursor()
        cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
        dados_fornecedor = cursor.fetchall()
        if not dados_fornecedor:
            self.label_Sem_Cad_Forn.setHidden(False)
            self.btn_Cad_Fornc.setHidden(False)
            self.Line_Nome_Fornc.setStyleSheet(f"background-color: {self.amarelo}")
            self.Line_Cod_Fornc.setStyleSheet(f"background-color: {self.amarelo}")
        else:
            self.label_Sem_Cad_Forn.setHidden(True)
            self.btn_Cad_Fornc.setHidden(True)
            self.Line_Nome_Fornc.setStyleSheet(f"background-color: {self.branco}")
            self.Line_Cod_Fornc.setStyleSheet(f"background-color: {self.branco}")

    def cadastra_produto(self):
        dados_alterados = self.extrair_dados_tabela_oc()

        testinho = 0

        for itens in dados_alterados:
            codigos_do_item, descricao_item, referencia_item, um, qtde_item, valor_unit, valor_total, \
                ipi, data_entrega, id_req, num_req = itens

            material_msg = codigos_do_item + " - " + descricao_item + " - " + referencia_item

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigos_do_item};")
            dados_produto = cursor.fetchall()
            if not dados_produto:
                cursor = conecta.cursor()
                cursor.execute(f"Insert into produto (ID, CODIGO, CONJUNTO, DESCRICAO, UNIDADE, "
                               f"QUANTIDADE, CUSTOUNITARIO, OBS, LOCALIZACAO, CUSTOESTRUTURA) "
                               f"values (GEN_ID(GEN_PRODUTO_ID,1), '{codigos_do_item}', 8, "
                               f"'{descricao_item}', '{um}', 0.00, 0.00, '{referencia_item}', 'A', 0.00);")
                conecta.commit()
                messagebox.showinfo(title="Aviso!", message=f'O produto {material_msg} foi cadastrado com sucesso!')

            testinho = testinho + 1
            testinho2 = testinho - 1
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigos_do_item};")
            dados_produto = cursor.fetchall()
            if not dados_produto:
                self.label_Sem_Cad_Prod.setHidden(False)
                self.btn_Cad_Produto.setHidden(False)
                self.table_Itens_OC.item(testinho2, 0).setBackground(QColor(self.amarelo))
                self.table_Itens_OC.item(testinho2, 1).setBackground(QColor(self.amarelo))
                self.table_Itens_OC.item(testinho2, 2).setBackground(QColor(self.amarelo))
                self.table_Itens_OC.item(testinho2, 3).setBackground(QColor(self.amarelo))
            else:
                self.label_Sem_Cad_Prod.setHidden(True)
                self.btn_Cad_Produto.setHidden(True)
                self.table_Itens_OC.item(testinho2, 0).setBackground(QColor(self.branco))
                self.table_Itens_OC.item(testinho2, 1).setBackground(QColor(self.branco))
                self.table_Itens_OC.item(testinho2, 2).setBackground(QColor(self.branco))
                self.table_Itens_OC.item(testinho2, 3).setBackground(QColor(self.branco))

    def limpar_tudo(self):
        self.Line_Numero_OC.clear()
        self.Line_Emissao_OC.clear()
        self.Line_Cod_Fornc.clear()
        self.Line_Nome_Fornc.clear()
        self.Line_Frete.clear()
        self.Line_Desconto.clear()
        self.Line_IPI.clear()
        self.Line_Mercadorias.clear()
        self.Line_Geral.clear()

        self.table_Itens_OC.setRowCount(0)
        self.table_Itens_Req.setRowCount(0)

        self.label_Sem_Cad_Forn.setHidden(True)
        self.btn_Cad_Fornc.setHidden(True)

        self.label_Sem_Cad_Prod.setHidden(True)
        self.btn_Cad_Produto.setHidden(True)

        self.codigo_item = 0
        self.last = 0
        self.number_of_pages = 0
        self.escolha_total = ""

        self.lista_completa_req = []
        self.index_arquivo_para_excluir = 0
        self.arquivo_para_excluir = []

        self.total_ipi = 0
        self.total_mercadorias = 0
        self.total_itens = 0


class AlignDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = QtCore.Qt.AlignCenter


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    telaordem = TelaOrdem()
    telaordem.show()
    sys.exit(qt.exec_())
