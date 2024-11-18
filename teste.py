import sys
from banco_dados.conexao import conecta
from banco_dados.controle_erros import grava_erro_banco
import os
from PIL import Image, ImageDraw, ImageFont
import inspect
import traceback
from pathlib import Path


class EnviaOrdensProducao:
    def __init__(self):
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.caminho_original = ""
        self.arq_original = ""
        self.num_desenho_arq = ""
        self.qtde_produto = 0
        self.cod_prod = ""
        self.descr_prod = ""
        self.ref_prod = ""
        self.um_prod = ""
        self.num_op = ""
        self.tipo = ""

        self.data_emissao = ""
        self.data_entrega = ""

        self.arquivos_pra_excluir = []

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

    def criar_texto(self, imgs, pos_horizontal, pos_vertical, texto, cor, fonte, largura_tra):
        try:
            draw = ImageDraw.Draw(imgs)
            draw.text((pos_horizontal, pos_vertical), texto, fill=cor, font=fonte, stroke_width=largura_tra)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_imagem_da_ficha1(self, cod, descr, ref, um, local):
        try:
            imgs = Image.open("ficha_prod_modelo.png")

            font_cod = ImageFont.truetype("tahoma.ttf", 100)
            font_descricao = ImageFont.truetype("tahoma.ttf", 60)

            self.criar_texto(imgs, 1510, 130, cod, (0, 0, 0), font_cod, 3)
            self.criar_texto(imgs, 680, 335, descr, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 615, 455, ref, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 1700, 455, um, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 645, 575, local, (0, 0, 0), font_descricao, 0)

            arquivo_final = f"ficha_produto_{cod}.png"
            imgs.save(arquivo_final)

            self.arquivos_pra_excluir.append(arquivo_final)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_imagem_da_ficha2(self, imagem_ficha1, cod, descr, ref, um, local):
        try:
            imgs = Image.open(imagem_ficha1)

            font_cod = ImageFont.truetype("tahoma.ttf", 100)
            font_descricao = ImageFont.truetype("tahoma.ttf", 60)

            self.criar_texto(imgs, 3050, 130, cod, (0, 0, 0), font_cod, 3)
            self.criar_texto(imgs, 2220, 335, descr, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 2155, 455, ref, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 3240, 455, um, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 2185, 575, local, (0, 0, 0), font_descricao, 0)

            arquivo_final = f"ficha_produto_{cod}.png"
            imgs.save(arquivo_final)

            self.arquivos_pra_excluir.append(arquivo_final)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def converte_varias_imagens_para_pdf(self, imagens_fichas):
        try:
            pil_images = [Image.open(imagem).convert("RGB") for imagem in imagens_fichas]

            desktop = Path.home() / "Desktop"
            nome_req = '\Ficha_final.pdf'
            caminho = str(desktop) + nome_req

            pil_images[0].save(caminho, save_all=True, append_images=pil_images[1:], format="PDF", resolution=100.0)

            print(f"PDF com várias páginas salvo em {caminho}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_comeco(self, lista_codigos_produtos):
        try:
            imagens_fichas = []

            for i in range(0, len(lista_codigos_produtos), 2):
                codigo_produto1 = lista_codigos_produtos[i]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.DESCRICAOCOMPLEMENTAR, ''), "
                               f"COALESCE(prod.obs, ''), "
                               f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, prod.localizacao, "
                               f"prod.unidade, prod.quantidade "
                               f"FROM produto as prod "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where codigo = {codigo_produto1};")
                detalhes_produto1 = cursor.fetchone()
                cod1, descr1, compl1, ref1, ncm1, conjunto1, local1, um1, saldo1 = detalhes_produto1

                imagem_ficha1 = self.cria_imagem_da_ficha1(cod1, descr1, ref1, um1, local1)

                if i + 1 < len(lista_codigos_produtos):
                    codigo_produto2 = lista_codigos_produtos[i + 1]

                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.DESCRICAOCOMPLEMENTAR, ''), "
                                   f"COALESCE(prod.obs, ''), "
                                   f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, prod.localizacao, "
                                   f"prod.unidade, prod.quantidade "
                                   f"FROM produto as prod "
                                   f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                                   f"where codigo = {codigo_produto2};")
                    detalhes_produto2 = cursor.fetchone()
                    cod2, descr2, compl2, ref2, ncm2, conjunto2, local2, um2, saldo2 = detalhes_produto2

                    imagem_ficha2 = self.cria_imagem_da_ficha2(imagem_ficha1, cod2, descr2, ref2, um2, local2)
                    imagens_fichas.append(imagem_ficha2)
                else:
                    imagens_fichas.append(imagem_ficha1)

            self.converte_varias_imagens_para_pdf(imagens_fichas)

            if self.arquivos_pra_excluir:
                for arqs in self.arquivos_pra_excluir:
                    print(arqs)
                    self.excluir_arquivo(arqs)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


lista_codigos_produtoss = [16334, 15295, 11981, 17188, 56391, 74116]

chama_classe = EnviaOrdensProducao()
chama_classe.manipula_comeco(lista_codigos_produtoss)
