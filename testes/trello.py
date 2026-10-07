import requests


class Trello:
    def __init__(self):
        self.base_url = "https://api.trello.com/1"
        self.api_key = "9183e9c4f9148d7cc9165f3bf0e56256"
        self.token = "0dee357803e80467cc00a92fb2ccb502c26325affe9d9d5ea2a38dfd3165aa4a"

    def _request(self, method, endpoint, params=None, data=None):
        url = f"{self.base_url}{endpoint}"

        params = params or {}
        params["key"] = self.api_key
        params["token"] = self.token

        response = requests.request(
            method=method,
            url=url,
            params=params,
            json=data
        )

        response.raise_for_status()

        if response.text:
            return response.json()

        return None

    def listar_quadros(self):
        """
        Retorna todos os quadros aos quais o usuário tem acesso.
        """
        return self._request(
            "GET",
            "/members/me/boards"
        )

    def buscar_quadro(self, quadro_id):
        """
        Retorna os dados de um quadro.
        """
        return self._request(
            "GET",
            f"/boards/{quadro_id}"
        )

    def listar_listas(self, quadro_id):
        """
        Retorna todas as listas de um quadro.
        """
        return self._request(
            "GET",
            f"/boards/{quadro_id}/lists"
        )

    def buscar_lista(self, lista_id):
        """
        Retorna os dados de uma lista.
        """
        return self._request(
            "GET",
            f"/lists/{lista_id}"
        )

    def listar_cartoes(self, lista_id):
        """
        Retorna todos os cartões de uma lista.
        """
        return self._request(
            "GET",
            f"/lists/{lista_id}/cards"
        )

    def buscar_cartao(self, cartao_id):
        """
        Retorna os dados de um cartão.
        """
        return self._request(
            "GET",
            f"/cards/{cartao_id}"
        )

    def criar_cartao(self, lista_id, nome, descricao=None, posicao="bottom"):
        """
        Cria um novo cartão em uma lista.
        """

        data = {
            "idList": lista_id,
            "name": nome,
            "pos": posicao
        }

        if descricao:
            data["desc"] = descricao

        return self._request(
            "POST",
            "/cards",
            data=data
        )

    def alterar_cartao(self, cartao_id, nome=None, descricao=None, lista_id=None, posicao=None):
        """
        Altera os dados de um cartão.

        Só altera os campos informados.
        """

        data = {}

        if nome is not None:
            data["name"] = nome

        if descricao is not None:
            data["desc"] = descricao

        if lista_id is not None:
            data["idList"] = lista_id

        if posicao is not None:
            data["pos"] = posicao

        if not data:
            raise ValueError("Nenhuma alteração foi informada.")

        return self._request(
            "PUT",
            f"/cards/{cartao_id}",
            data=data
        )

    def mover_cartao(self, cartao_id, lista_id, posicao="bottom"):
        """
        Move um cartão para outra lista.
        """

        return self.alterar_cartao(
            cartao_id=cartao_id,
            lista_id=lista_id,
            posicao=posicao
        )

    def excluir_cartao(self, cartao_id):
        """
        Exclui definitivamente um cartão.
        """

        return self._request(
            "DELETE",
            f"/cards/{cartao_id}"
        )

    def adicionar_comentario(self, cartao_id, texto):
        """
        Adiciona um comentário ao cartão.
        """

        return self._request(
            "POST",
            f"/cards/{cartao_id}/actions/comments",
            data={
                "text": texto
            }
        )

    def listar_etiquetas_quadro(self, quadro_id):
        """
        Lista as etiquetas disponíveis no quadro.
        """

        return self._request(
            "GET",
            f"/boards/{quadro_id}/labels"
        )

    def adicionar_etiqueta(self, cartao_id, etiqueta_id):
        """
        Adiciona uma etiqueta ao cartão.
        """

        return self._request(
            "POST",
            f"/cards/{cartao_id}/idLabels",
            data={
                "value": etiqueta_id
            }
        )

    def remover_etiqueta(self, cartao_id, etiqueta_id):
        """
        Remove uma etiqueta do cartão.
        """

        return self._request(
            "DELETE",
            f"/cards/{cartao_id}/idLabels/{etiqueta_id}"
        )

    def listar_membros_quadro(self, quadro_id):
        """
        Lista os membros do quadro.
        """

        return self._request(
            "GET",
            f"/boards/{quadro_id}/members"
        )

    def adicionar_membro(self, cartao_id, membro_id):
        """
        Adiciona um membro ao cartão.
        """

        return self._request(
            "POST",
            f"/cards/{cartao_id}/idMembers",
            data={
                "value": membro_id
            }
        )

    def remover_membro(self, cartao_id, membro_id):
        """
        Remove um membro do cartão.
        """

        return self._request(
            "DELETE",
            f"/cards/{cartao_id}/idMembers/{membro_id}"
        )

    def criar_checklist(self, cartao_id, nome):
        """
        Cria um checklist dentro do cartão.
        """

        return self._request(
            "POST",
            "/checklists",
            data={
                "idCard": cartao_id,
                "name": nome
            }
        )

    def listar_checklists(self, cartao_id):
        """
        Lista os checklists do cartão.
        """

        return self._request(
            "GET",
            f"/cards/{cartao_id}/checklists"
        )

    def adicionar_item_checklist(self, checklist_id, nome, posicao="bottom"):
        """
        Adiciona um item ao checklist.
        """

        return self._request(
            "POST",
            f"/checklists/{checklist_id}/checkItems",
            data={
                "name": nome,
                "pos": posicao
            }
        )

    def marcar_item_checklist(self, cartao_id, item_id, marcado=True):
        """
        Marca ou desmarca um item do checklist.
        """

        estado = "complete" if marcado else "incomplete"

        return self._request(
            "PUT",
            f"/cards/{cartao_id}/checkItem/{item_id}",
            data={
                "state": estado
            }
        )

    def adicionar_anexo_url(self, cartao_id, url, nome=None):
        params = {
            "url": url
        }

        if nome:
            params["name"] = nome

        return self._request(
            "POST",
            f"/cards/{cartao_id}/attachments",
            params=params
        )

    def listar_acoes_cartao(self, cartao_id):
        """
        Retorna o histórico de ações do cartão.
        """
        return self._request(
            "GET",
            f"/cards/{cartao_id}/actions"
        )

    def listar_anexos_cartao(self, cartao_id):
        return self._request(
            "GET",
            f"/cards/{cartao_id}/attachments"
        )

    def excluir_anexo(self, cartao_id, anexo_id):
        return self._request(
            "DELETE",
            f"/cards/{cartao_id}/attachments/{anexo_id}"
        )