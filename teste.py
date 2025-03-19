import random


# Função para treinar a IA (perceptron)
def treinar_ia():
    # Inicializando pesos e bias aleatórios
    peso1, peso2, bias = random.random(), random.random(), random.random()

    # Taxa de aprendizado
    taxa_aprendizado = 0.1

    # Dados de entrada e o resultado esperado (saída)
    entradas = [
        [0, 0],  # Não está chovendo, não tem guarda-chuva
        [0, 1],  # Não está chovendo, tem guarda-chuva
        [1, 0],  # Está chovendo, não tem guarda-chuva
        [1, 1]  # Está chovendo, tem guarda-chuva
    ]
    saidas = [1, 1, 0, 1]  # O que a IA deve aprender: [Sim, Sim, Não, Sim]

    # Vamos treinar a IA por várias iterações (tentativas)
    for _ in range(1000):  # Número de tentativas (épocas)
        for i in range(len(entradas)):
            # Calculando a soma ponderada das entradas
            soma = entradas[i][0] * peso1 + entradas[i][1] * peso2 + bias
            # Função de ativação: se a soma for maior que 0, o resultado é 1 (Sim), caso contrário, 0 (Não)
            resultado = 1 if soma > 0 else 0

            # Calculando o erro (quanto a IA errou)
            erro = saidas[i] - resultado

            # Ajustando os pesos e o bias com base no erro
            peso1 += taxa_aprendizado * erro * entradas[i][0]
            peso2 += taxa_aprendizado * erro * entradas[i][1]
            bias += taxa_aprendizado * erro

    return peso1, peso2, bias


# Função para testar a IA (verificar se ela aprendeu corretamente)
def testar_ia(peso1, peso2, bias):
    entradas = [
        [0, 0],  # Não está chovendo, não tem guarda-chuva
        [0, 1],  # Não está chovendo, tem guarda-chuva
        [1, 0],  # Está chovendo, não tem guarda-chuva
        [1, 1]  # Está chovendo, tem guarda-chuva
    ]

    for i in range(len(entradas)):
        soma = entradas[i][0] * peso1 + entradas[i][1] * peso2 + bias
        resultado = 1 if soma > 0 else 0  # 1 = Sim, 0 = Não
        print(
            f"Está chovendo? {entradas[i][0]} | Tem guarda-chuva? {entradas[i][1]} -> Devo sair de casa? {'Sim' if resultado == 1 else 'Não'}")


# Treinando a IA
peso1, peso2, bias = treinar_ia()

# Testando a IA
testar_ia(peso1, peso2, bias)
