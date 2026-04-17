import subprocess
from typing import TypedDict, List, Literal


class AgendamentoConfig(TypedDict):
    tipo: Literal["diario", "semanal", "mensal"]
    dias_semana: List[str]
    dia_mes: int

# ==============================
# CONFIGURAÇÕES GERAIS
# ==============================

USUARIO = "Anderson"

nome_arquivo = "plano_pcp"
HORARIO = "11:30"
sub_pasta = "pcp"

# IMPORTANTE: ":" não pode no nome da tarefa
NOME_TAREFA = f"{HORARIO.replace(':', '_')}_{nome_arquivo}"

PYTHONW = r"C:\Users\Anderson\PycharmProjects\robo_boby\.venv\Scripts\pythonw.exe"
SCRIPT = rf"C:\Users\Anderson\PycharmProjects\robo_boby\scripts\{sub_pasta}\{nome_arquivo}.py"


# ==============================
# CONFIGURAÇÃO DE AGENDAMENTO
# ==============================

"""
COMO USAR:

1) DIÁRIO
    "tipo": "diario"

2) SEMANAL
    "tipo": "semanal"
    "dias_semana": ["MON", "WED", "FRI"]

    Dias válidos:
        MON = segunda
        TUE = terça
        WED = quarta
        THU = quinta
        FRI = sexta
        SAT = sábado
        SUN = domingo

3) MENSAL
    "tipo": "mensal"
    "dia_mes": 10   # dia do mês (1 a 31)
"""

AGENDAMENTO: AgendamentoConfig = {
    "tipo": "diario", # "diario", "semanal", "mensal"
    "dias_semana": ["TUE"], # usado só no semanal
    "dia_mes": 1, # usado só no mensal
}


# ==============================
# FUNÇÕES
# ==============================

def deletar_tarefa():
    subprocess.run(
        ["schtasks", "/delete", "/tn", NOME_TAREFA, "/f"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

def criar_tarefa():
    cmd = [
        "schtasks",
        "/create",
        "/tn", NOME_TAREFA,
        "/tr", f'"{PYTHONW}" "{SCRIPT}"',
        "/st", HORARIO,
        "/it",
        "/f",
        "/rl", "highest",
    ]

    tipo = AGENDAMENTO["tipo"]

    if tipo == "diario":
        cmd.extend(["/sc", "daily"])

    elif tipo == "semanal":
        cmd.extend([
            "/sc", "weekly",
            "/d", ",".join(AGENDAMENTO["dias_semana"])
        ])

    elif tipo == "mensal":
        cmd.extend([
            "/sc", "monthly",
            "/d", str(AGENDAMENTO["dia_mes"])
        ])

    subprocess.run(cmd, check=True)

    print(cmd)


# ==============================
# EXECUÇÃO
# ==============================

if __name__ == "__main__":
    try:
        print("Removendo tarefa antiga (se existir)...")
        deletar_tarefa()

        print("Criando nova tarefa...")
        criar_tarefa()

        print("✅ Tarefa criada com sucesso!")

    except subprocess.CalledProcessError as e:
        print("❌ Erro ao criar tarefa:")
        print(e)