import subprocess

# ==============================
# CONFIGURAÇÕES
# ==============================

NOME_TAREFA = "07_35_data_encerrar.py"

USUARIO = "Anderson"  # usuário do Windows

PYTHONW = r"C:\Users\Anderson\PycharmProjects\robo_boby\.venv\Scripts\pythonw.exe"
SCRIPT = r"C:\Users\Anderson\PycharmProjects\robo_boby\data_encerrar.py"

HORARIO = "07:35"


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
    comando_execucao = f'"{PYTHONW}" "{SCRIPT}"'

    cmd = [
        "schtasks",
        "/create",
        "/tn", NOME_TAREFA,
        "/tr", comando_execucao,
        "/sc", "daily",
        "/st", HORARIO,
        "/ru", USUARIO,
        "/f"
    ]

    subprocess.run(cmd, check=True)


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