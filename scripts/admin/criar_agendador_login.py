import subprocess
import os
import ctypes

print("Usuário:", os.getlogin())
print("Administrador:", ctypes.windll.shell32.IsUserAnAdmin())

USUARIO = "Anderson"

NOME_TAREFA = "Notificar_Login"

PYTHONW = r"C:\Users\Anderson\PycharmProjects\robo_boby\.venv\Scripts\pythonw.exe"
SCRIPT = r"C:\Users\Anderson\PycharmProjects\robo_boby\scripts\admin\login_windows.py"


def deletar_tarefa():
    subprocess.run(
        ["schtasks", "/delete", "/tn", NOME_TAREFA, "/f"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def criar_tarefa():
    cmd = [
        "schtasks",
        "/create",
        "/tn", NOME_TAREFA,
        "/tr", f'"{PYTHONW}" "{SCRIPT}"',
        "/sc", "ONLOGON",
        "/f",
    ]

    print(cmd)

    resultado = subprocess.run(
        cmd,
        capture_output=True,
        text=True
    )

    print("Código:", resultado.returncode)
    print("Saída:")
    print(resultado.stdout)
    print("Erro:")
    print(resultado.stderr)


if __name__ == "__main__":
    try:
        print("Removendo tarefa antiga...")
        deletar_tarefa()

        print("Criando tarefa...")
        criar_tarefa()

        print("✅ Tarefa criada com sucesso!")

    except subprocess.CalledProcessError as e:
        print(e)