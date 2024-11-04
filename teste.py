class TaskManager:
    def __init__(self):
        self.tasks = []  # Inicializa uma lista vazia para armazenar as tarefas

    def add_task(self, task):
        """Adiciona uma nova tarefa à lista."""
        self.tasks.append(task)
        print(f'Tarefa "{task}" adicionada.')

    def remove_task(self, task):
        """Remove uma tarefa da lista."""
        if task in self.tasks:
            self.tasks.remove(task)
            print(f'Tarefa "{task}" removida.')
        else:
            print(f'Tarefa "{task}" não encontrada.')

    def list_tasks(self):
        """Lista todas as tarefas."""
        if not self.tasks:
            print("Nenhuma tarefa encontrada.")
        else:
            print("Tarefas:")
            for idx, task in enumerate(self.tasks, start=1):
                print(f"{idx}. {task}")

def main():
    task_manager = TaskManager()  # Cria uma instância do gerenciador de tarefas

    while True:
        print("\nMenu:")
        print("1. Adicionar Tarefa")
        print("2. Remover Tarefa")
        print("3. Listar Tarefas")
        print("4. Sair")
        choice = input("Escolha uma opção: ")

        if choice == '1':
            task = input("Digite a tarefa a ser adicionada: ")
            task_manager.add_task(task)
        elif choice == '2':
            task = input("Digite a tarefa a ser removida: ")
            task_manager.remove_task(task)
        elif choice == '3':
            task_manager.list_tasks()
        elif choice == '4':
            print("Saindo...")
            break
        else:
            print("Opção inválida! Tente novamente.")

if __name__ == "__main__":
    main()
