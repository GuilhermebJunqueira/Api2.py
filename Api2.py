import openpyxl

# Lista para armazenar os perfis dos alunos
alunos = []

# Função para exibir perfis de alunos
def exibir_alunos():
    for i, aluno in enumerate(alunos, start=1):
        print(f"Aluno {i}:")
        for chave, valor in aluno.items():
            print(f"{chave}: {valor}")
        print()

# Função para ler dados do Excel e adicionar aos perfis dos alunos
def ler_dados_excel():
    try:
        arquivo_excel = input("Digite o nome do arquivo Excel (.xlsx) com os dados: ")
        workbook = openpyxl.load_workbook(arquivo_excel)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            nome, matricula, curso, idade, grupos, membros_grupos, prazos_entrega = row
            aluno = {
                'Nome': nome,
                'Matrícula': matricula,
                'Curso': curso,
                'Idade': idade,
                'Grupos': grupos.split(','),
                'Membros dos Grupos': membros_grupos.split(','),
                'Prazos de Entrega': prazos_entrega.split(',')
            }
            alunos.append(aluno)

        print(f"Dados do arquivo '{arquivo_excel}' lidos e adicionados com sucesso!")

    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo Excel: {e}")

# Loop principal do programa
while True:
    print("Opções:")
    print("1 - Exibir perfis de alunos")
    print("2 - Ler dados de um arquivo Excel")
    print("3 - Sair")

    opcao = input("Escolha uma opção: ")

    if opcao == '1':
        exibir_alunos()
    elif opcao == '2':
        ler_dados_excel()
    elif opcao == '3':
        print("Saindo do programa.")
        break
    else:
        print("Opção inválida. Tente novamente.")
