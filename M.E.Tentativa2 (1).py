"""Imports"""
import pandas

"""Verificando arquivo"""
try:
    arquivo_prof = pandas.read_excel("Professores.xlsx")
except FileNotFoundError:
    coluna_prof = ["CNTPS", "NOME", "DATA DE NASCIMENTO"]

try:
    arquivo_alunos = pandas.read_excel("Alunos.xlsx")
except FileNotFoundError:
    coluna_alunos = ["MATRICULA", "NOME", "MATRICULA DO PROFESSOR"]

try:
    arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
except FileNotFoundError:
    coluna_discipli = ["CÓDIGO", "NOME", "MATRICULA DO PROFESSOR"]

try:
    arquivo_medias = pandas.read_excel("Medias.xlsx")
except FileNotFoundError:
    coluna_medias = ["CÓDIGO", "MATRICULA DO ALUNO", "MEDIA"]

try:
    arquivo_faltas = pandas.read_excel("Faltas.xlsx")
except FileNotFoundError:
    coluna_faltas = ["CÓDIGO", "MATRICULA DO ALUNO", "FALTAS"]


"""Criando Funções de cadastro"""

def red_message(message):
    print(f"\033[1;30;101m{message}\033[1;0;0m")


def cadastro_prof():
    while True:
        matricula = input("Informe a matricula do professor: ")
        try:
            arquivo_prof = pandas.read_excel("Professores.xlsx")
            if int(matricula) not in arquivo_prof["MATRICULA"].values:
                nome = str.upper(input("Informe o nome do professor: "))
                data = input("Informe a data de nascimento do professor: ")
                arquivo_prof.loc[len(arquivo_prof)] = [matricula, nome, data]
                arquivo_prof.to_excel("Professores.xlsx", index=False)
            else:
                print("Matricula já cadastrada.")
        except FileNotFoundError:
            nome = str.upper(input("Informe o nome do professor: "))
            data = input("Informe a data de nascimento do professor: ")
            matriculas = []
            nomes = []
            datas = []
            matriculas.append(matricula)
            nomes.append(nome)
            datas.append(data)
            arquivo_prof = pandas.DataFrame(list(zip(matriculas, nomes, datas)), columns=coluna_prof)
            arquivo_prof.to_excel("Professores.xlsx", index=False)

        except ValueError:
            red_message("Valor informado não é válido.")

        vontade = str.upper(input("Cadastrar outro professor?(s/n): "))
        if vontade == "N":
            break
        elif vontade != "S":
            print("Comando não compreendido")


def cadastro_aluno():
    while True:
        matricula = input("Informe a matricula do aluno: ")
        try:
            arquivo_aluno = pandas.read_excel("Alunos.xlsx")
            if int(matricula) not in arquivo_aluno["MATRICULA"].values:
                nome = str.upper(input("Informe o nome do aluno: "))
                data = input("Informe a data de nascimento do aluno: ")
                arquivo_aluno.loc[len(arquivo_aluno)] = [matricula, nome, data]
                arquivo_aluno.to_excel("Alunos.xlsx", index=False)
            else:
                red_message("Matricula já cadastrada.")
        except FileNotFoundError:
            nome = str.upper(input("Informe o nome do aluno: "))
            data = input("Informe a data de nascimento do aluno: ")
            matriculas = []
            nomes = []
            datas = []
            matriculas.append(matricula)
            nomes.append(nome)
            datas.append(data)
            arquivo_aluno = pandas.DataFrame(list(zip(matriculas, nomes, datas)), columns=coluna_alunos)
            arquivo_aluno.to_excel("Alunos.xlsx", index=False)

        except ValueError:
            red_message("Valor informado não é válido.")

        vontade = str.upper(input("Cadastrar outro aluno? (s/n): "))
        if vontade == "N":
            break
        elif vontade != "S":
            print("Comando não compreendido")


def cadastrar_disciplina():
    while True:
        try:
            arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
            try:
                arquivo_prof = pandas.read_excel("Professores.xlsx")
                codigo = input("Informe o código da disciplina: ")
                if int(codigo) not in arquivo_discipli["CÓDIGO"].values:
                    nome = str.upper(input("Informe o nome da disciplina: "))
                    matricula = input("Informe a matricula do professor ministrante: ")
                    if int(matricula) in arquivo_prof["MATRICULA"].values:
                        arquivo_discipli.loc[len(arquivo_discipli)] = [codigo, nome, matricula]
                        arquivo_discipli.to_excel("Disciplinas.xlsx", index=False)

                    else:
                        red_message("Professor não cadastrado.")

                else:
                    red_message("Disciplina já cadastrada.")

            except FileNotFoundError:
                red_message("Nenhum professor cadastrado")
                break

        except FileNotFoundError:
            try:
                arquivo_prof = pandas.read_excel("Professores.xlsx")
                codigo = input("Informe o código da disciplina: ")
                nome = str.upper(input("Informe o nome da disciplina: "))
                matricula = input("Informe a matricula do professor ministrante: ")

                if int(matricula) in arquivo_prof["MATRICULA"].values:
                    codigos = []
                    nomes = []
                    matri_prof = []
                    codigos.append(codigo)
                    nomes.append(nome)
                    matri_prof.append(matricula)
                    arquivo_alunos = pandas.DataFrame(list(zip(codigos, nomes, matri_prof)), columns=coluna_discipli)
                    arquivo_alunos.to_excel("Disciplinas.xlsx", index=False)

            except FileNotFoundError:
                red_message("Não existem professores cadastrados.")
                break

        except ValueError:
            red_message("Valor informado não é válido.")

        vontade = str.upper(input("Cadastrar outra disciplina?(s/n): "))
        if vontade == "N":
            break
        elif vontade != "S":
            print("Comando não compreendido")


def cadastrar_medias():
    while True:
        try:
            arquivo_medias = pandas.read_excel("Medias.xlsx")
            try:
                arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
                codigo = input("Informe o código da disciplina: ")
                if int(codigo) in arquivo_discipli["CÓDIGO"].values:
                    try:
                        arquivo_alunos = pandas.read_excel("Alunos.xlsx")
                        matricula = str.upper(input("Informe a matricula do aluno: "))
                        if int(matricula) in arquivo_alunos["MATRICULA"].values:
                            media = input("Informe a média do aluno: ")
                            if 0 <= int(media) <= 10:
                                arquivo_medias.loc[len(arquivo_medias)] = [codigo, matricula, media]
                                arquivo_medias.to_excel("Medias.xlsx", index=False)
                            else:
                                red_message("Média não pode ser menor que 0 ou maior que 10")
                        else:
                            red_message("Aluno não cadastrado.")
                    except FileNotFoundError:
                        red_message("Nenhum aluno cadastrado.")
                        break
                else:
                    red_message("Disciplina não cadastrada.")
            except FileNotFoundError:
                red_message("Nenhuma disciplina cadastrada.")
                break
        except FileNotFoundError:
            try:
                arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
                codigo = input("Informe o código da disciplina: ")
                if int(codigo) in arquivo_discipli["CÓDIGO"].values:
                    try:
                        arquivo_alunos = pandas.read_excel("Alunos.xlsx")
                        matricula = input("Informe a matricula do aluno: ")
                        if int(matricula) in arquivo_alunos["MATRICULA"].values:
                            media = input("Informe a média do aluno: ")
                            if 0 <= int(media) <= 10:
                                codigos = []
                                matriculas = []
                                medias = []
                                codigos.append(codigo)
                                matriculas.append(matricula)
                                medias.append(media)
                                arquivo_medias = pandas.DataFrame(list(zip(codigos, matriculas, medias)), columns=coluna_medias)
                                arquivo_medias.to_excel("Medias.xlsx", index=False)
                            else:
                                red_message("Média não pode ser menor que 0 ou maior que 10.")
                        else:
                            red_message("Aluno não cadastrado")
                    except FileNotFoundError:
                        red_message("Nenhum aluno cadastrado")
                        break
                else:
                    red_message("Disciplina não cadastrada.")
            except FileNotFoundError:
                red_message("Nenhuma disciplina cadastrada")
                break

        vontade = str.upper(input("Cadastrar outra média?(s/n): "))
        if vontade == "N":
            break
        elif vontade != "S":
            print("Comando não compreendido")


def cadastrar_faltas():
    while True:
        try:
            arquivo_faltas = pandas.read_excel("Faltas.xlsx")
            try:
                arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
                codigo = input("Informe o código da disciplina: ")
                if int(codigo) in arquivo_discipli["CÓDIGO"].values:
                    try:
                        arquivo_alunos = pandas.read_excel("Alunos.xlsx")
                        matricula = str.upper(input("Informe a matricula do aluno: "))
                        if int(matricula) in arquivo_alunos["MATRICULA"].values:
                            falta = input("Informe as faltas do aluno: ")
                            arquivo_faltas.loc[len(arquivo_faltas)] = [codigo, matricula, falta]
                            arquivo_faltas.to_excel("Medias.xlsx", index=False)

                        else:
                            red_message("Aluno não cadastrado.")
                    except FileNotFoundError:
                        red_message("Nenhum aluno cadastrado.")
                        break
                else:
                    red_message("Disciplina não cadastrada.")
            except FileNotFoundError:
                red_message("Nenhuma disciplina cadastrada.")
                break
        except FileNotFoundError:
            try:
                arquivo_discipli = pandas.read_excel("Disciplinas.xlsx")
                codigo = input("Informe o código da disciplina: ")
                if int(codigo) in arquivo_discipli["CÓDIGO"].values:
                    try:
                        arquivo_alunos = pandas.read_excel("Alunos.xlsx")
                        matricula = input("Informe a matricula do aluno: ")
                        if int(matricula) in arquivo_alunos["MATRICULA"].values:
                            falta = input("Informe o número de faltas:")
                            codigos = []
                            matriculas = []
                            faltas_lista = []
                            codigos.append(codigo)
                            matriculas.append(matricula)
                            faltas_lista.append(falta)
                            arquivo_faltas = pandas.DataFrame(list(zip(codigos, matriculas, faltas_lista)), columns=coluna_faltas)
                            arquivo_faltas.to_excel("Faltas.xlsx", index=False)
                        else:
                            red_message("Aluno não cadastrado")
                    except FileNotFoundError:
                        red_message("Nenhum aluno cadastrado")
                        break
                else:
                    red_message("Disciplina não cadastrada.")
            except FileNotFoundError:
                red_message("Nenhuma disciplina cadastrada")
                break

        vontade = str.upper(input("Cadastrar outra disciplina?(s/n): "))
        if vontade == "N":
            break
        elif vontade != "S":
            print("Comando não compreendido")


def criar_relatorio():
    print("Em desenvolvimento")


"""Código principal"""
while True:
    vontade = input("1 - Cadastrar Professores.\n2 - Cadastrar Alunos.\n3 - Cadastrar Disciplinas.\n"
                    "4 - Cadastro de médias.\n5 - Cadastro de Faltas\n6 - Criar Relatório\n7 - Sair: ")

    if vontade == "1":
        cadastro_prof()

    elif vontade == "2":
        cadastro_aluno()

    elif vontade == "3":
        cadastrar_disciplina()

    elif vontade == "4":
        cadastrar_medias()

    elif vontade == "5":
        cadastrar_faltas()

    elif vontade == "6":
        criar_relatorio()

    elif vontade == "7":
        break

    else:
        print("Comando não reconhecido.")
