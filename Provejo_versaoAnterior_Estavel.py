import os
import time
import pandas as pd
from datetime import datetime, timedelta
from openpyxl.styles import Alignment

# Função Limpar tela
def limparTela():
    sistema_operacional = os.name
    if sistema_operacional == "windows" or sistema_operacional == "Windows":
        os.system("cls")
    else:
        os.system("cls")

# Função Encerrar Programa
def encerrarPrograma():
    print("Encerrando Programa", end="")
    for i in range(3):
        time.sleep(0.5)
        print(".", end="")
    time.sleep(0.3)
    print("")

# Função Exibição de menu
def menuEpi():
    print("\nEscolha um Opção")
    print("1 - Novo Registro")
    print("2 - Sair")

# Função cadastro de data
def dataEntrega():
    while True:
        data_string = input("Insira a data de Entrega (DD/MM/AAAA): ")
        try:
            # Tentar converter a data usando o formato DD/MM/YYYY
            data = datetime.strptime(data_string, "%d/%m/%Y")
        except ValueError:
            try:
                # Se falhar, tentar converter usando o formato DD/MM/YY
                data = datetime.strptime(data_string, "%d/%m/%y")
            except ValueError:
                print("Formato de data inválido. Por favor, insira no formato DD/MM/AAAA ou DD/MM/AA.")
                continue

        data_atual = datetime.now()

        # Verificar se a data inserida é menor ou igual à data atual
        data_vencimento = data + timedelta(days=365) # cálculo de vencimento da epi
        if data <= data_atual:
            return data.strftime("%d/%m/%Y"), data_vencimento.strftime("%d/%m/%Y")
        else:
            print("Você inseriu uma data futura. Por favor, insira uma data válida.")

# Função com laço de repetição na Epi
def laco_Epi():
    lista = []
    epiLista = ["Calça", "Camisa", "Bota", "Boné", "Crachá", "Cracha"]
    while True:
        try:
            # Nome
            nomeEpi = str(input("Digite o nome da Epi: "))
            nEpi = "".join(word.capitalize() for word in nomeEpi.split())
            while nEpi not in epiLista:
                print("Essa Epi não está registrada")
                nomeEpi = str(input("Digite o nome da Epi: "))
                nEpi = "".join(word.capitalize() for word in nomeEpi.split())
            # Tamanho
            tamanhoEpi = str(input("Digite o tamanho da Epi: ")).upper()  # case maiúscula para todas as letras
            while len(tamanhoEpi) == 0:
                print("Digite um Tamanho para a Epi")
                tamanhoEpi = str(input("Digite o tamanho da Epi: ")).upper()

            # Quantidade
            while True:
                try:
                    qtd_ = int(input("Digite a quantidade: "))
                    break
                except ValueError:
                    print("Digite uma quantidade -> (Em Números)")

            # Data
            data_entrega, data_vencimento = dataEntrega()
            global separador
            lista.append((separador, nEpi, tamanhoEpi, qtd_, data_entrega, data_vencimento))

            print("\nFoi enviado mais alguma Epi ?")
            escolha_ = input("Sim ou Não ? ").lower()
            if escolha_ == "sim" or escolha_ == "sim ":
                for i in range(1):
                    separador += 1
                print("")
                continue
            elif escolha_ == "não" or escolha_ == "nao" or escolha_ == "não ":
                break
            else:
                print("Opção Não Identificada.")
                print("Continuar ? (sim ou não) ", end="")
                ence_conti = input("").lower()
                if ence_conti == "sim":
                    separador += 1
                    print("")
                    continue
                else:
                    break
        except ValueError:
            print("Dado inserido inválido")
            time.sleep(2)
            limparTela()
    separador = 1
    return lista

# Função de cadastro Inicial
def cadastroEpi():
    while True:
        try:
            # Início Menu
            menuEpi()
            escolhido = input("\nDigite a sua escolha: ").lower()
            # Lista com o nome dos postos presentes na rede
            postos = ["Xpres", "Valente", "Rei Davi", "Querubim", "Prj", "Riu Una", "Rosa Flor", "Elefantinho", "Pel",
                      "Pv", "Rd", "Ru", "Rf"]
            # Condicional Menu
            if escolhido == "1" or escolhido == "novo registro":
                nomePosto = str(input("Digite o nome do posto: "))
                nPosto = " ".join(word.capitalize() for word in nomePosto.split())  # case maiúscula para a primeira letra de cada palavra
                # Enquanto o nome do posto digitado não estiver dentro da lista
                while nPosto not in postos:
                    print("Nome Inválido")
                    nomePosto = str(input("Digite o nome do posto: ")).capitalize()
                    nPosto = " ".join(word.capitalize() for word in nomePosto.split())

                nomeFuncio = str(input("Digite o nome do Funcionário: ")).capitalize()
                # Enquanto o nome Funcionário for menor ou igual a 1 caractere
                while len(nomeFuncio) <= 1:
                    print("Digite o seu nome, por favor!")
                    nomeFuncio = str(input("Digite o nome do Funcionário: ")).capitalize()

                lacoEpi = laco_Epi()

                # Printar na tela e salvar as informações
                limparTela()
                print("+----------------------------------+")
                print("|        Informações Cadastradas   |")
                print("+----------------------------------+")
                print(f"\nNome do Posto: {nPosto}")
                print(f"Nome do Funcionário: {nomeFuncio}")

                for item in lacoEpi:
                    separador, nomeEpi, tamanhoEpi, qtd_, data_entrega, data_vencimento = item
                    print(f"\n{separador} Epi")
                    print(f"Nome da Epi: {nomeEpi}")
                    print(f"Tamanho da Epi: {tamanhoEpi}")
                    print(f"Quantidade: {qtd_}")
                    print(f"Data de Entrega: {data_entrega}")
                    print(f"Data de Vencimento: {data_vencimento}")

                print("-" * 20)  # Adiciona uma linha divisória entre os itens

                # Salvar Informações
                salvar = input("\nSalvar Informações? (Sim ou Não): ").lower()

                if salvar == "sim":
                    print("\nRegistro(s) Salvo(s)")
                    time.sleep(1)
                else:
                    print("Registro(s) não Salvo(s)!")
                    time.sleep(2)

            # Opção Finalizar Programa
            elif escolhido == "2" or escolhido == "sair":
                encerrarPrograma()
                break
            # Caso nenhuma opção seja reconhecida
            else:
                print("\nOps! Opção não encontrada, Josa.")

            # Função limpar a tela do programa
            limparTela()

        except ValueError:
            print("\nOps! Digite a quantidade em Números.")
            time.sleep(1)
            limparTela()
        except Exception:
            print("Ops! Algo de errado não está certo, Josa!")
            time.sleep(1)
            limparTela()

# Escopo Principal
separador = 1
cadastroEpi()
