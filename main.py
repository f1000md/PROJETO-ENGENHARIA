import pyautogui
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import time
import os

os.system("cls")

def realizar_acao(nome_acao, categoria, parametro):
    try:
        inicio = time.time()

        tipo = categoria.lower()

        if tipo == "texto":
            pyautogui.write(parametro, interval=0.1)
        elif tipo == "tecla":
            pyautogui.press(parametro)
        elif tipo == "espera":
            time.sleep(float(parametro))
        elif tipo == "hotkey":
            teclas = parametro.split('+')
            pyautogui.hotkey(*teclas)
        elif tipo == "click":
            if parametro.startswith('(') and parametro.endswith(')'):
                coords = parametro.strip("()")
                dados = coords.split(',')
                pos = {}
                for item in dados:
                    chave, valor = item.split('=')
                    pos[chave.strip().lower()] = int(valor.strip())
                if 'x' in pos and 'y' in pos:
                    pyautogui.click(pos['x'], pos['y'])
                else:
                    print("Erro: coordenadas incompletas.")
            else:
                return (nome_acao, "Coordenadas mal formatadas", 0)
        else:
            return (nome_acao, "Categoria de ação inválida", 0)

        duracao = round(time.time() - inicio, 2)
        return (nome_acao, "Executado com sucesso", duracao)

    except Exception as erro:
        return (nome_acao, f"Falha: {str(erro)}", 0)

def salvar_relatorio(lista_registros):
    wb = Workbook()
    aba = wb.active
    aba.append(["Ação", "Resultado", "Tempo (s)"])

    for item in lista_registros:
        aba.append(item)

    pasta_destino = "Relatorios"
    os.makedirs(pasta_destino, exist_ok=True)

    nome_arquivo = f"Relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    caminho_completo = os.path.join(pasta_destino, nome_arquivo)
    wb.save(caminho_completo)

    return caminho_completo

def iniciar_execucao():
    try:
        tarefas_df = pd.read_csv("tarefas.csv")
    except FileNotFoundError:
        print("Arquivo 'tarefas.csv' não localizado!")
        return

    print("Preparando execução automática...")
    for contagem in range(3, 0, -1):
        print(f"Início em {contagem} segundo(s)...")
        time.sleep(1)
        os.system("cls")

    acoes_executadas = []
    erros = 0
    desconhecidas = 0

    for _, linha in tarefas_df.iterrows():
        resultado = realizar_acao(linha["Tarefa"], linha["Tipo"], linha["Dado"])
        acoes_executadas.append(resultado)
        print(f"• [{resultado[1]}] - {linha['Tarefa']} ({resultado[2]}s)")

        if resultado[1] == "Executado com sucesso":
            continue
        elif resultado[1] == "Categoria de ação inválida":
            desconhecidas += 1
        else:
            erros += 1

    caminho_arquivo = salvar_relatorio(acoes_executadas)

    print("\nExecução concluída!")
    print(f"Relatório salvo em: {caminho_arquivo}")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    if desconhecidas:
        print(f"Ações desconhecidas: {desconhecidas}")
    if erros:
        print(f"Falhas detectadas: {erros}")

if __name__ == '__main__':
    iniciar_execucao()
