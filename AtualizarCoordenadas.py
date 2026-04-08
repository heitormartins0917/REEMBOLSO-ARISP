import os
import json
import time
import pyautogui

# CONFIGURAÇÕES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_CONFIG = os.path.join(BASE_DIR, "config")
ARQUIVO_COORDENADAS = os.path.join(PASTA_CONFIG, "coordenadas.json")

os.makedirs(PASTA_CONFIG, exist_ok=True)

TEMPO_CAPTURA = 4  # segundos para posicionar o mouse

# FUNÇÕES

def capturar_posicao(nome_campo):
    print("\n" + "=" * 50)
    print(f"Prepare-se para capturar: {nome_campo}")
    input("Pressione ENTER para começar a contagem...")

    for i in range(TEMPO_CAPTURA, 0, -1):
        print(f"Capturando em {i}...", end="\r")
        time.sleep(1)

    pos = pyautogui.position()
    print(" " * 50, end="\r")  # limpa linha
    print(f"{nome_campo} capturado em: Point(x={pos.x}, y={pos.y})")
    return {"x": pos.x, "y": pos.y}

# EXECUÇÃO

print("\n======================================")
print(" CALIBRADOR DE COORDENADAS - SHARE ")
print("======================================")

print("\nFuncionamento:")
print("- Para cada campo, você vai apertar ENTER")
print(f"- Depois terá {TEMPO_CAPTURA} segundos para posicionar o mouse")
print("- O script vai capturar sozinho a posição atual")
print("- NÃO precisa clicar, só posicionar o mouse\n")

coordenadas = {}

coordenadas["valor_estimado"] = capturar_posicao("Valor Estimado")
coordenadas["obs"] = capturar_posicao("Obs")
coordenadas["pasta_processo"] = capturar_posicao("Pasta do Processo")
coordenadas["confirmar"] = capturar_posicao("Confirmar")
coordenadas["docto"] = capturar_posicao("Docto.")

with open(ARQUIVO_COORDENADAS, "w", encoding="utf-8") as f:
    json.dump(coordenadas, f, indent=4, ensure_ascii=False)

print("\n======================================")
print(" COORDENADAS SALVAS COM SUCESSO ")
print("======================================")
print(f"Arquivo salvo em:\n{ARQUIVO_COORDENADAS}")

input("\nPressione ENTER para sair...")