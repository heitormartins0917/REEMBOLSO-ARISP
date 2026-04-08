import os
import re
import fitz  # pymupdf
import pandas as pd
from datetime import datetime

# =========================
# CONFIGURAÇÕES DO PROJETO
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_ENTRADA = os.path.join(BASE_DIR, "entrada_pdf")
PASTA_SAIDA = os.path.join(BASE_DIR, "saida_excel")

os.makedirs(PASTA_ENTRADA, exist_ok=True)
os.makedirs(PASTA_SAIDA, exist_ok=True)

# =========================
# FUNÇÕES AUXILIARES
# =========================

def limpar_texto(texto):
    if not texto:
        return ""
    return re.sub(r"\s+", " ", texto).strip()

def normalizar_valor(valor_str):
    if not valor_str:
        return ""
    return valor_str.strip().replace(" ", "")

def extrair_texto_pdf(caminho_pdf):
    try:
        doc = fitz.open(caminho_pdf)
        texto_completo = []

        for pagina in doc:
            texto_completo.append(pagina.get_text())

        doc.close()
        return limpar_texto(" ".join(texto_completo))

    except Exception as e:
        return f"ERRO_LEITURA_PDF: {e}"

def extrair_valor_total_pdf(texto_pdf):
    if not texto_pdf or texto_pdf.startswith("ERRO_LEITURA_PDF"):
        return ""

    padroes_prioritarios = [
        r"Valor\s*Total[:\s]*R\$\s*([\d\.,]+)",
        r"VALOR\s*TOTAL[:\s]*R\$\s*([\d\.,]+)",
        r"Total[:\s]*R\$\s*([\d\.,]+)",
        r"Valor\s*a\s*Pagar[:\s]*R\$\s*([\d\.,]+)",
        r"Valor\s*do\s*Documento[:\s]*R\$\s*([\d\.,]+)",
    ]

    for padrao in padroes_prioritarios:
        match = re.search(padrao, texto_pdf, re.IGNORECASE)
        if match:
            return normalizar_valor(match.group(1))

    valores_encontrados = re.findall(r"R\$\s*([\d\.,]+)", texto_pdf, re.IGNORECASE)

    if len(valores_encontrados) >= 1:
        return normalizar_valor(valores_encontrados[-1])

    return ""

def identificar_tipo_por_codigo(codigo):
    codigo = str(codigo).strip().upper()

    if re.fullmatch(r"PO\d+[A-Z]?", codigo):
        return "Prévia"

    return "Matrícula"

def identificar_tipo_final(codigo, valor_total):
    tipo_codigo = identificar_tipo_por_codigo(codigo)

    if tipo_codigo == "Prévia":
        return "Prévia"

    if valor_total == "7,62":
        return "Prévia"

    return "Matrícula"

def extrair_dados_nome_arquivo(nome_arquivo):
    nome_sem_ext = os.path.splitext(nome_arquivo)[0]
    partes = nome_sem_ext.split("-")

    codigo = partes[0].strip() if len(partes) > 0 else ""
    gcpj = partes[1].strip() if len(partes) > 1 else ""
    pasta = partes[2].strip() if len(partes) > 2 else ""

    return codigo, gcpj, pasta

def validar_extracao(codigo, gcpj, pasta, valor_total, texto_pdf):
    if not codigo or not gcpj or not pasta:
        return "ERRO"

    if not texto_pdf or texto_pdf.startswith("ERRO_LEITURA_PDF"):
        return "ERRO"

    if not valor_total:
        return "ERRO"

    return "OK"

# =========================
# PROCESSAMENTO PRINCIPAL
# =========================

print("\n==============================")
print(" EXTRATOR DE REEMBOLSOS ARISP ")
print("==============================\n")

arquivos_pdf = sorted([
    f for f in os.listdir(PASTA_ENTRADA)
    if f.lower().endswith(".pdf")
])

if not arquivos_pdf:
    print("Nenhum PDF encontrado na pasta 'entrada_pdf'.")
    input("\nPressione Enter para sair...")
    raise SystemExit

print(f"{len(arquivos_pdf)} PDF(s) encontrado(s).\n")

dados = []

for i, arquivo in enumerate(arquivos_pdf, start=1):
    caminho_pdf = os.path.join(PASTA_ENTRADA, arquivo)

    print(f"[{i}/{len(arquivos_pdf)}] Processando: {arquivo}")

    codigo, gcpj, pasta = extrair_dados_nome_arquivo(arquivo)
    texto_pdf = extrair_texto_pdf(caminho_pdf)
    valor_total = extrair_valor_total_pdf(texto_pdf)
    tipo = identificar_tipo_final(codigo, valor_total)
    status_extracao = validar_extracao(codigo, gcpj, pasta, valor_total, texto_pdf)

    # Regra extra de segurança
    tipo_codigo = identificar_tipo_por_codigo(codigo)
    if tipo_codigo == "Matrícula" and valor_total == "7,62":
        status_extracao = "ERRO"

    dados.append({
        "Tipo": tipo,
        "Matricula": codigo,
        "GCPJ": gcpj,
        "Pasta": pasta,
        "Valor Total": valor_total,
        "Status Extração": status_extracao,
        "Status SHARE": ""
    })

df = pd.DataFrame(dados)

colunas_finais = [
    "Tipo",
    "Matricula",
    "GCPJ",
    "Pasta",
    "Valor Total",
    "Status Extração",
    "Status SHARE"
]

df = df[colunas_finais]

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
nome_excel = f"Reembolso_ARISP_{timestamp}.xlsx"
caminho_excel = os.path.join(PASTA_SAIDA, nome_excel)

df.to_excel(caminho_excel, index=False)

total_ok = (df["Status Extração"] == "OK").sum()
total_erro = (df["Status Extração"] == "ERRO").sum()
total_matriculas = (df["Tipo"] == "Matrícula").sum()
total_previas = (df["Tipo"] == "Prévia").sum()

print("\n==============================")
print(" PROCESSAMENTO CONCLUÍDO ")
print("==============================")
print(f"Total de PDFs: {len(df)}")
print(f"Matrículas: {total_matriculas}")
print(f"Prévias: {total_previas}")
print(f"Extração OK: {total_ok}")
print(f"Extração ERRO: {total_erro}")

print("\nPlanilha criada em:")
print(caminho_excel)

input("\nPressione Enter para sair...")