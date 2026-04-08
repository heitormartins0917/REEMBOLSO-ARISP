# REEMBOLSO ARISP

Automação em Python para apoiar o fluxo de **extração de dados de PDFs** e **lançamento automático no sistema SHARE**, com controle por planilha Excel.

---

# OBJETIVO DO PROJETO

Este projeto foi criado para automatizar um processo operacional repetitivo de reembolsos, reduzindo:

- tempo manual
- chance de erro humano
- retrabalho
- cansaço operacional

O fluxo atual do projeto está dividido em **duas grandes etapas**:

1. **Extrair dados dos PDFs**
2. **Lançar automaticamente no SHARE**

Além disso, o projeto possui um sistema de **calibração de coordenadas**, para adaptar a automação a diferentes computadores, resoluções e posições de tela.

--

# DESCRIÇÃO DE CADA ARQUIVO

## `ExtrairValoresMatricula.py`
Script responsável por:

- ler os PDFs da pasta `entrada_pdf`
- extrair:
  - tipo (Matrícula ou Prévia)
  - matrícula / número da prévia
  - GCPJ
  - pasta
  - valor total
- gerar uma planilha Excel com os dados
- marcar se a extração foi bem-sucedida

---

## `LancarSHAREMatricula.py`
Script responsável por:

- abrir a planilha mais recente da pasta `saida_excel`
- ler os dados extraídos
- ignorar linhas já lançadas
- ignorar linhas com erro de extração
- preencher automaticamente o sistema SHARE
- marcar no Excel quais linhas já foram lançadas
- pintar as linhas lançadas de verde

---

## `AtualizarCoordenadas.py`
Script responsável por:

- capturar as coordenadas da tela do sistema SHARE
- salvar essas posições em um arquivo JSON
- permitir recalibrar a automação em outro computador ou tela

---

## `requirements.txt`
Lista das bibliotecas Python necessárias para o projeto.

---

## `README.md`
Este arquivo de documentação.

---

# DESCRIÇÃO DAS PASTAS

## `entrada_pdf/`
Pasta onde devem ser colocados os **PDFs que serão processados**.

### Importante:
Os arquivos podem vir com nomes variados, mas o sistema espera que o nome contenha, nessa ordem:

```text
CODIGO-GCPJ-PASTA.pdf
```

### Exemplos:
```text
R100016-2600046952-482257.pdf
86055-2100974897-188501.pdf
PO12345988C-2100974897-188501.pdf
```

---

## `saida_excel/`
Pasta onde serão salvas as planilhas geradas automaticamente pelo extrator.

Exemplo:
```text
Reembolso_ARISP_2026-04-07_14-32-11.xlsx
```

---

## `logs/`
Pasta reservada para logs futuros do projeto.

Atualmente ainda não utilizada de forma avançada, mas mantida para evolução futura.

---

## `config/`
Pasta de configurações do projeto.

Atualmente armazena:

### `coordenadas.json`
Arquivo com as coordenadas da tela utilizadas pelo lançador automático.

---

# COMO INSTALAR O PROJETO NO VS CODE

---

## 1) INSTALAR O PYTHON

Baixe e instale o Python no site oficial:

- https://www.python.org/downloads/

### Durante a instalação:
marque a opção:

```text
Add Python to PATH
```

Se o Python já estiver instalado, pode seguir.

---

## 2) ABRIR O PROJETO NO VS CODE

1. Abra o VS Code
2. Clique em **Arquivo > Abrir Pasta**
3. Selecione a pasta:

```text
REEMBOLSO
```

---

## 3) INSTALAR AS BIBLIOTECAS

Abra o terminal do VS Code e rode:

```powershell
python -m pip install -r requirements.txt
```

Se o comando `python` não funcionar, tente:

```powershell
py -m pip install -r requirements.txt
```

---

# CONTEÚDO DO `requirements.txt`

Crie o arquivo `requirements.txt` com este conteúdo:

```txt
pandas
openpyxl
pymupdf
pyautogui
pyperclip
```

---

# FLUXO DE USO DO PROJETO

O projeto funciona em **3 etapas principais**.

---

# ETAPA 1 — CALIBRAR AS COORDENADAS (SE NECESSÁRIO)

> Esta etapa só precisa ser feita se:
> - for outro computador
> - mudar resolução
> - mudar escala do Windows
> - mudar posição da janela do navegador
> - mudar layout do SHARE

Rode:

```powershell
python AtualizarCoordenadas.py
```

O script vai pedir que você capture os campos:

- Valor Estimado
- Obs
- Pasta do Processo
- Confirmar
- Docto.

Ao final, será criado/atualizado o arquivo:

```text
config/coordenadas.json
```

---

# ETAPA 2 — EXTRAIR DADOS DOS PDFs

Coloque os PDFs dentro da pasta:

```text
entrada_pdf
```

Depois rode:

```powershell
python ExtrairValoresMatricula.py
```

O script irá:

- ler os PDFs
- extrair os dados
- gerar uma planilha Excel automaticamente

Essa planilha será salva em:

```text
saida_excel
```

---

# ETAPA 3 — LANÇAR NO SHARE

Abra o sistema SHARE na tela correta de lançamento.

Depois rode:

```powershell
python LancarSHAREMatricula.py
```

O script irá:

- localizar a planilha mais recente
- ler apenas linhas válidas
- ignorar linhas já lançadas
- preencher automaticamente os campos do sistema
- clicar em confirmar
- atualizar a planilha

---

# ESTRUTURA DA PLANILHA GERADA

A planilha gerada possui as seguintes colunas:

| Coluna | Descrição |
|--------|-----------|
| Tipo | Identifica se é Matrícula ou Prévia |
| Matricula | Código principal do processo |
| GCPJ | Número GCPJ |
| Pasta | Número da pasta |
| Valor Total | Valor extraído do PDF |
| Status Extração | Resultado da leitura do PDF |
| Status SHARE | Resultado do lançamento no sistema |

---

# SIGNIFICADO DOS STATUS

## `Status Extração`

Indica se o PDF foi lido corretamente.

### Valores possíveis:
- `OK`
- `ERRO`

---

## `Status SHARE`

Indica se a linha já foi lançada no sistema SHARE.

### Valores possíveis:
- vazio = ainda não lançado
- `OK` = lançado com sucesso
- `ERRO` = tentativa com problema (caso seja implementado depois)

---

# COMO O LANÇADOR DECIDE O QUE FAZER

O `LancarSHAREMatricula.py` segue estas regras:

## Ele PROCESSA apenas linhas com:
- `Status Extração = OK`
- `Status SHARE` vazio

## Ele IGNORA linhas com:
- `Status Extração = ERRO`
- `Status SHARE = OK`

Isso permite:

- continuar de onde parou
- evitar lançamentos duplicados
- usar a mesma planilha em dias diferentes

---

# SEGURANÇA DA AUTOMAÇÃO

A automação usa o fail-safe do `pyautogui`.

## Como parar imediatamente:
jogue o mouse para um canto da tela.

Isso interrompe o robô com segurança.

---

# IMPORTANTE SOBRE A TELA DO SHARE

Para o lançador funcionar corretamente, a tela precisa estar:

- no mesmo layout calibrado
- com os campos visíveis
- pronta para digitação
- sem interferência manual durante a execução

### Durante a execução:
- não mexa no mouse
- não mexa no teclado
- não troque de janela

---

# OBSERVAÇÕES IMPORTANTES

## 1) O primeiro lançamento é diferente
O sistema SHARE pode apresentar uma mensagem ou comportamento visual diferente **somente após o primeiro lançamento**.

Por isso, o script já foi ajustado para tratar essa diferença usando o campo **Docto.** como referência de navegação entre os próximos lançamentos.

---

## 2) O campo Pasta pode vir como decimal
Em alguns casos, o Excel pode interpretar a pasta como número decimal (exemplo: `1114240.0`).

O script já possui tratamento para converter isso automaticamente em:

```text
1114240
```

---

## 3) Coordenadas podem variar
Mesmo pequenas mudanças na tela podem afetar a automação.

Se algo estiver clicando no lugar errado, a primeira suspeita deve ser:

- coordenadas desatualizadas

Nesse caso, rode novamente:

```powershell
python AtualizarCoordenadas.py
```

---

# EXEMPLO DE USO IDEAL

Fluxo recomendado de operação:

1. Colocar PDFs em `entrada_pdf`
2. Rodar `ExtrairValoresMatricula.py`
3. Conferir rapidamente a planilha gerada
4. Abrir o SHARE na tela correta
5. Rodar `LancarSHAREMatricula.py`
6. Acompanhar a execução sem interferir

---

# FUTURAS MELHORIAS POSSÍVEIS

Este projeto foi pensado para crescer por etapas.

Melhorias futuras possíveis:

- validação automática da planilha antes do lote
- logs detalhados de execução
- tratamento de erros do SHARE
- diferenciação de fluxos entre Matrícula e Prévia
- interface gráfica simples
- separação por módulos
- suporte a múltiplos processos operacionais

---

# RESUMO FINAL

Este projeto automatiza dois trabalhos principais:

## 1. Extração
Transforma PDFs em uma planilha organizada.

## 2. Lançamento
Usa essa planilha para lançar os dados automaticamente no sistema SHARE.

Com isso, o processo fica:

- mais rápido
- mais consistente
- mais seguro
- mais escalável

---