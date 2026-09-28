# Guia de manutenção e customização

Este documento existe para que **qualquer pessoa consiga dar continuidade a
este projeto**, mesmo sem ter acompanhado o desenvolvimento original — a
ideia é a mesma de um projeto open source: código pequeno, sem mágica
escondida, documentado o suficiente para outra pessoa pegar e seguir.

Não é preciso ser programador experiente para mexer aqui. O projeto é
pequeno de propósito (menos de 700 linhas ao todo) e cada arquivo tem uma
responsabilidade única. Se você sabe editar uma planilha e consegue seguir
um passo a passo, consegue fazer os ajustes mais comuns descritos abaixo.

## Visão geral do projeto

O que o programa faz (resumo — detalhes no `README.md`): lê duas planilhas
do SAP (Realizados e Compromisso), filtra as linhas de UPG, aplica
automaticamente classificações já feitas em meses anteriores (guardadas
num "banco" em Excel) e deixa só as linhas novas para classificação manual.

```
sap_upg/
  core.py     # TODA a lógica de negócio (ler planilha, filtrar, classificar, salvar)
  cli.py      # comandos de linha de comando ("preparar"/"consolidar") — só monta os
              # argumentos e chama as funções de core.py
  gui.py      # interface gráfica (Tkinter) — também só monta a tela e chama core.py
  appdata.py  # onde a interface gráfica guarda a última pasta/arquivo usados

Abrir_Classificador_UPG.bat / .pyw   # atalhos de duplo clique para abrir a GUI
Instalar_Dependencias.bat            # instala Python (se faltar) e as dependências, com duplo clique
requirements.txt                     # dependências Python (pandas, openpyxl, ttkbootstrap)
```

Regra importante para quem for mexer: **toda a lógica de negócio mora em
`core.py`**. `cli.py` e `gui.py` são só "capas" — eles não decidem nada
sozinhos, só coletam o que o usuário informou (por linha de comando ou por
clique) e chamam as funções de `core.py` (principalmente
`executar_preparar` e `executar_consolidar`). Se você adicionar uma
funcionalidade nova, o normal é: a lógica entra em `core.py`, e só depois
você expõe ela como um argumento de linha de comando e/ou um campo na
janela.

## Pontos de customização mais comuns

### 1. O SAP mudou o nome de uma coluna

Onde: `sap_upg/core.py`, logo no topo:

```python
DEFAULT_PO_COL_REALIZADOS = "Documento de compras"
DEFAULT_PO_COL_COMPROMISSO = "Nº doc.de referência"
```

Basta editar essas duas strings para o nome novo. Não precisa mexer em
mais nada — `cli.py` e `gui.py` usam essas constantes como valor padrão.

A coluna OBJETO **não** precisa desse cuidado: o casamento de nome já
ignora maiúsculas/minúsculas (`_resolver_nome_coluna` em `core.py`), então
pequenas variações não quebram nada. Só edite `coluna_objeto: str =
"OBJETO"` (na classe `ReportSpec`) se o nome da coluna mudar de verdade
(ex: virar "Descrição do objeto").

### 2. Trocar a palavra-chave de filtro (hoje é "UPG")

Onde: função `filtrar_upg` em `core.py`:

```python
def filtrar_upg(df: pd.DataFrame, coluna_objeto: str) -> pd.DataFrame:
    mascara = df[coluna_objeto].astype(str).str.contains("UPG", case=False, na=False)
    return df.loc[mascara].copy()
```

Se um dia esse mesmo processo precisar ser reaproveitado para outro tipo
de projeto/centro de custo (não só UPG), o ponto certo para generalizar é
esse `"UPG"` — dá para transformar num parâmetro (`palavra_chave: str =
"UPG"`) e propagar até `executar_preparar`, `cli.py` e `gui.py`, do mesmo
jeito que já foi feito para as colunas de PO.

### 3. Onde o banco de classificações fica salvo por padrão

Onde: `sap_upg/appdata.py`:

```python
PASTA_APP = Path.home() / "Documents" / "Classificador_UPG"
```

Isso só afeta o valor sugerido na interface gráfica (o usuário sempre pode
trocar pelo botão "Procurar..."). Se a equipe quiser um local padrão
compartilhado (ex: uma pasta de rede), é só trocar esse caminho.

### 4. Rótulos e cores de status na planilha gerada

Onde: topo de `core.py`:

```python
STATUS_HERDADO = "OK (herdado)"
STATUS_PENDENTE = "PENDENTE"
PENDENTE_FILL = PatternFill(start_color="FFF2CC", ...)  # cor de destaque das linhas pendentes
```

### 5. Atualizar a versão do Python instalada automaticamente

Onde: topo de `Instalar_Dependencias.bat`:

```bat
set "PY_VERSION=3.12.7"
```

Esse é o Python baixado automaticamente de python.org quando a máquina não
tem Python instalado. Troque só esse número para atualizar (confira antes
em https://www.python.org/downloads/windows/ se existe instalador
`amd64.exe` para a versão escolhida).

### 6. Adicionar um terceiro relatório do SAP (além de Realizados/Compromisso)

Não é um caso trivial, mas o caminho é: em `executar_preparar` (`core.py`),
a lista `especificacoes` monta um item por relatório — adicionar um
terceiro item nessa lista, mais uma constante `ABA_<NOME>`, cobre a maior
parte. Depois é só expor os novos parâmetros em `cli.py`
(`construir_parser`) e, se fizer sentido, um campo novo na janela
(`gui.py`).

## Por que as coisas foram feitas assim (para quem for mexer)

- **Banco em Excel, não em banco de dados de verdade (SQLite etc)**: a
  prioridade foi qualquer pessoa da equipe conseguir abrir, conferir e, se
  precisar, editar o histórico de classificações sem precisar de Python.
  Se o volume crescer muito (dezenas de milhares de POs) e o Excel ficar
  lento, essa é a primeira peça a trocar — a interface de
  `carregar_banco`/`salvar_banco` em `core.py` foi pensada para poder ser
  substituída por outro armazenamento sem mudar o resto do código.
- **Chave de reaproveitamento é só o número do PO**: hoje o desenho assume
  que um PO nunca precisa de duas classificações diferentes ao mesmo tempo.
  Se isso deixar de ser verdade (ex: parcelado por item/centro de custo),
  a chave em `aplicar_classificacoes`/`consolidar_no_banco` precisa virar
  composta (PO + outra coluna), e isso teria efeito em várias funções.
- **A GUI nunca chama métodos da janela de dentro da thread de trabalho**:
  o Tkinter não é seguro para isso (já causou um bug real durante o
  desenvolvimento — "main thread is not in main loop"). Toda comunicação
  da thread de trabalho para a tela passa pela fila `self._fila_ui` em
  `gui.py`. Se for adicionar uma ação nova que rode em background, siga
  esse mesmo padrão (`self._enfileirar(...)`), nunca chame `self.after`
  nem widgets diretamente de dentro de uma `threading.Thread`.

## Como testar depois de mexer em algo

Hoje não existe suíte de testes automatizados (o projeto é pequeno o
suficiente para não ter sido necessário até agora). O jeito de validar uma
mudança é gerar planilhas de exemplo e rodar o fluxo completo:

```bash
pip install -r requirements.txt
python -m sap_upg preparar --pasta caminho/com/2/arquivos/de/teste --saida teste.xlsx --banco teste_banco.xlsx
# abra teste.xlsx, preencha a coluna Classificacao nas linhas PENDENTE, salve
python -m sap_upg consolidar --trabalho teste.xlsx --banco teste_banco.xlsx
# rode "preparar" de novo com os mesmos arquivos: as linhas que você classificou
# devem vir como "OK (herdado)" dessa vez
```

Se o projeto crescer bastante, vale a pena introduzir testes automatizados
com `pytest` cobrindo pelo menos `filtrar_upg`, `aplicar_classificacoes` e
`consolidar_no_banco` — são funções puras (recebem dado, devolvem dado),
fáceis de testar sem precisar de arquivos reais do SAP.

## Como levar uma mudança para quem usa o programa

Não existe instalador único — o "programa" é a própria pasta do projeto.
Depois de mexer no código:

1. Teste localmente (seção acima).
2. Se o projeto estiver num repositório Git (recomendado), commit e push.
3. Substitua a pasta na máquina de quem usa pela versão atualizada (ou,
   se lá também tiver Git configurado, um `git pull` resolve).
4. Só é preciso rodar `Instalar_Dependencias.bat` de novo se
   `requirements.txt` mudou.

## Licença de uso

Este projeto foi desenvolvido para uso interno. Não há uma licença open
source formal anexada a ele — se quem for dono do projeto daqui para
frente quiser publicá-lo abertamente (ex: no GitHub, para outras empresas
reaproveitarem), essa é uma decisão de quem estiver com essa
responsabilidade no momento, e vale confirmar antes com quem cuida desse
tipo de assunto na empresa. Até lá, o espírito deste documento é apenas
garantir que **dentro da empresa**, ninguém fique travado por falta de
contexto para dar manutenção.
