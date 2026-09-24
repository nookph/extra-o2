# Classificador de UPG (Realizados x Compromisso)

> **Nota:** este repositório é um sandbox isolado para experimentar mudanças
> de design (principalmente na interface gráfica) sem afetar o repositório
> original `extra-o`, que continua sendo a versão estável em uso.

Automação em Python para substituir o fluxo manual/Power Automate de
classificar, todo mês, as linhas de UPG das planilhas do SAP (Realizados e
Compromisso), reaproveitando automaticamente as classificações já feitas em
meses anteriores.

## Como funciona

Só **uma pessoa** precisa ter Python instalado (quem hoje processa as
planilhas do SAP). O resultado de cada etapa é sempre um arquivo **Excel
comum**, que qualquer colega já sabe abrir e editar — ninguém mais precisa
instalar nada.

1. **`preparar`**: lê as duas planilhas brutas exportadas do SAP, filtra
   apenas as linhas em que a coluna OBJETO contém "UPG", e cria um Excel de
   trabalho com duas abas (`Realizados` e `Compromisso`) mais um `Resumo`.
   - Linhas cujo número de PO já está classificado (em qualquer um dos dois
     relatórios, de qualquer mês anterior) vêm com a classificação
     preenchida automaticamente (`Status = OK (herdado)`).
   - Linhas novas vêm em branco e destacadas em amarelo (`Status =
     PENDENTE`), para o time classificar manualmente como já faz hoje.
2. Time classifica manualmente as linhas `PENDENTE`, preenchendo a coluna
   `Classificacao` (e `Observacao`, se quiser) direto no Excel gerado.
3. **`consolidar`**: lê esse Excel já classificado e grava (upsert) as
   classificações novas/alteradas no banco persistente `classificacoes.xlsx`
   (uma linha por PO). No mês seguinte, o passo 1 já aplica tudo de novo
   automaticamente.

Como a chave é o **número do PO**, um PO que já foi classificado em
Realizados também é reconhecido se aparecer em Compromisso (e vice-versa) —
resolvendo o caso de reclassificação (ex: item de outro centro de custo que
não deveria estar ali).

## Instalação

Sem terminal: clique duas vezes em **`Instalar_Dependencias.bat`**. Ele
verifica se o Python está instalado (e explica como instalar caso não
esteja) e instala pandas/openpyxl sozinho — só precisa fazer isso uma vez,
na máquina de quem vai usar o programa.

Se preferir pela linha de comando:

```bash
pip install -r requirements.txt
```

## Interface gráfica (para quem não quer usar terminal)

Tem uma versão com janela e botões, sem precisar abrir CMD/PowerShell:

- Clique duas vezes em **`Abrir_Classificador_UPG.bat`** (ou `Abrir_Classificador_UPG.pyw`) na pasta do projeto.
- Abre uma janela com:
  - **Passo 1**: escolher a pasta com os 2 arquivos do SAP (o programa acha sozinho qual é Realizado/Compromisso), onde salvar a planilha do mês, e onde fica o banco de classificações (já vem com um local padrão em `Documentos/Classificador_UPG`). Botão **"Gerar planilha do mês"**.
  - **Passo 2**: escolher a planilha já classificada. Botão **"Salvar classificações no banco"**.
  - Uma área de mensagens mostra o mesmo progresso do modo linha de comando (quantas linhas de UPG, quantas herdadas, avisos).
  - A planilha gerada abre sozinha ao terminar, e o programa lembra a última pasta/arquivo usados — do segundo mês em diante é só conferir e clicar.

Pré-requisito: a pessoa que vai usar esse atalho precisa ter Python
instalado na máquina e ter rodado o `Instalar_Dependencias.bat` (ou o
`pip install` acima) uma vez — o `.bat`/`.pyw` da interface só abre a
janela, sem terminal visível.

**Este é o caminho recomendado em ambiente corporativo.** Diferente de um
`.exe`, o `.bat`/`.pyw` não é um programa novo e desconhecido — ele só
aciona o Python que já está instalado e liberado na máquina, então costuma
passar sem problemas por políticas de TI que bloqueiam executáveis
desconhecidos (ver aviso abaixo).

### Sobre gerar um .exe (não recomendado se a TI bloqueia executáveis)

É possível empacotar tudo num único `.exe` com o
[PyInstaller](https://pyinstaller.org/) (rodando numa máquina Windows):

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name ClassificadorUPG Abrir_Classificador_UPG.pyw
```

**Atenção:** em empresas com política de bloqueio de executáveis não
aprovados (AppLocker/WDAC), esse `.exe` será bloqueado — e não tem como
contornar isso reempacotando de outro jeito, porque o bloqueio é por
política, não por o `.exe` ser malicioso. Se isso acontecer:

- Use o `.bat`/`.pyw` em vez do `.exe` (não é um executável novo, ver acima).
- Se até o `.bat` for bloqueado (algumas políticas também restringem
  scripts), crie um **atalho do Windows (.lnk)** em vez de rodar o `.bat`
  diretamente: botão direito na área de trabalho → Novo → Atalho → no
  campo do local, aponte para o `pythonw.exe` (ex:
  `C:\Users\voce\AppData\Local\Programs\Python\Python312\pythonw.exe`)
  seguido de `-m sap_upg.gui`, e em "Iniciar em" coloque a pasta do
  projeto. Um atalho `.lnk` só aponta para um programa já aprovado
  (o próprio Python), então normalmente não cai nas mesmas regras de
  bloqueio de scripts.
- Se nada disso funcionar, o `.exe` (ou o próprio Python) precisa ser
  liberado pela TI — é um pedido normal para ferramentas internas
  (whitelisting por hash/caminho), não indica nenhum problema com o
  programa em si.

## Uso mensal (linha de comando)

Os nomes das colunas de PO já vêm configurados como padrão (confirmados
pelo usuário): `Documento de compras` em Realizados e `Nº doc.de
referência` em Compromisso.

A coluna OBJETO é reconhecida **sem diferenciar maiúsculas/minúsculas**
(`OBJETO`, `Objeto`, `objeto`... tanto faz), porque na prática o SAP não
usa sempre a mesma capitalização nas duas planilhas — isso já foi
observado em produção (`OBJETO` numa exportação, `Objeto` na outra) e o
script lida com isso sozinho, sem precisar de ajuste.

Os arquivos que o SAP exporta também não precisam ser apontados pelo nome
exato: o nome sempre muda por causa da data (ex:
`ZMM_Realizado_24092026.xlsx`), mas sempre contém a palavra **Realizado**
ou **Compromisso**. Basta indicar a pasta onde os dois arquivos foram
salvos (`--pasta`) que o script encontra sozinho qual é qual — e avisa se
achar mais de um candidato (ex: sobrou um arquivo do mês passado na
pasta), usando o mais recente e listando o que foi ignorado.

```bash
# 1) Gerar o Excel de trabalho do mês (aponte a pasta onde salvou os 2 arquivos do SAP)
python -m sap_upg preparar \
    --pasta "C:/Users/voce/Downloads" \
    --saida "UPG_Setembro_para_classificar.xlsx"

# 2) Depois de classificar as linhas PENDENTE no Excel gerado, consolidar:
python -m sap_upg consolidar \
    --trabalho "UPG_Setembro_para_classificar.xlsx"
```

Se preferir apontar os arquivos manualmente (ex: nomes atípicos, ou
arquivos em pastas diferentes), ainda é possível usar `--realizados` e
`--compromisso` com o caminho exato — nesse caso a busca automática por
pasta é ignorada.

Por padrão o banco fica em `classificacoes.xlsx` na pasta atual; use
`--banco caminho/para/o/banco.xlsx` para guardá-lo em outro lugar (ex: uma
pasta compartilhada da equipe, para todos acompanharem o histórico).

Se o SAP mudar o nome de alguma coluna no futuro, sobrescreva com
`--po-col-realizados "..."` / `--po-col-compromisso "..."` em qualquer um
dos dois comandos.

### Se a coluna OBJETO tiver outro nome de verdade (não só maiúscula/
minúscula, que já é tratado automaticamente), ou a planilha tiver linhas
de cabeçalho extras (comum em exportações do SAP)

```bash
python -m sap_upg preparar \
    --objeto-col "Descricao do Objeto" \
    --header-realizados 2 \
    --header-compromisso 1 \
    ...
```

`--header-realizados`/`--header-compromisso` indicam em qual linha (contando
do zero) está o cabeçalho de verdade, caso o SAP exporte linhas de título
antes da tabela.

Se um nome de coluna estiver errado, o script mostra a lista de colunas
disponíveis na planilha para facilitar o ajuste.

## Estrutura do projeto

```
sap_upg/
  core.py     # lógica de filtro UPG, normalização de PO, banco e formatação
  cli.py      # comandos "preparar" e "consolidar" (linha de comando)
  gui.py      # interface gráfica (Tkinter) para quem não usa terminal
  appdata.py  # guarda a última pasta/arquivo usados pela interface gráfica
Abrir_Classificador_UPG.bat / .pyw   # atalhos para abrir a interface gráfica com duplo clique
Instalar_Dependencias.bat            # instala pandas/openpyxl com duplo clique (rodar uma vez)
```

## Próximos ajustes possíveis

- Se o mesmo PO puder legitimamente precisar de classificações diferentes
  em linhas diferentes (ex: parcelado por item/centro de custo), a chave de
  match precisa deixar de ser só o PO — hoje o desenho assume PO único.

## Manutenção e customização

Quer trocar nomes de coluna, adaptar para outro tipo de filtro além de
"UPG", entender por que o projeto foi feito de um jeito e não de outro, ou
simplesmente dar continuidade a isso sem ter acompanhado o desenvolvimento
original? Veja **[CONTRIBUTING.md](CONTRIBUTING.md)**.
