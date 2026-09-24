"""Interface de linha de comando para o classificador de UPG.

Os nomes das colunas de PO já vêm com padrão configurado (Realizados usa
"Documento de compras", Compromisso usa "Nº doc.de referência"), então no
dia a dia não é preciso informá-los — só sobrescreva com --po-col-... se o
SAP mudar o nome da coluna.

Os arquivos exportados pelo SAP também não precisam ser apontados pelo nome
exato: como o nome sempre muda por causa da data, mas sempre contém
"Realizado" ou "Compromisso", basta apontar a pasta onde eles foram salvos
(--pasta) que o script encontra sozinho.

Uso típico todo mês (arquivos salvos na pasta "Downloads", por exemplo):

    # 1) Gera o Excel de trabalho do mês, já com classificações antigas aplicadas
    python -m sap_upg preparar \\
        --pasta "Downloads" \\
        --saida "UPG_Setembro_para_classificar.xlsx"

    # 2) Depois que o time preencher a coluna "Classificacao" nas linhas
    #    marcadas como PENDENTE, roda para salvar no banco definitivo:
    python -m sap_upg consolidar \\
        --trabalho "UPG_Setembro_para_classificar.xlsx"

O banco (classificacoes.xlsx) é reaproveitado automaticamente em todos os
meses seguintes: quando um PO já classificado aparecer de novo (em
Realizados ou em Compromisso), o script preenche a classificação sozinho.

Quem preferir não usar a linha de comando pode abrir a interface gráfica:

    python -m sap_upg.gui
"""

from __future__ import annotations

import argparse

from . import core


def _comando_preparar(args: argparse.Namespace) -> None:
    caminho_saida = core.executar_preparar(
        pasta=args.pasta,
        realizados=args.realizados,
        compromisso=args.compromisso,
        po_col_realizados=args.po_col_realizados,
        po_col_compromisso=args.po_col_compromisso,
        objeto_col=args.objeto_col,
        aba_realizados=args.aba_realizados,
        aba_compromisso=args.aba_compromisso,
        header_realizados=args.header_realizados,
        header_compromisso=args.header_compromisso,
        banco=args.banco,
        saida=args.saida,
    )
    print()
    print("Preencha a coluna 'Classificacao' (e 'Observacao', se quiser) nas linhas destacadas em amarelo (PENDENTE).")
    print("Depois rode o comando 'consolidar' para salvar essas classificações no banco.")
    print(f"Arquivo: {caminho_saida.resolve()}")


def _comando_consolidar(args: argparse.Namespace) -> None:
    core.executar_consolidar(
        trabalho=args.trabalho,
        po_col_realizados=args.po_col_realizados,
        po_col_compromisso=args.po_col_compromisso,
        banco=args.banco,
    )


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sap_upg",
        description="Automação de classificação de UPG a partir dos relatórios SAP (Realizados/Compromisso).",
    )
    subparsers = parser.add_subparsers(dest="comando", required=True)

    p_preparar = subparsers.add_parser("preparar", help="Gera o Excel mensal de trabalho já com classificações antigas aplicadas.")
    p_preparar.add_argument(
        "--realizados",
        default=None,
        help="Caminho do arquivo de Realizados do SAP. Se omitido, procura na --pasta um arquivo com 'Realizado' no nome.",
    )
    p_preparar.add_argument(
        "--compromisso",
        default=None,
        help="Caminho do arquivo de Compromisso do SAP. Se omitido, procura na --pasta um arquivo com 'Compromisso' no nome.",
    )
    p_preparar.add_argument(
        "--pasta",
        default=".",
        help="Pasta onde procurar os arquivos do SAP quando --realizados/--compromisso não forem informados (padrão: pasta atual).",
    )
    p_preparar.add_argument(
        "--po-col-realizados",
        default=core.DEFAULT_PO_COL_REALIZADOS,
        help=f"Nome da coluna de PO na planilha de Realizados (padrão: '{core.DEFAULT_PO_COL_REALIZADOS}').",
    )
    p_preparar.add_argument(
        "--po-col-compromisso",
        default=core.DEFAULT_PO_COL_COMPROMISSO,
        help=f"Nome da coluna de PO na planilha de Compromisso (padrão: '{core.DEFAULT_PO_COL_COMPROMISSO}').",
    )
    p_preparar.add_argument("--objeto-col", default="OBJETO", help="Nome da coluna OBJETO (padrão: OBJETO).")
    p_preparar.add_argument("--aba-realizados", default=0, help="Nome ou índice da aba em Realizados (padrão: primeira aba).")
    p_preparar.add_argument("--aba-compromisso", default=0, help="Nome ou índice da aba em Compromisso (padrão: primeira aba).")
    p_preparar.add_argument("--header-realizados", type=int, default=0, help="Linha (0-indexada) do cabeçalho em Realizados.")
    p_preparar.add_argument("--header-compromisso", type=int, default=0, help="Linha (0-indexada) do cabeçalho em Compromisso.")
    p_preparar.add_argument("--banco", default=str(core.DEFAULT_BANCO), help="Caminho do banco de classificações (padrão: classificacoes.xlsx).")
    p_preparar.add_argument("--saida", required=True, help="Caminho do Excel de trabalho a ser gerado.")
    p_preparar.set_defaults(func=_comando_preparar)

    p_consolidar = subparsers.add_parser("consolidar", help="Salva no banco as classificações preenchidas manualmente no Excel de trabalho.")
    p_consolidar.add_argument("--trabalho", required=True, help="Caminho do Excel de trabalho já classificado.")
    p_consolidar.add_argument(
        "--po-col-realizados",
        default=core.DEFAULT_PO_COL_REALIZADOS,
        help=f"Nome da coluna de PO na aba Realizados do arquivo de trabalho (padrão: '{core.DEFAULT_PO_COL_REALIZADOS}').",
    )
    p_consolidar.add_argument(
        "--po-col-compromisso",
        default=core.DEFAULT_PO_COL_COMPROMISSO,
        help=f"Nome da coluna de PO na aba Compromisso do arquivo de trabalho (padrão: '{core.DEFAULT_PO_COL_COMPROMISSO}').",
    )
    p_consolidar.add_argument("--banco", default=str(core.DEFAULT_BANCO), help="Caminho do banco de classificações (padrão: classificacoes.xlsx).")
    p_consolidar.set_defaults(func=_comando_consolidar)

    return parser


def main() -> None:
    parser = construir_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
