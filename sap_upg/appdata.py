"""Configuração persistida da interface gráfica (última pasta/arquivo usados).

Guardada em uma pasta fixa no perfil do usuário, para funcionar
independente de onde o programa é executado (script, atalho ou .exe).
"""

from __future__ import annotations

import json
from pathlib import Path

PASTA_APP = Path.home() / "Documents" / "Classificador_UPG"
ARQUIVO_CONFIG = PASTA_APP / "config.json"
BANCO_PADRAO = PASTA_APP / "classificacoes.xlsx"


def carregar_config() -> dict:
    if ARQUIVO_CONFIG.exists():
        try:
            return json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def salvar_config(config: dict) -> None:
    PASTA_APP.mkdir(parents=True, exist_ok=True)
    ARQUIVO_CONFIG.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
