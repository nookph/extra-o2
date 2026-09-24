"""Interface gráfica simples (Tkinter) do Classificador de UPG.

Feita para quem não quer/precisa abrir terminal: dois botões, um para cada
etapa do mês. O Tkinter já vem embutido no Python, não precisa instalar
nada além do que o `requirements.txt` já pede (pandas, openpyxl).

Para abrir:
    python -m sap_upg.gui
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from typing import Callable
from tkinter import filedialog, messagebox, ttk

from . import appdata, core


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Classificador de UPG")
        self.geometry("760x640")
        self.minsize(700, 560)

        self._config = appdata.carregar_config()
        # As threads de trabalho nunca chamam métodos do Tkinter diretamente
        # (não é seguro fazer isso fora da thread principal). Elas só
        # colocam funções nesta fila; quem executa é sempre o loop principal,
        # via _consumir_fila_ui.
        self._fila_ui: queue.Queue[Callable[[], None]] = queue.Queue()
        self._processando = False

        self._construir_layout()
        self.after(100, self._consumir_fila_ui)

    # ------------------------------------------------------------------ #
    # Layout
    # ------------------------------------------------------------------ #
    def _construir_layout(self) -> None:
        pad = {"padx": 10, "pady": 6}

        titulo = ttk.Label(
            self,
            text="Classificador de UPG (Realizados x Compromisso)",
            font=("Segoe UI", 13, "bold"),
        )
        titulo.pack(anchor="w", **pad)

        # ---- Passo 1: gerar a planilha do mês ----
        etapa1 = ttk.LabelFrame(self, text="Passo 1 — Gerar a planilha do mês")
        etapa1.pack(fill="x", **pad)

        self.var_pasta = tk.StringVar(value=self._config.get("pasta", str(Path.home())))
        self._linha_arquivo(
            etapa1, "Pasta com os 2 arquivos do SAP:", self.var_pasta, self._escolher_pasta
        )

        nome_sugerido = f"UPG_{date.today().strftime('%Y-%m')}_para_classificar.xlsx"
        pasta_saida_padrao = self._config.get("pasta_saida", str(appdata.PASTA_APP))
        self.var_saida = tk.StringVar(value=str(Path(pasta_saida_padrao) / nome_sugerido))
        self._linha_arquivo(
            etapa1, "Salvar planilha de trabalho como:", self.var_saida, self._escolher_saida
        )

        self.var_banco = tk.StringVar(value=self._config.get("banco", str(appdata.BANCO_PADRAO)))
        self._linha_arquivo(
            etapa1, "Banco de classificações (histórico):", self.var_banco, self._escolher_banco
        )

        self.btn_preparar = ttk.Button(
            etapa1, text="Gerar planilha do mês", command=self._executar_preparar
        )
        self.btn_preparar.pack(anchor="e", padx=10, pady=(4, 10))

        # ---- Passo 2: salvar classificações no banco ----
        etapa2 = ttk.LabelFrame(self, text="Passo 2 — Salvar as classificações no banco")
        etapa2.pack(fill="x", **pad)

        self.var_trabalho = tk.StringVar()
        self._linha_arquivo(
            etapa2,
            "Planilha de trabalho já classificada:",
            self.var_trabalho,
            self._escolher_trabalho,
        )

        self.btn_consolidar = ttk.Button(
            etapa2, text="Salvar classificações no banco", command=self._executar_consolidar
        )
        self.btn_consolidar.pack(anchor="e", padx=10, pady=(4, 10))

        # ---- Mensagens ----
        etapa3 = ttk.LabelFrame(self, text="Mensagens")
        etapa3.pack(fill="both", expand=True, **pad)

        self.txt_log = tk.Text(etapa3, height=14, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=8)

    def _linha_arquivo(self, container, rotulo, variavel, comando_procurar) -> None:
        linha = ttk.Frame(container)
        linha.pack(fill="x", padx=10, pady=4)
        ttk.Label(linha, text=rotulo, width=34, anchor="w").pack(side="left")
        entrada = ttk.Entry(linha, textvariable=variavel)
        entrada.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(linha, text="Procurar...", command=comando_procurar).pack(side="left")

    # ------------------------------------------------------------------ #
    # Seletores de arquivo/pasta
    # ------------------------------------------------------------------ #
    def _escolher_pasta(self) -> None:
        inicial = self.var_pasta.get() or str(Path.home())
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com os arquivos do SAP", initialdir=inicial
        )
        if pasta:
            self.var_pasta.set(pasta)

    def _escolher_saida(self) -> None:
        atual = Path(self.var_saida.get())
        caminho = filedialog.asksaveasfilename(
            title="Salvar planilha de trabalho como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=atual.name,
            initialdir=str(atual.parent),
        )
        if caminho:
            self.var_saida.set(caminho)

    def _escolher_banco(self) -> None:
        atual = Path(self.var_banco.get())
        caminho = filedialog.asksaveasfilename(
            title="Selecione (ou crie) o arquivo do banco de classificações",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=atual.name,
            initialdir=str(atual.parent),
        )
        if caminho:
            self.var_banco.set(caminho)

    def _escolher_trabalho(self) -> None:
        inicial = Path(self.var_saida.get()).parent
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de trabalho já classificada",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(inicial),
        )
        if caminho:
            self.var_trabalho.set(caminho)

    # ------------------------------------------------------------------ #
    # Fila thread-safe: threads de trabalho só enfileiram, o loop principal
    # (agendado via self.after) é quem de fato mexe nos widgets.
    # ------------------------------------------------------------------ #
    def _enfileirar(self, acao: Callable[[], None]) -> None:
        self._fila_ui.put(acao)

    def _inserir_log(self, mensagem: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _log(self, mensagem: str) -> None:
        self._enfileirar(lambda: self._inserir_log(mensagem))

    def _consumir_fila_ui(self) -> None:
        try:
            while True:
                acao = self._fila_ui.get_nowait()
                acao()
        except queue.Empty:
            pass
        self.after(100, self._consumir_fila_ui)

    # ------------------------------------------------------------------ #
    # Ações
    # ------------------------------------------------------------------ #
    def _travar_botoes(self, travar: bool) -> None:
        estado = "disabled" if travar else "normal"
        self.btn_preparar.configure(state=estado)
        self.btn_consolidar.configure(state=estado)

    def _executar_preparar(self) -> None:
        if self._processando:
            return
        pasta = self.var_pasta.get().strip()
        saida = self.var_saida.get().strip()
        banco = self.var_banco.get().strip()
        if not saida:
            messagebox.showwarning("Falta informação", "Informe onde salvar a planilha de trabalho.")
            return
        if not pasta:
            messagebox.showwarning("Falta informação", "Selecione a pasta com os arquivos do SAP.")
            return

        Path(saida).parent.mkdir(parents=True, exist_ok=True)
        Path(banco).parent.mkdir(parents=True, exist_ok=True)

        self._processando = True
        self._travar_botoes(True)
        self._log(f"\n--- Gerando planilha do mês ({date.today().strftime('%d/%m/%Y')}) ---")

        def tarefa() -> None:
            try:
                caminho_gerado = core.executar_preparar(
                    pasta=pasta, banco=banco, saida=saida, log=self._log
                )
                self._log(
                    "Preencha a coluna 'Classificacao' nas linhas amarelas (PENDENTE) "
                    "e depois use o Passo 2."
                )
                self._salvar_config(pasta=pasta, pasta_saida=str(Path(saida).parent), banco=banco)
                self._enfileirar(lambda: self.var_trabalho.set(str(caminho_gerado)))
                self._enfileirar(
                    lambda: messagebox.showinfo(
                        "Pronto", f"Planilha gerada com sucesso:\n{caminho_gerado}"
                    )
                )
                self._enfileirar(lambda: self._abrir_arquivo(caminho_gerado))
            except Exception as exc:  # noqa: BLE001 - mostrado direto para o usuário final
                self._log(f"ERRO: {exc}")
                self._enfileirar(lambda exc=exc: messagebox.showerror("Erro ao gerar planilha", str(exc)))
            finally:
                self._processando = False
                self._enfileirar(lambda: self._travar_botoes(False))

        threading.Thread(target=tarefa, daemon=True).start()

    def _executar_consolidar(self) -> None:
        if self._processando:
            return
        trabalho = self.var_trabalho.get().strip()
        banco = self.var_banco.get().strip()
        if not trabalho:
            messagebox.showwarning(
                "Falta informação", "Selecione a planilha de trabalho já classificada."
            )
            return

        self._processando = True
        self._travar_botoes(True)
        self._log("\n--- Salvando classificações no banco ---")

        def tarefa() -> None:
            try:
                _, alterados, total = core.executar_consolidar(
                    trabalho=trabalho, banco=banco, log=self._log
                )
                self._salvar_config(banco=banco)
                self._enfileirar(
                    lambda: messagebox.showinfo(
                        "Pronto",
                        f"{alterados} classificação(ões) nova(s)/alterada(s) salva(s).\n"
                        f"Total de POs no banco: {total}",
                    )
                )
            except Exception as exc:  # noqa: BLE001 - mostrado direto para o usuário final
                self._log(f"ERRO: {exc}")
                self._enfileirar(
                    lambda exc=exc: messagebox.showerror("Erro ao salvar classificações", str(exc))
                )
            finally:
                self._processando = False
                self._enfileirar(lambda: self._travar_botoes(False))

        threading.Thread(target=tarefa, daemon=True).start()

    def _salvar_config(self, **novos) -> None:
        self._config.update(novos)
        appdata.salvar_config(self._config)

    @staticmethod
    def _abrir_arquivo(caminho: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{caminho}"')
            else:
                os.system(f'xdg-open "{caminho}"')
        except OSError:
            pass


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
