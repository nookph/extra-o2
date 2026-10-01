"""Interface gráfica (ttkbootstrap) do Classificador de UPG.

Feita para quem não quer/precisa abrir terminal: dois passos, cada um com
seu botão de ação. Usa ttkbootstrap por cima do Tkinter (que já vem
embutido no Python) para um visual mais moderno, com tema claro/escuro,
barra de progresso e log colorido.

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
from tkinter import filedialog, messagebox


def _avisar_erro_fatal(titulo: str, mensagem: str) -> None:
    """Mostra uma janela de erro usando só tkinter puro (sem ttkbootstrap).

    Os atalhos de duplo clique abrem via `pythonw`, que não tem janela de
    console: se uma dependência não estiver instalada, o programa some sem
    deixar rastro nenhum. Por isso, qualquer falha logo na importação ou na
    inicialização precisa aparecer aqui, numa janela própria, em vez de só
    fechar sozinho.
    """
    raiz = tk.Tk()
    raiz.withdraw()
    messagebox.showerror(titulo, mensagem)
    raiz.destroy()


try:
    import ttkbootstrap as ttb
    from ttkbootstrap.constants import BOTH, LEFT, PRIMARY, SECONDARY, SUCCESS, X

    from . import appdata, core
except ImportError as _exc:
    _avisar_erro_fatal(
        "Classificador de UPG — dependência faltando",
        "Não consegui abrir o programa porque uma dependência do Python "
        "não está instalada corretamente:\n\n"
        f"{_exc}\n\n"
        "Rode de novo o 'Instalar_Dependencias.bat' (duplo clique) e "
        "confira se ele termina sem mensagens em vermelho antes de "
        "tentar abrir o programa de novo.",
    )
    sys.exit(1)

TEMA = "litera"


class App(ttb.Window):
    def __init__(self) -> None:
        super().__init__(title="Classificador de UPG", themename=TEMA)
        self.geometry("820x820")
        self.minsize(760, 640)

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
        raiz = ttb.Frame(self, padding=16)
        raiz.pack(fill=BOTH, expand=True)

        self._cabecalho(raiz)
        self._passo_selo = self._indicador_etapas(raiz)

        self._card_etapa1(raiz)
        self._card_etapa2(raiz)
        self._card_mensagens(raiz)

    def _cabecalho(self, raiz: ttb.Frame) -> None:
        cabecalho = ttb.Frame(raiz)
        cabecalho.pack(fill=X, pady=(0, 4))

        ttb.Label(
            cabecalho,
            text="Classificador de UPG",
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")
        ttb.Label(
            cabecalho,
            text="Realizados x Compromisso — automação das planilhas do SAP",
            bootstyle=SECONDARY,
        ).pack(anchor="w", pady=(2, 12))

    def _indicador_etapas(self, raiz: ttb.Frame) -> dict[str, ttb.Label]:
        barra = ttb.Frame(raiz)
        barra.pack(fill=X, pady=(0, 12))

        selo1 = ttb.Label(
            barra, text="① Gerar planilha", bootstyle=f"{PRIMARY}-inverse", padding=(10, 4)
        )
        selo1.pack(side=LEFT)
        ttb.Label(barra, text=" → ", bootstyle=SECONDARY).pack(side=LEFT)
        selo2 = ttb.Label(
            barra, text="② Salvar no banco", bootstyle=f"{SECONDARY}-inverse", padding=(10, 4)
        )
        selo2.pack(side=LEFT)
        return {"1": selo1, "2": selo2}

    def _card_etapa1(self, raiz: ttb.Frame) -> None:
        etapa1 = ttb.Labelframe(
            raiz, text="Passo 1 — Gerar a planilha do mês", padding=12, bootstyle=PRIMARY
        )
        etapa1.pack(fill=X, pady=(0, 12))

        self.var_pasta = tk.StringVar(value=self._config.get("pasta", str(Path.home())))
        self._linha_arquivo(
            etapa1, "Pasta com os 2 arquivos do SAP", self.var_pasta, self._escolher_pasta
        )

        nome_sugerido = f"UPG_{date.today().strftime('%Y-%m')}_para_classificar.xlsx"
        pasta_saida_padrao = self._config.get("pasta_saida", str(appdata.PASTA_APP))
        self.var_saida = tk.StringVar(value=str(Path(pasta_saida_padrao) / nome_sugerido))
        self._linha_arquivo(
            etapa1, "Salvar planilha de trabalho como", self.var_saida, self._escolher_saida
        )

        self.var_banco = tk.StringVar(value=self._config.get("banco", str(appdata.BANCO_PADRAO)))
        self._linha_arquivo(
            etapa1, "Banco de classificações (histórico)", self.var_banco, self._escolher_banco
        )

        self.btn_preparar = ttb.Button(
            etapa1,
            text="Gerar planilha do mês",
            command=self._executar_preparar,
            bootstyle=PRIMARY,
        )
        self.btn_preparar.pack(anchor="e", pady=(8, 0))

    def _card_etapa2(self, raiz: ttb.Frame) -> None:
        etapa2 = ttb.Labelframe(
            raiz, text="Passo 2 — Salvar as classificações no banco", padding=12, bootstyle=SECONDARY
        )
        etapa2.pack(fill=X, pady=(0, 12))

        self.var_trabalho = tk.StringVar()
        self._linha_arquivo(
            etapa2,
            "Planilha de trabalho já classificada",
            self.var_trabalho,
            self._escolher_trabalho,
        )

        self.btn_consolidar = ttb.Button(
            etapa2,
            text="Salvar classificações no banco",
            command=self._executar_consolidar,
            bootstyle=SUCCESS,
        )
        self.btn_consolidar.pack(anchor="e", pady=(8, 0))

    def _card_mensagens(self, raiz: ttb.Frame) -> None:
        etapa3 = ttb.Labelframe(raiz, text="Mensagens", padding=8)
        etapa3.pack(fill=BOTH, expand=True)

        self.progresso = ttb.Progressbar(
            etapa3, mode="indeterminate", bootstyle=f"{PRIMARY}-striped"
        )
        # Só é exibida (pack) enquanto uma tarefa está rodando.

        barra_log = ttb.Frame(etapa3)
        barra_log.pack(fill=X, pady=(0, 4))
        ttb.Button(
            barra_log, text="Limpar", command=self._limpar_log, bootstyle=f"{SECONDARY}-link"
        ).pack(side="right")

        self.txt_log = ttb.ScrolledText(etapa3, height=12, wrap="word", auto_hide=True)
        self.txt_log.pack(fill=BOTH, expand=True)
        self.txt_log.text.configure(state="disabled")
        for tag, cor in (("sucesso", "#2fa84f"), ("erro", "#d9534f"), ("info", "#6c757d")):
            self.txt_log.text.tag_configure(tag, foreground=cor)

    def _linha_arquivo(self, container, rotulo, variavel, comando_procurar) -> None:
        linha = ttb.Frame(container)
        linha.pack(fill=X, pady=4)
        ttb.Label(linha, text=rotulo, bootstyle=SECONDARY).pack(anchor="w")

        sublinha = ttb.Frame(linha)
        sublinha.pack(fill=X, pady=(2, 0))
        entrada = ttb.Entry(sublinha, textvariable=variavel)
        entrada.pack(side=LEFT, fill=X, expand=True, padx=(0, 6))
        ttb.Button(
            sublinha, text="Procurar...", command=comando_procurar, bootstyle=f"{SECONDARY}-outline"
        ).pack(side=LEFT)

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

    def _limpar_log(self) -> None:
        self.txt_log.text.configure(state="normal")
        self.txt_log.text.delete("1.0", "end")
        self.txt_log.text.configure(state="disabled")

    def _inserir_log(self, mensagem: str) -> None:
        tag = "info"
        if mensagem.startswith("ERRO"):
            tag = "erro"
        elif "sucesso" in mensagem.lower() or "salva" in mensagem.lower():
            tag = "sucesso"
        self.txt_log.text.configure(state="normal")
        self.txt_log.text.insert("end", mensagem + "\n", tag)
        self.txt_log.text.see("end")
        self.txt_log.text.configure(state="disabled")

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
        if travar:
            self.progresso.pack(fill=X, pady=(0, 8), before=self.txt_log)
            self.progresso.start(12)
        else:
            self.progresso.stop()
            self.progresso.pack_forget()

    def _marcar_etapa_concluida(self, numero: str) -> None:
        selo = self._passo_selo[numero]
        estilo = PRIMARY if numero == "1" else SUCCESS
        selo.configure(bootstyle=f"{estilo}-inverse", text=selo.cget("text") + " ✓")

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
                self._enfileirar(lambda: self._marcar_etapa_concluida("1"))
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
                self._enfileirar(lambda: self._marcar_etapa_concluida("2"))
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
    try:
        app = App()
        app.mainloop()
    except Exception as exc:  # noqa: BLE001 - mostrado direto para o usuário final
        _avisar_erro_fatal(
            "Classificador de UPG — erro ao abrir",
            f"O programa fechou por causa de um erro inesperado:\n\n{exc}",
        )


if __name__ == "__main__":
    main()
