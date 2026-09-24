"""Atalho para abrir a interface gráfica com um clique duplo (Windows).

Sem terminal: a extensão .pyw abre com o "pythonw", que não mostra janela
preta de comando.
"""

from sap_upg.gui import main

if __name__ == "__main__":
    main()
