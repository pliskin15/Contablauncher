"""Painel unificado de Adiantamentos: Balancete × Financeiro e Balancete × Adiantamento."""

import tkinter as tk
from tkinter import ttk

from core import theme
from adiantamento.balancete_financeiro import App as BalanceteFinanceiro
from adiantamento.vs_balancete import CruzamentoFrame


def _build_content(win) -> None:
    win.title("Adiantamentos")
    win.geometry("1180x720")
    win.minsize(900, 560)
    theme.apply_theme(win)

    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True)

    aba1 = BalanceteFinanceiro(nb)
    aba2 = CruzamentoFrame(nb)
    nb.add(aba1, text="Balancete × Financeiro")
    nb.add(aba2, text="Balancete × Adiantamento")


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()
