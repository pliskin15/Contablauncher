"""Painel unificado de conciliadores: Conciliador de Contas (989/1677) e Livro de Entradas × Contabilidade."""

import tkinter as tk
from tkinter import ttk

from core import theme
from conciliador import contas, livro_entrada


def _build_content(win) -> None:
    win.title("Conciliador")
    win.geometry("1440x860")
    win.minsize(1100, 680)
    theme.apply_theme(win)

    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True)

    aba_contas = tk.Frame(nb, bg=contas.COR["bg"])
    aba_livro = tk.Frame(nb, bg=livro_entrada.BG)
    nb.add(aba_contas, text="Conciliador de Contas")
    nb.add(aba_livro, text="Livro Fiscal Entrada")

    contas.ConciliadorApp(aba_contas).pack(fill="both", expand=True)
    livro_entrada._build_content(aba_livro)


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()
