"""Painel unificado de notas fiscais: Pendentes e Em Atraso."""

import tkinter as tk
from tkinter import ttk

from core import theme
from notas.pendentes import NotasPendentesFrame
from notas.em_atraso import NotasEmAtrasoFrame


def _build_content(win) -> None:
    win.title("Notas Fiscais")
    win.geometry("720x620")
    win.minsize(620, 560)
    theme.apply_theme(win)

    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True)

    nb.add(NotasPendentesFrame(nb), text="Pendentes")
    nb.add(NotasEmAtrasoFrame(nb), text="Em Atraso")


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()
