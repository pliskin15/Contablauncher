"""
prossegur/creditos.py
Módulo de Créditos — Prossegur

Importa o Razão Modelo I (xlsx) e exibe apenas as linhas cujo
HISTORICO seja "NOSSO DEPOSITO - LOJA.XX", mostrando DATA, LOJA e DÉBITO.
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

from core import ui_theme as theme
from core import database as db


# ──────────────────────────────────────────────────────────────────────────── #
#  Helper: extrai número da loja do campo HISTORICO                             #
# ──────────────────────────────────────────────────────────────────────────── #

_RE_LOJA = re.compile(r"NOSSO DEPOSITO\s*-\s*LOJA\.(\d+)", re.IGNORECASE)


def _extrair_loja(historico: str) -> str | None:
    """Retorna '1', '12', etc., ou None se não bater o padrão."""
    m = _RE_LOJA.search(str(historico))
    if m:
        return str(int(m.group(1)))   # remove zero à esquerda → '1', '12'
    return None


# ──────────────────────────────────────────────────────────────────────────── #
#  Frame principal do módulo                                                    #
# ──────────────────────────────────────────────────────────────────────────── #

class ProssegurCreditos(ttk.Frame):

    def __init__(self, master, **kwargs):
        super().__init__(master, style="App.TFrame", **kwargs)
        db.inicializar()
        self._resumo: list[dict] = []
        self._build()
        try:
            self._carregar_do_banco()
        except Exception as exc:
            self._lbl_arquivo.configure(
                text=f"⚠  Banco não inicializado: {exc}",
                fg=theme.ACCENT_ORANGE,
            )

    # ------------------------------------------------------------------ #
    #  Layout                                                               #
    # ------------------------------------------------------------------ #
    def _build(self):
        # ── Toolbar ──────────────────────────────────────────────────── #
        toolbar = tk.Frame(self, bg=theme.CARD_BG, pady=6)
        toolbar.pack(fill="x")

        tk.Button(
            toolbar,
            text="Importar Planilha",
            command=self._importar,
            bg=theme.PRIMARY, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=6,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left", padx=10)

        self._lbl_arquivo = tk.Label(
            toolbar, text="Dados carregados do banco local",
            bg=theme.CARD_BG, fg=theme.FG_MUTED,
            font=("Segoe UI", 9),
        )
        self._lbl_arquivo.pack(side="left", padx=6)

        self._lbl_total = tk.Label(
            toolbar, text="",
            bg=theme.CARD_BG, fg=theme.FG_TEXT,
            font=("Segoe UI Semibold", 10),
        )
        self._lbl_total.pack(side="right", padx=16)

        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")

        # ── Treeview ─────────────────────────────────────────────────── #
        cols = ("data", "loja", "debito")
        frame_tree = ttk.Frame(self, style="App.TFrame")
        frame_tree.pack(fill="both", expand=True)

        self._tree = ttk.Treeview(
            frame_tree, columns=cols, show="headings",
            style="Treeview", selectmode="browse",
        )

        self._tree.heading("data",   text="Data")
        self._tree.heading("loja",   text="Loja")
        self._tree.heading("debito", text="Débito")

        self._tree.column("data",   width=160, anchor="center")
        self._tree.column("loja",   width=100, anchor="center")
        self._tree.column("debito", width=200, anchor="e")

        sb = ttk.Scrollbar(frame_tree, orient="vertical",
                           command=self._tree.yview,
                           style="Vertical.TScrollbar")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        tk.Label(
            self,
            text='Exibindo apenas lançamentos "NOSSO DEPOSITO - LOJA.XX"',
            bg=theme.BG_APP, fg=theme.FG_MUTED, font=("Segoe UI", 9),
        ).pack(side="bottom", pady=6)

    # ------------------------------------------------------------------ #
    #  Banco → tela                                                         #
    # ------------------------------------------------------------------ #
    def _carregar_do_banco(self):
        self._resumo = db.get_creditos_prossegur()
        self._preencher_tabela()

    # ------------------------------------------------------------------ #
    #  Importação                                                           #
    # ------------------------------------------------------------------ #
    def _importar(self):
        path = filedialog.askopenfilename(
            title="Selecionar Razão Modelo I — Prossegur Créditos",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if not path:
            return

        try:
            # Cabeçalho real na linha 4 (índice 3), igual à Brinks
            df = pd.read_excel(path, header=3, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]

            # Colunas obrigatórias
            required = {"DATA", "HISTORICO", "DEBITO"}
            missing = required - set(df.columns)
            if missing:
                messagebox.showerror(
                    "Colunas não encontradas",
                    f"Faltando: {', '.join(missing)}\n\n"
                    f"Colunas detectadas:\n{', '.join(df.columns)}",
                )
                return

            # Filtra apenas linhas de depósito de loja
            df["_loja"] = df["HISTORICO"].apply(_extrair_loja)
            df = df[df["_loja"].notna()].copy()

            if df.empty:
                messagebox.showwarning(
                    "Nenhum lançamento encontrado",
                    'Nenhuma linha com "NOSSO DEPOSITO - LOJA.XX" foi encontrada.',
                )
                return

            df["DEBITO"] = pd.to_numeric(df["DEBITO"], errors="coerce").fillna(0.0)

            registros = []
            for _, row in df.iterrows():
                registros.append({
                    "data":          str(row["DATA"]).strip(),
                    "loja":          row["_loja"],
                    "historico":     str(row["HISTORICO"]).strip(),
                    "debito":        float(row["DEBITO"]),
                    "sequencia":     str(row.get("SEQUENCIA", row.get("SEQUENCI/", ""))).strip(),
                    "lote":          str(row.get("LOTE", "")).strip(),
                    "voucher":       str(row.get("VOUCHER", "")).strip(),
                    "doc_nro":       str(row.get("DOC. NRO.", row.get("DOC.NRO.", ""))).strip(),
                    "centro_custo":  str(row.get("CENTRO DE CUSTO", "")).strip(),
                    "conta_partida": str(row.get("CONTA PARTIDA", "")).strip(),
                    "arquivo_origem": os.path.basename(path),
                })

            inseridos, ignorados = db.insert_creditos_prossegur(registros)

            msg = f"✔  {inseridos} novo(s) inserido(s)"
            if ignorados:
                msg += f"   ·   {ignorados} duplicata(s) ignorada(s)"
            self._lbl_arquivo.configure(text=msg, fg=theme.ACCENT_GREEN)

            self._carregar_do_banco()

        except Exception as exc:
            messagebox.showerror("Erro ao importar", str(exc))

    # ------------------------------------------------------------------ #
    #  Preenchimento da tabela                                              #
    # ------------------------------------------------------------------ #
    def _preencher_tabela(self):
        for row in self._tree.get_children():
            self._tree.delete(row)

        total_geral = 0.0
        for i, r in enumerate(self._resumo):
            debito_fmt = (
                f"R$ {r['debito']:,.2f}"
                .replace(",", "X").replace(".", ",").replace("X", ".")
            )
            tag = "par" if i % 2 == 0 else "impar"
            self._tree.insert("", "end", iid=str(i), values=(
                r["data"], r["loja"], debito_fmt,
            ), tags=(tag,))
            total_geral += r["debito"]

        self._tree.tag_configure("par",   background=theme.CARD_BG)
        self._tree.tag_configure("impar", background=theme.SURFACE2)

        total_fmt = (
            f"Total geral:  R$ {total_geral:,.2f}"
            .replace(",", "X").replace(".", ",").replace("X", ".")
        )
        self._lbl_total.configure(text=total_fmt)
