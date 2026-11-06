"""
prossegur/depositos.py
Módulo de Depósitos — Prossegur

Planilha esperada (colunas mínimas):
  - Data da transação   → data do depósito (ex.: "01-01-2026 04:06:00")
  - Nome do Cash Today  → nome do cofre, mapeado via prossegur_map do JSON
  - Tipo de transação   → usado para filtrar apenas "Depósito"
  - Montante            → valor do depósito
  - Nome do cliente     → razão social (opcional, para detalhe)
  - Nome do usuario     → depositante (opcional, para detalhe)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from datetime import datetime

from core import ui_theme as theme
from core import database as db

import sys

def _base_dir() -> str:
    """Retorna a pasta raiz — funciona tanto em dev quanto no exe PyInstaller."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _encontrar_json() -> str:
    """Procura codigos_cofres.json em locais possíveis e retorna o primeiro encontrado."""
    candidatos = [
        os.path.join(_base_dir(), "codigos_cofres.json"),       # mesma pasta do módulo
        os.path.join(_base_dir(), "..", "codigos_cofres.json"), # pasta pai
        os.path.join(os.getcwd(), "codigos_cofres.json"),       # diretório de trabalho
    ]
    for path in candidatos:
        if os.path.exists(path):
            return os.path.normpath(path)
    return os.path.normpath(candidatos[0])  # fallback padrão

MAPA_JSON = _encontrar_json()


# ──────────────────────────────────────────────────────────────────────────── #
#  Helpers de mapeamento                                                        #
# ──────────────────────────────────────────────────────────────────────────── #

def _carregar_mapa() -> dict:
    caminho = _encontrar_json()   # reavalia em tempo de execução
    if os.path.exists(caminho):
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    return {"prossegur_map": {}}


def _lojas_por_nome(nome_cofre: str) -> list[str]:
    """Retorna lista de lojas a partir do nome do cofre Prossegur."""
    mapa = _carregar_mapa()
    return mapa.get("prossegur_map", {}).get(nome_cofre.strip(), [])


def _parse_data(s: str) -> str:
    """
    Normaliza a data da planilha para dd/mm/yyyy.
    Aceita formatos como '01-01-2026 04:06:00', '2026-01-01', '01/01/2026', etc.
    """
    s = str(s).strip()
    formatos = (
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
    )
    for fmt in formatos:
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s   # devolve como está se não reconhecer


# ──────────────────────────────────────────────────────────────────────────── #
#  Janela de detalhe (duplo clique)                                             #
# ──────────────────────────────────────────────────────────────────────────── #

class DetalheDepositoProssegur(tk.Toplevel):
    def __init__(self, parent, data: str, loja: str, lancamentos: list[dict]):
        super().__init__(parent)
        self.title(f"Depósitos Prossegur  —  Loja {loja}  —  {data}")
        self.geometry("720x400")
        self.minsize(580, 300)
        self.resizable(True, True)
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)
        self._build(data, loja, lancamentos)
        self._center(parent)

    def _build(self, data: str, loja: str, lancamentos: list[dict]):
        # cabeçalho
        header = tk.Frame(self, bg=theme.PRIMARY)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"  Loja {loja}   ·   {data}   ·   {len(lancamentos)} lançamento(s)",
            bg=theme.PRIMARY, fg="white",
            font=("Segoe UI Semibold", 11), pady=10, anchor="w",
        ).pack(fill="x", padx=12)

        # tabela
        cols = ("data_hora", "tipo", "valor", "cofre", "depositante", "cliente")
        frame = ttk.Frame(self, style="App.TFrame")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            style="Treeview", selectmode="browse")

        tree.heading("data_hora",   text="Data / Hora")
        tree.heading("tipo",        text="Tipo")
        tree.heading("valor",       text="Montante")
        tree.heading("cofre",       text="Nome do Cofre")
        tree.heading("depositante", text="Depositante")
        tree.heading("cliente",     text="Cliente")

        tree.column("data_hora",   width=145, anchor="center")
        tree.column("tipo",        width=80,  anchor="center")
        tree.column("valor",       width=110, anchor="e")
        tree.column("cofre",       width=180, anchor="w")
        tree.column("depositante", width=130, anchor="w")
        tree.column("cliente",     width=160, anchor="w")

        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview,
                           style="Vertical.TScrollbar")
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        total = 0.0
        for i, lc in enumerate(lancamentos):
            valor = float(lc.get("valor", 0) or 0)
            total += valor
            tag = "par" if i % 2 == 0 else "impar"
            tree.insert("", "end", values=(
                lc.get("data_hora", ""),
                lc.get("tipo", ""),
                f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                lc.get("nome_cofre", ""),
                lc.get("depositante", ""),
                lc.get("cliente", ""),
            ), tags=(tag,))

        tree.tag_configure("par",   background=theme.CARD_BG)
        tree.tag_configure("impar", background=theme.SURFACE2)

        # rodapé com total
        footer = tk.Frame(self, bg=theme.BG_APP, pady=8)
        footer.pack(fill="x")
        total_fmt = (
            f"R$ {total:,.2f}"
            .replace(",", "X").replace(".", ",").replace("X", ".")
        )
        tk.Label(
            footer,
            text=f"Total do dia:   {total_fmt}",
            bg=theme.BG_APP, fg=theme.FG_TEXT,
            font=("Segoe UI Semibold", 11),
        ).pack(side="right", padx=16)

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w}x{h}+{pw - w // 2}+{ph - h // 2}")


# ──────────────────────────────────────────────────────────────────────────── #
#  Frame principal do módulo                                                    #
# ──────────────────────────────────────────────────────────────────────────── #

class ProssegurDepositos(ttk.Frame):

    def __init__(self, master, **kwargs):
        super().__init__(master, style="App.TFrame", **kwargs)
        self._resumo: list[dict] = []
        db.inicializar()
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

        tk.Button(
            toolbar,
            text="Remapear Lojas",
            command=self._remapear_lojas,
            bg=theme.ACCENT_ORANGE, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=6,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left", padx=(0, 10))

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

        # tabela resumo
        cols = ("data", "loja", "qtd", "valor")
        frame_tree = ttk.Frame(self, style="App.TFrame")
        frame_tree.pack(fill="both", expand=True)

        self._tree = ttk.Treeview(
            frame_tree, columns=cols, show="headings",
            style="Treeview", selectmode="browse",
        )

        self._tree.heading("data",  text="Data do Depósito")
        self._tree.heading("loja",  text="Loja")
        self._tree.heading("qtd",   text="Qtd.")
        self._tree.heading("valor", text="Valor Total do Dia")

        self._tree.column("data",  width=160, anchor="center")
        self._tree.column("loja",  width=80,  anchor="center")
        self._tree.column("qtd",   width=60,  anchor="center")
        self._tree.column("valor", width=220, anchor="e")

        sb = ttk.Scrollbar(frame_tree, orient="vertical",
                           command=self._tree.yview,
                           style="Vertical.TScrollbar")
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._tree.bind("<Double-1>", self._on_duplo_clique)

        tk.Label(
            self,
            text="Dê duplo clique em uma linha para ver os lançamentos do dia.",
            bg=theme.BG_APP, fg=theme.FG_MUTED, font=("Segoe UI", 9),
        ).pack(side="bottom", pady=6)

    # ------------------------------------------------------------------ #
    #  Banco → tela                                                         #
    # ------------------------------------------------------------------ #
    def _carregar_do_banco(self):
        self._resumo = db.get_resumo_depositos_prossegur()
        self._preencher_tabela()

    # ------------------------------------------------------------------ #
    #  Importação                                                           #
    # ------------------------------------------------------------------ #
    def _importar(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha Prossegur — Depósitos",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")],
        )
        if not path:
            return

        try:
            # Suporte a CSV também
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)

            df.columns = [c.strip() for c in df.columns]

            # Colunas obrigatórias
            required = {"Data da transação", "Nome do Cash Today", "Montante"}
            missing = required - set(df.columns)
            if missing:
                messagebox.showerror(
                    "Colunas não encontradas",
                    f"Faltando: {', '.join(missing)}\n\n"
                    f"Colunas encontradas:\n{', '.join(df.columns)}",
                )
                return

            # Filtra apenas linhas de Depósito (e Coleta, se quiser incluir)
            # Se a coluna 'Tipo de transação' existir, filtra; senão importa tudo
            if "Tipo de transação" in df.columns:
                df = df[df["Tipo de transação"].str.strip().str.lower() == "depósito"].copy()

            df["Montante"] = pd.to_numeric(
                df["Montante"].str.replace(",", ".") if df["Montante"].dtype == object else df["Montante"],
                errors="coerce"
            ).fillna(0.0)

            # Monta registros para o banco
            registros = []
            nao_mapeados = set()

            for _, row in df.iterrows():
                nome_cofre = str(row["Nome do Cash Today"]).strip()
                data_raw   = str(row["Data da transação"]).strip()
                data_dep   = _parse_data(data_raw)
                montante   = float(row["Montante"])
                tipo       = str(row.get("Tipo de transação", "Depósito")).strip()
                depositante= str(row.get("Nome do usuario",  "")).strip()
                cliente    = str(row.get("Nome do cliente",  "")).strip()
                moeda      = str(row.get("Moeda", "BRL")).strip()

                lojas = _lojas_por_nome(nome_cofre)
                if not lojas:
                    nao_mapeados.add(nome_cofre)
                    # mesmo sem mapeamento, salva com loja=None para não perder
                    lojas = [None]

                for loja in lojas:
                    registros.append({
                        "data_deposito": data_dep,
                        "data_hora":     data_raw,
                        "nome_cofre":    nome_cofre,
                        "loja":          loja,
                        "valor":         montante,
                        "tipo":          tipo,
                        "depositante":   depositante,
                        "cliente":       cliente,
                        "moeda":         moeda,
                    })

            inseridos, ignorados = db.insert_depositos_prossegur(
                registros, os.path.basename(path)
            )

            msg = f"✔  {inseridos} novo(s) inserido(s)"
            if ignorados:
                msg += f"   ·   {ignorados} duplicata(s) ignorada(s)"
            if nao_mapeados:
                msg += f"   ·   ⚠ {len(nao_mapeados)} cofre(s) sem mapeamento"

            self._lbl_arquivo.configure(
                text=msg,
                fg=theme.ACCENT_GREEN if not nao_mapeados else theme.ACCENT_ORANGE,
            )

            if nao_mapeados:
                messagebox.showwarning(
                    "Cofres sem mapeamento",
                    "Os seguintes nomes de cofre não foram encontrados no JSON:\n\n"
                    + "\n".join(sorted(nao_mapeados))
                    + "\n\nVerifique o arquivo codigos_cofres.json.",
                )

            self._carregar_do_banco()

        except Exception as exc:
            messagebox.showerror("Erro ao importar", str(exc))

    # ------------------------------------------------------------------ #
#  Remapeamento de lojas existentes no banco                            #
    # ------------------------------------------------------------------ #
    def _remapear_lojas(self):
        """Percorre todos os registros com loja=None no banco e aplica o mapa atual."""
        mapa = _carregar_mapa().get("prossegur_map", {})
        if not mapa:
            messagebox.showwarning(
                "Mapa vazio",
                "Nenhum mapeamento encontrado no JSON.\nVerifique o arquivo codigos_cofres.json.",
            )
            return

        atualizados = db.remapear_lojas_prossegur(mapa)
        msg = f"✔  {atualizados} registro(s) remapeado(s)"
        self._lbl_arquivo.configure(text=msg, fg=theme.ACCENT_GREEN)
        self._carregar_do_banco()

    # ------------------------------------------------------------------ #
    #  Preenchimento da tabela resumo                                        #
    # ------------------------------------------------------------------ #
    def _preencher_tabela(self):
        for row in self._tree.get_children():
            self._tree.delete(row)

        total_geral = 0.0
        for i, r in enumerate(self._resumo):
            valor_fmt = (
                f"R$ {r['total']:,.2f}"
                .replace(",", "X").replace(".", ",").replace("X", ".")
            )
            tag = "par" if i % 2 == 0 else "impar"
            self._tree.insert("", "end", iid=str(i), values=(
                r["data_corte"],
                r["loja"] if r["loja"] else "—",
                r["qtd"],
                valor_fmt,
            ), tags=(tag,))
            total_geral += r["total"]

        self._tree.tag_configure("par",   background=theme.CARD_BG)
        self._tree.tag_configure("impar", background=theme.SURFACE2)

        total_fmt = (
            f"Total geral:  R$ {total_geral:,.2f}"
            .replace(",", "X").replace(".", ",").replace("X", ".")
        )
        self._lbl_total.configure(text=total_fmt)

    # ------------------------------------------------------------------ #
    #  Duplo clique → detalhe                                               #
    # ------------------------------------------------------------------ #
    def _on_duplo_clique(self, event):
        sel = self._tree.selection()
        if not sel:
            return

        idx    = int(sel[0])
        resumo = self._resumo[idx]
        data   = resumo["data_corte"]
        loja   = resumo["loja"] or "—"

        lancamentos = db.get_lancamentos_dia_prossegur(data, loja)
        DetalheDepositoProssegur(self.winfo_toplevel(), data, loja, lancamentos)