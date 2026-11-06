"""
protheus/painel.py — Gerador de TXT de lançamentos contábeis para o Protheus.

Interface do módulo (Frame embutível no menu). A lógica do layout do TXT
(validar, montar_linha, ler_linha...) está em protheus/txt_layout.py.
"""

import json
import os
import shutil
import tkinter as tk
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import ttk, messagebox, filedialog

from core import theme
from protheus import excel_massa
from protheus.txt_layout import (
    CAMPOS, CONTA_VAZIA, ENCODING, FILIAL_VAZIA, ITEM_VAZIO, LIMPAR_APOS_ADD, LPS, TAM,
    fmt_brl, ler_linha, montar_linha, ultimo_dia_mes_anterior, validar,
)

NOME_RASCUNHO = "lancamentos_rascunho.json"


def pasta_dados() -> Path:
    """Pasta de dados do app (rascunho do gerador): %APPDATA%\\Cofre\\protheus."""
    base = os.environ.get("APPDATA")
    raiz = Path(base) / "Cofre" if base else Path.home() / ".cofre"
    pasta = raiz / "protheus"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def caminho_rascunho() -> Path:
    """Caminho do rascunho; traz para a pasta de dados o arquivo antigo (ao lado do app), se existir."""
    novo = pasta_dados() / NOME_RASCUNHO
    antigo = Path(__file__).resolve().parent.parent / NOME_RASCUNHO
    if not novo.exists() and antigo.exists():
        try:
            shutil.move(str(antigo), str(novo))
        except OSError:
            pass
    return novo


class GeradorFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=theme.BG_APP)
        self.rascunho = caminho_rascunho()
        self.lancamentos: list[dict] = []
        self.editando: int | None = None
        self.vars = {k: tk.StringVar() for k in TAM}

        self._montar_cabecalho()
        self._montar_barra_acoes()      # botões no topo: ficam visíveis em qualquer resolução
        self._montar_form()
        self._montar_grade()
        self._carregar_rascunho()

        if not self.vars["data"].get():
            self.vars["data"].set(ultimo_dia_mes_anterior())
        self.bind("<Destroy>", self._ao_destruir)

    @property
    def _janela(self):
        return self.winfo_toplevel()

    # ---------- layout
    def _montar_cabecalho(self):
        bar = tk.Frame(self, bg=theme.CARD_BG)
        bar.pack(fill="x")
        tk.Label(bar, text="Gerador TXT — Lançamentos Protheus", font=("Segoe UI", 14, "bold"),
                 bg=theme.CARD_BG, fg=theme.FG_TEXT).pack(side="left", padx=20, pady=12)
        tk.Label(bar, text="Importação em lote (lançamento padrão)", font=("Segoe UI", 9),
                 bg=theme.CARD_BG, fg=theme.FG_MUTED).pack(side="left", pady=12)
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")

    def _montar_form(self):
        frm = ttk.LabelFrame(self, text="Lançamento", padding=10)
        frm.pack(fill="x", padx=10, pady=(10, 5))

        # Lançamento padrão (posições 001-003 do TXT): escolha obrigatória, sem valor automático
        style = ttk.Style(self)
        style.configure("LP.Toolbutton", background=theme.SURFACE2, foreground=theme.FG_TEXT,
                        padding=(18, 6), font=("Segoe UI", 10, "bold"), borderwidth=1, relief="flat")
        style.map("LP.Toolbutton",
                  background=[("selected", theme.PRIMARY), ("active", theme.GRID_LINE)],
                  foreground=[("selected", "white")])
        ttk.Label(frm, text="Lançamento Padrão").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=(0, 6))
        lp_box = ttk.Frame(frm)
        lp_box.grid(row=0, column=1, columnspan=3, sticky="w", pady=(0, 6))
        self.lp_botoes = []
        for lp in LPS:
            b = ttk.Radiobutton(lp_box, text=lp, value=lp, variable=self.vars["lp"], style="LP.Toolbutton")
            b.pack(side="left", padx=(0, 6))
            self.lp_botoes.append(b)

        self.entries = {}
        layout = [
            ("data", 1, 0), ("valor", 1, 1),
            ("conta_deb", 2, 0), ("conta_cred", 2, 1),
            ("filial_deb", 3, 0), ("filial_cred", 3, 1),
            ("cc_deb", 4, 0), ("cc_cred", 4, 1),
            ("item_deb", 5, 0), ("item_cred", 5, 1),
        ]
        rotulos = {k: r for k, r, _ in CAMPOS}
        for k, lin, col in layout:
            ttk.Label(frm, text=rotulos[k]).grid(row=lin, column=col * 2, sticky="w", padx=(0, 6), pady=3)
            e = ttk.Entry(frm, textvariable=self.vars[k], width=22)
            e.grid(row=lin, column=col * 2 + 1, sticky="w", pady=3, padx=(0, 30))
            self.entries[k] = e

        ttk.Label(frm, text="Histórico").grid(row=6, column=0, sticky="w", pady=3)
        e = ttk.Entry(frm, textvariable=self.vars["historico"], width=90)
        e.grid(row=6, column=1, columnspan=3, sticky="we", pady=3)
        self.entries["historico"] = e
        self.lbl_hist = ttk.Label(frm, text="0/100", foreground=theme.FG_MUTED)
        self.lbl_hist.grid(row=6, column=4, sticky="w", padx=6)
        self.vars["historico"].trace_add("write", self._contar_hist)

        # filial crédito acompanha a de débito enquanto não for alterada
        self.entries["filial_deb"].bind("<FocusOut>", self._copiar_filial)
        self.entries["cc_deb"].bind("<FocusOut>", lambda e: None)

        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=5, sticky="w", pady=(10, 0))
        self.btn_add = ttk.Button(btns, text="Adicionar (Enter)", command=self.adicionar)
        self.btn_add.pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Cancelar edição", command=self.cancelar_edicao).pack(side="left", padx=6)
        ttk.Button(btns, text="Limpar campos", command=self.limpar_form).pack(side="left", padx=6)

        for e in self.entries.values():
            e.bind("<Return>", lambda ev: self.adicionar())

        self.vars["lp"].trace_add("write", self._aplicar_lp)

    def _aplicar_lp(self, *_):
        """101: contas e filiais dos dois lados obrigatórias |
        102: conta, filial e item de CRÉDITO vazios | 103: conta, filial e item de DÉBITO vazios."""
        lp = self.vars["lp"].get()
        vazios = (CONTA_VAZIA.get(lp), FILIAL_VAZIA.get(lp), ITEM_VAZIO.get(lp))
        for k in ("conta_deb", "conta_cred", "filial_deb", "filial_cred", "item_deb", "item_cred"):
            if k in vazios:
                self.vars[k].set("")
                self.entries[k].state(["disabled"])
            else:
                self.entries[k].state(["!disabled"])

    def _foco_conta(self):
        """Foco na primeira conta que aceita digitação (a bloqueada do LP escolhido é pulada)."""
        k = "conta_cred" if CONTA_VAZIA.get(self.vars["lp"].get()) == "conta_deb" else "conta_deb"
        self.entries[k].focus_set()

    def _montar_barra_acoes(self):
        """Todos os botões em cima, em duas linhas:
        1) arquivos — modelo/importação Excel, abrir TXT … total e GERAR TXT
        2) lista — editar, duplicar, excluir, limpar"""
        barra = ttk.Frame(self, padding=(10, 8, 10, 2))
        barra.pack(fill="x")

        l1 = ttk.Frame(barra)
        l1.pack(fill="x")
        ttk.Button(l1, text="⬇  Baixar modelo Excel", command=self.baixar_modelo).pack(side="left", padx=(0, 6))
        ttk.Button(l1, text="⬆  Importar Excel preenchido...", command=self.importar_excel).pack(side="left", padx=6)
        ttk.Button(l1, text="Abrir TXT...", command=self.abrir_txt).pack(side="left", padx=6)
        ttk.Button(l1, text="Gerar TXT", style="Primary.TButton", command=self.gerar_txt).pack(side="right")
        self.lbl_total = ttk.Label(l1, text="", font=("Segoe UI", 10, "bold"))
        self.lbl_total.pack(side="right", padx=15)

        l2 = ttk.Frame(barra)
        l2.pack(fill="x", pady=(6, 0))
        ttk.Button(l2, text="Editar", command=self.editar).pack(side="left", padx=(0, 6))
        ttk.Button(l2, text="Duplicar", command=self.duplicar).pack(side="left", padx=6)
        ttk.Button(l2, text="Excluir", command=self.excluir).pack(side="left", padx=6)
        ttk.Button(l2, text="Limpar tudo", command=self.limpar_tudo).pack(side="left", padx=6)
        ttk.Label(l2, text="Selecione linhas na lista para editar, duplicar ou excluir (duplo clique edita).",
                  foreground=theme.FG_MUTED).pack(side="left", padx=14)
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x", pady=(8, 0))

    def _montar_grade(self):
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("n", "lp", "data", "conta_deb", "conta_cred", "valor", "historico",
                "filial_deb", "filial_cred", "cc_deb", "cc_cred", "item_deb", "item_cred")
        titulos = ("#", "LP", "Data", "Déb", "Créd", "Valor", "Histórico",
                   "Fil Déb", "Fil Créd", "CC Déb", "CC Créd", "Item Déb", "Item Créd")
        larg = (40, 50, 85, 80, 80, 110, 300, 75, 75, 70, 70, 100, 100)

        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="extended")
        for c, t, w in zip(cols, titulos, larg):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="e" if c in ("valor", "n") else "w",
                             stretch=(c == "historico"))
        sb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", lambda e: self.editar())
        self.tree.bind("<Delete>", lambda e: self.excluir())

    # ---------- eventos do formulário
    def _contar_hist(self, *_):
        n = len(self.vars["historico"].get())
        self.lbl_hist.config(text=f"{n}/100", foreground=theme.ACCENT_RED if n > 100 else theme.FG_MUTED)

    def _copiar_filial(self, _=None):
        if FILIAL_VAZIA.get(self.vars["lp"].get()) == "filial_cred":
            return                              # LP 102: a filial crédito deve ficar vazia
        if not self.vars["filial_cred"].get().strip():
            self.vars["filial_cred"].set(self.vars["filial_deb"].get().upper())

    def _ler_form(self) -> dict:
        return {k: v.get() for k, v in self.vars.items()}

    def limpar_form(self):
        for k in LIMPAR_APOS_ADD:
            self.vars[k].set("")
        self._foco_conta()

    def cancelar_edicao(self):
        self.editando = None
        self.btn_add.config(text="Adicionar (Enter)")
        self.limpar_form()

    # ---------- ações
    def adicionar(self):
        try:
            d = validar(self._ler_form())
        except ValueError as e:
            messagebox.showwarning("Validação", str(e), parent=self._janela)
            return

        if self.editando is not None:
            self.lancamentos[self.editando] = d
            self.editando = None
            self.btn_add.config(text="Adicionar (Enter)")
        else:
            self.lancamentos.append(d)

        self._atualizar()
        self.limpar_form()

    def _selecionados(self) -> list[int]:
        return sorted(int(i) for i in self.tree.selection())

    def editar(self):
        sel = self._selecionados()
        if not sel:
            return
        idx = sel[0]
        d = self.lancamentos[idx]
        for k in TAM:
            val = d[k]
            if k == "valor":
                val = fmt_brl(Decimal(val))
            self.vars[k].set(val)
        self.editando = idx
        self.btn_add.config(text=f"Salvar alteração (linha {idx + 1})")
        self._foco_conta()

    def duplicar(self):
        for idx in self._selecionados():
            self.lancamentos.append(dict(self.lancamentos[idx]))
        self._atualizar()

    def excluir(self):
        sel = self._selecionados()
        if not sel:
            return
        if not messagebox.askyesno("Excluir", f"Excluir {len(sel)} lançamento(s)?", parent=self._janela):
            return
        for idx in reversed(sel):
            del self.lancamentos[idx]
        self.cancelar_edicao()
        self._atualizar()

    def limpar_tudo(self):
        if self.lancamentos and messagebox.askyesno(
                "Limpar tudo", "Apagar todos os lançamentos da lista?", parent=self._janela):
            self.lancamentos.clear()
            self.cancelar_edicao()
            self._atualizar()

    def abrir_txt(self):
        caminho = filedialog.askopenfilename(
            parent=self._janela, filetypes=[("Arquivo texto", "*.txt"), ("Todos", "*.*")])
        if not caminho:
            return
        novos, erros = [], []
        try:
            with open(caminho, encoding=ENCODING, errors="replace") as f:
                for n, linha in enumerate(f, 1):
                    if not linha.strip():
                        continue
                    try:
                        novos.append(validar(ler_linha(linha)))
                    except Exception as e:
                        erros.append(f"Linha {n}: {e}")
        except OSError as e:
            messagebox.showerror("Abrir TXT", f"Não foi possível ler o arquivo.\n\n{e}", parent=self._janela)
            return
        if novos and self.lancamentos and messagebox.askyesno(
                "Abrir TXT", "Substituir a lista atual?\n(Não = acrescentar ao final)", parent=self._janela):
            self.lancamentos.clear()
        self.lancamentos.extend(novos)
        self._atualizar()
        msg = f"{len(novos)} lançamento(s) carregado(s)."
        if erros:
            msg += "\n\nIgnorados:\n" + "\n".join(erros[:15])
        messagebox.showinfo("Abrir TXT", msg, parent=self._janela)

    # ---------- input em massa (Excel)
    def baixar_modelo(self):
        caminho = filedialog.asksaveasfilename(
            parent=self._janela, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile="Modelo_Lancamentos_Protheus.xlsx", title="Salvar modelo de planilha")
        if not caminho:
            return
        try:
            excel_massa.gerar_modelo(caminho)
        except OSError as e:
            messagebox.showerror(
                "Modelo Excel", "Não foi possível gravar o arquivo.\nSe ele estiver aberto no Excel, "
                                f"feche-o e tente de novo.\n\n{e}", parent=self._janela)
            return
        if messagebox.askyesno("Modelo Excel", f"Modelo salvo em:\n{caminho}\n\nAbrir agora?",
                               parent=self._janela):
            try:
                os.startfile(caminho)
            except (AttributeError, OSError):
                pass

    def importar_excel(self):
        caminho = filedialog.askopenfilename(
            parent=self._janela, title="Selecione a planilha preenchida",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not caminho:
            return
        try:
            novos, erros = excel_massa.ler_planilha(caminho)
        except ValueError as e:
            messagebox.showwarning("Importar Excel", str(e), parent=self._janela)
            return
        except Exception as e:
            messagebox.showerror("Importar Excel",
                                 f"Não foi possível ler a planilha.\nSe ela estiver aberta, feche-a e tente de novo.\n\n{e}",
                                 parent=self._janela)
            return

        if not novos:
            msg = "Nenhum lançamento válido na planilha."
            if erros:
                msg += "\n\nIgnorados:\n" + "\n".join(erros[:15])
            messagebox.showwarning("Importar Excel", msg, parent=self._janela)
            return

        if self.lancamentos and messagebox.askyesno(
                "Importar Excel", "Substituir a lista atual?\n(Não = acrescentar ao final)", parent=self._janela):
            self.lancamentos.clear()
        self.lancamentos.extend(novos)
        self.cancelar_edicao()
        self._atualizar()

        total = sum(Decimal(d["valor"]) for d in novos)
        msg = f"{len(novos)} lançamento(s) importado(s).\nTotal: R$ {fmt_brl(total)}"
        if erros:
            msg += f"\n\n{len(erros)} linha(s) ignorada(s) por erro:\n" + "\n".join(erros[:15])
            if len(erros) > 15:
                msg += f"\n… e mais {len(erros) - 15}."
        (messagebox.showwarning if erros else messagebox.showinfo)("Importar Excel", msg, parent=self._janela)

    def _caractere_invalido(self) -> str | None:
        """Procura caractere que o cp1252 não aceita (o arquivo sairia incompleto ao gravar)."""
        for i, d in enumerate(self.lancamentos, 1):
            try:
                montar_linha(d).encode(ENCODING)
            except UnicodeEncodeError as e:
                return (f"Lançamento {i}: o caractere “{e.object[e.start]}” não existe no "
                        f"encoding {ENCODING} usado pelo Protheus. Corrija o histórico/campos.")
        return None

    def gerar_txt(self):
        if not self.lancamentos:
            messagebox.showwarning("Gerar TXT", "Nenhum lançamento na lista.", parent=self._janela)
            return
        problema = self._caractere_invalido()
        if problema:
            messagebox.showwarning("Gerar TXT", problema, parent=self._janela)
            return
        caminho = filedialog.asksaveasfilename(
            parent=self._janela, defaultextension=".txt",
            filetypes=[("Arquivo texto", "*.txt")],
            initialfile=f"IMPORTAR_{datetime.now():%Y%m%d_%H%M}.txt")
        if not caminho:
            return
        linhas = [montar_linha(d) for d in self.lancamentos]
        try:
            with open(caminho, "w", encoding=ENCODING, newline="\r\n") as f:
                f.write("\n".join(linhas) + "\n")
        except OSError as e:
            messagebox.showerror(
                "Gerar TXT", f"Não foi possível gravar o arquivo.\nSe ele estiver aberto em outro "
                             f"programa, feche-o e tente de novo.\n\n{e}", parent=self._janela)
            return

        total = sum(Decimal(d["valor"]) for d in self.lancamentos)
        messagebox.showinfo(
            "Gerar TXT",
            f"Arquivo gerado com {len(linhas)} linha(s).\nTotal: R$ {fmt_brl(total)}\n\n{caminho}",
            parent=self._janela)

    # ---------- grade, totais e rascunho
    def _atualizar(self):
        self.tree.delete(*self.tree.get_children())
        for i, d in enumerate(self.lancamentos):
            self.tree.insert("", "end", iid=str(i), values=(
                i + 1, d["lp"], d["data"], d["conta_deb"], d["conta_cred"],
                fmt_brl(Decimal(d["valor"])), d["historico"],
                d["filial_deb"], d["filial_cred"], d["cc_deb"], d["cc_cred"],
                d["item_deb"], d["item_cred"]))
        total = sum((Decimal(d["valor"]) for d in self.lancamentos), Decimal("0"))
        self.lbl_total.config(text=f"{len(self.lancamentos)} lançamento(s)  |  Total: R$ {fmt_brl(total)}")
        self._salvar_rascunho()

    def _salvar_rascunho(self):
        try:
            dados = {"lancamentos": self.lancamentos,
                     "form": {k: self.vars[k].get() for k in ("lp", "data", "conta_cred", "filial_deb", "filial_cred", "cc_cred", "item_cred")}}
            self.rascunho.write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")
        except (OSError, tk.TclError):
            pass

    def _carregar_rascunho(self):
        if self.rascunho.exists():
            try:
                dados = json.loads(self.rascunho.read_text(encoding="utf-8"))
                self.lancamentos = dados.get("lancamentos", [])
                for d in self.lancamentos:
                    # rascunho gravado antes do campo existir: todos eram 101 (preservado, não é preenchimento novo)
                    d.setdefault("lp", "101")
                    for k in TAM:                      # campos criados depois (ex.: item contábil) ficam vazios
                        d.setdefault(k, "")
                for k, v in dados.get("form", {}).items():
                    if k in self.vars:
                        self.vars[k].set(v)
            except (OSError, json.JSONDecodeError):
                self.lancamentos = []
        self._aplicar_lp()          # a conta que deve ficar vazia no LP restaurado é limpa e bloqueada
        self._atualizar()

    def _ao_destruir(self, ev):
        if ev.widget is self:       # o menu fecha o módulo destruindo a janela; guarda o que está digitado
            self._salvar_rascunho()


def _build_content(win) -> None:
    win.title("Gerador TXT - Lançamentos Protheus")
    # cabe em qualquer resolução: tamanho ideal 1150x650, limitado ao que a tela comporta
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    w, h = min(1150, sw - 40), min(650, sh - 120)
    win.geometry(f"{w}x{h}")
    win.minsize(min(950, w), min(480, h))
    win.configure(bg=theme.BG_APP)
    theme.apply_theme(win)
    GeradorFrame(win).pack(fill="both", expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()
