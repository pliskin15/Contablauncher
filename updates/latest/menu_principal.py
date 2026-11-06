"""
menu_principal.py
Menu geral do sistema — ponto de entrada principal.
Ponto de entrada único do app — janela no visual do Prisma (cabeçalho + guias), ver core/shell.py.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import importlib
import traceback
# Imports estaticos — PyInstaller empacota todas as dependencias automaticamente
import conciliacao.cartao
import conciliacao.conciliacao_contabil
import conciliacao.pix_qrcode
import brinks.painel
import prossegur.painel
import json
import os
from core import ui_theme as theme
from core import auth, login, versao, usuarios_painel, shell
from core.theme import (
    T, aplicar_estilos_ttk, botao_tema,
    registrar_callback, tema_atual,
)
from core.lojas import _open_lojas
import brinks.depositos
import brinks.creditos
import prossegur.depositos
import prossegur.creditos

# ══════════════════════════════════════════════════════════════════════════════
#  COFRE — lógica de mapeamento
# ══════════════════════════════════════════════════════════════════════════════

MAPA_JSON = os.path.join(os.path.dirname(__file__), "codigos_cofres.json")

def _carregar_mapa() -> dict:
    if os.path.exists(MAPA_JSON):
        with open(MAPA_JSON, encoding="utf-8") as f:
            return json.load(f)
    return {"lojas": [], "brinks_map": {}, "prossegur_map": {}}


# ══════════════════════════════════════════════════════════════════════════════
#  COFRE — Frame reutilizável (era _CofreFrame em menu_principal.py)
# ══════════════════════════════════════════════════════════════════════════════

class _CofreFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=theme.BG_APP)
        
        # DEBUG — apague depois
        import sys, traceback
        log_path = os.path.join(os.path.dirname(sys.executable), "debug.log")
        
        try:
            from brinks.depositos import BrinksDepositos
            from brinks.creditos import BrinksCreditos
            from brinks.painel import BrinksPainel
            from prossegur.depositos import ProssegurDepositos
            from prossegur.creditos import ProssegurCreditos
            from prossegur.painel import ProssegurPainel
            with open(log_path, "a") as f:
                f.write("IMPORTS OK\n")
        except Exception as e:
            with open(log_path, "a") as f:
                f.write(f"IMPORT ERRO:\n{traceback.format_exc()}\n")

        self._mods = {
            ("Prossegur", "Painel"):    ProssegurPainel,
            ("Prossegur", "Depósitos"): ProssegurDepositos,
            ("Prossegur", "Créditos"):  ProssegurCreditos,
            ("Brinks",    "Painel"):    BrinksPainel,
            ("Brinks",    "Depósitos"): BrinksDepositos,
            ("Brinks",    "Créditos"):  BrinksCreditos,
        }

        theme.apply_theme(master)
        self._build_ui()

    def _build_ui(self):
        toolbar = ttk.Frame(self, style="Toolbar.TFrame", height=52)
        toolbar.pack(fill="x", side="top")
        toolbar.pack_propagate(False)

        ttk.Label(toolbar, text="🏦 Cofre",
                  style="Toolbar.TLabel",
                  font=("Segoe UI Semibold", 14)).pack(side="left", padx=20, pady=10)
        ttk.Label(toolbar, text=versao.texto(), style="Toolbar.TLabel",
                  font=("Segoe UI", 10)).pack(side="right", padx=20)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        body = ttk.Frame(self, style="App.TFrame")
        body.pack(fill="both", expand=True)

        self._sidebar = ttk.Frame(body, style="Card.TFrame", width=220)
        self._sidebar.pack(fill="y", side="left")
        self._sidebar.pack_propagate(False)

        self._content = ttk.Frame(body, style="App.TFrame")
        self._content.pack(fill="both", expand=True, side="left")

        status_bar = ttk.Frame(self, style="Status.TFrame", height=26)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self._status_var = tk.StringVar(value="Pronto.")
        tk.Label(status_bar, textvariable=self._status_var,
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=12, pady=4)

        self._active_btn = None
        self._home_btn   = None
        self._build_sidebar()
        self._go_home()

    def _build_sidebar(self):
        self._home_btn = self._nav_btn(
            self._sidebar, "🏠  Início", self._go_home,
            padx=10, pady=(10, 2), fill="x")

        ttk.Separator(self._sidebar, orient="horizontal").pack(
            fill="x", padx=16, pady=(10, 0))

        self._section_label(self._sidebar, "PROSSEGUR", padx=16, pady=(18, 4))
        for label, mod in [("📊  Painel", "Painel"),
                            ("💰  Depósitos", "Depósitos"),
                            ("📋  Créditos",  "Créditos")]:
            self._nav_btn(self._sidebar, label,
                          lambda m=mod: self._open_module("Prossegur", m),
                          padx=10, pady=2, fill="x")

        ttk.Separator(self._sidebar, orient="horizontal").pack(
            fill="x", padx=16, pady=14)

        self._section_label(self._sidebar, "BRINKS", padx=16, pady=(2, 4))
        for label, mod in [("📊  Painel", "Painel"),
                            ("💰  Depósitos", "Depósitos"),
                            ("📋  Créditos",  "Créditos")]:
            self._nav_btn(self._sidebar, label,
                          lambda m=mod: self._open_module("Brinks", m),
                          padx=10, pady=2, fill="x")

        spacer = ttk.Frame(self._sidebar, style="Card.TFrame")
        spacer.pack(fill="both", expand=True)

        ttk.Separator(self._sidebar, orient="horizontal").pack(fill="x", padx=16)

        self._nav_btn(self._sidebar, "📤  Exportar Mapeamento",
                      self._exportar_mapeamento,
                      padx=10, pady=(6, 2), fill="x")

    def _section_label(self, parent, text, **pack_opts):
        ttk.Label(parent, text=text,
                  background=theme.CARD_BG, foreground=theme.FG_MUTED,
                  font=("Segoe UI Semibold", 9)).pack(anchor="w", **pack_opts)

    def _nav_btn(self, parent, text, command, **pack_opts):
        btn = tk.Button(parent, text=text,
                        bg=theme.CARD_BG, fg=theme.FG_TEXT,
                        activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                        relief="flat", bd=0, cursor="hand2",
                        anchor="w", padx=12, pady=7,
                        font=("Segoe UI", 11))
        btn.pack(**pack_opts)
        btn.configure(command=lambda b=btn, c=command: self._nav_click(b, c))
        btn.bind("<Enter>", lambda e, b=btn: self._btn_hover(b, True))
        btn.bind("<Leave>", lambda e, b=btn: self._btn_hover(b, False))
        return btn

    def _btn_hover(self, btn, entering):
        if btn is self._active_btn:
            return
        btn.configure(bg=theme.SURFACE2 if entering else theme.CARD_BG)

    def _nav_click(self, btn, command):
        if self._active_btn:
            self._active_btn.configure(bg=theme.CARD_BG, fg=theme.FG_TEXT)
        self._active_btn = btn
        if btn:
            btn.configure(bg=theme.PRIMARY, fg="white")
        command()

    def _clear_content(self):
        for w in self._content.winfo_children():
            w.destroy()

    def _go_home(self):
        if self._active_btn and self._active_btn is not self._home_btn:
            self._active_btn.configure(bg=theme.CARD_BG, fg=theme.FG_TEXT)
        self._active_btn = self._home_btn
        if self._home_btn:
            self._home_btn.configure(bg=theme.PRIMARY, fg="white")
        self._show_home()

    def _show_home(self):
        """Início: já exibe o Painel da Prossegur e o da Brinks, em abas."""
        self._clear_content()
        self._status_var.set("Início  ›  Painéis")

        nb = ttk.Notebook(self._content)
        nb.pack(fill="both", expand=True)
        self._home_nb = nb

        # cada aba é criada só quando aberta pela primeira vez
        abas = {}
        for empresa in ("Prossegur", "Brinks"):
            area = ttk.Frame(nb, style="App.TFrame")
            nb.add(area, text=f"  📊  {empresa}  ")
            abas[str(area)] = (empresa, area, [False])

        def _carregar(_event=None):
            empresa, area, feito = abas[nb.select()]
            self._status_var.set(f"Início  ›  Painel {empresa}")
            if feito[0]:
                return
            feito[0] = True
            cls = self._mods.get((empresa, "Painel"))
            try:
                cls(area).pack(fill="both", expand=True)
            except Exception:
                tk.Label(area, text=traceback.format_exc(), bg=theme.BG_APP, fg="red",
                         font=("Courier", 8), justify="left",
                         wraplength=800).pack(padx=10, pady=10, anchor="nw")

        nb.bind("<<NotebookTabChanged>>", _carregar)
        nb.after(50, _carregar)

    def _open_module(self, empresa: str, modulo: str):
        self._clear_content()
        self._status_var.set(f"{empresa}  ›  {modulo}")

        cor_empresa = theme.ACCENT_CYAN if empresa == "Prossegur" else theme.ACCENT_ORANGE

        header = ttk.Frame(self._content, style="Card.TFrame")
        header.pack(fill="x")
        tk.Label(header, text=f"  {empresa}  ›  {modulo}",
                 bg=theme.CARD_BG, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 13), pady=12, anchor="w"
                 ).pack(fill="x", padx=8)
        tk.Frame(header, bg=cor_empresa, height=3).pack(fill="x")

        area = ttk.Frame(self._content, style="App.TFrame")
        area.pack(fill="both", expand=True)

        cls = self._mods.get((empresa, modulo))
        if cls:
            try:
                cls(area).pack(fill="both", expand=True)
            except Exception:
                tk.Label(area,
                         text=traceback.format_exc(),
                         bg=theme.BG_APP, fg="red",
                         font=("Courier", 8),
                         justify="left",
                         wraplength=800).pack(padx=10, pady=10, anchor="nw")
            return

        tk.Label(area,
                 text=f"cls é None para ({empresa}, {modulo})\nMods:\n" +
                      "\n".join(str(k) for k in self._mods.keys()),
                 bg=theme.BG_APP, fg="orange",
                 font=("Courier", 9),
                 justify="left").pack(padx=10, pady=10)
        
    def _exportar_mapeamento(self):
        xlsx_path = filedialog.askopenfilename(
            title="Selecionar planilha de cofres",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not xlsx_path:
            return
        try:
            import pandas as pd
            df = pd.read_excel(xlsx_path, dtype=str)
            df.columns = [c.strip() for c in df.columns]
            required = {"UF", "LOJA", "NOME", "COFRE"}
            missing = required - set(df.columns)
            if missing:
                messagebox.showerror(
                    "Colunas não encontradas",
                    f"A planilha não contém: {', '.join(missing)}\n"
                    f"Colunas encontradas: {', '.join(df.columns)}")
                return
            lojas, brinks_map, prossegur_map = [], {}, {}
            for _, row in df.iterrows():
                uf    = str(row["UF"]).strip()
                loja  = str(row["LOJA"]).strip()
                nome  = str(row["NOME"]).strip()
                cofre = str(row["COFRE"]).strip()
                lojas.append({"uf": uf, "loja": loja, "nome": nome, "cofre": cofre})
                brinks_map[cofre] = loja
                prossegur_map.setdefault(nome, [])
                if loja not in prossegur_map[nome]:
                    prossegur_map[nome].append(loja)
            output = {"lojas": lojas, "brinks_map": brinks_map, "prossegur_map": prossegur_map}
            with open(MAPA_JSON, "w", encoding="utf-8") as f:
                json.dump(output, f, ensure_ascii=False, indent=2)
            self._status_var.set(
                f"Mapeamento exportado  —  {len(lojas)} lojas  "
                f"({len(brinks_map)} cofres Brinks, {len(prossegur_map)} nomes Prossegur)")
            messagebox.showinfo(
                "Exportação concluída",
                f"JSON gerado com sucesso!\n\n"
                f"  Lojas:            {len(lojas)}\n"
                f"  Cofres Brinks:    {len(brinks_map)}\n"
                f"  Nomes Prossegur:  {len(prossegur_map)}\n\n"
                f"Arquivo: {MAPA_JSON}")
        except Exception as exc:
            messagebox.showerror("Erro na exportação", str(exc))


# ══════════════════════════════════════════════════════════════════════════════
#  CONCILIADOR — cards
# ══════════════════════════════════════════════════════════════════════════════

_MODULOS_CONCILIADOR = [
    {
        "titulo":     "Conciliação Cartões",
        "subtitulo":  "Débito / Crédito — Getnet",
        "icone":      "💳",
        "modulo":     "conciliacao.cartao",
        "funcao":     "Conciliacao_Cartao",
        "cor_acento": "AZUL",
    },
    {
        "titulo":     "Conciliação PIX Maquineta",
        "subtitulo":  "PIX via terminal Getnet",
        "icone":      "📲",
        "modulo":     "conciliacao.conciliacao_contabil",
        "funcao":     "ConciliacaoContabil",
        "cor_acento": "VERDE",
    },
    {
        "titulo":     "Conciliação PIX QrCode",
        "subtitulo":  "PIX via QrCode",
        "icone":      "📲",
        "modulo":     "conciliacao.pix_qrcode",
        "funcao":     "Conciliacao_Pix",
        "cor_acento": "AMARELO",
    },
    
]


# Mapa estatico modulo -> funcao (PyInstaller detecta todos os imports)
_MAPA_MODULOS = {
    "conciliacao.cartao":        conciliacao.cartao.Conciliacao_Cartao,
    "conciliacao.conciliacao_contabil": conciliacao.conciliacao_contabil.ConciliacaoContabil,
    "conciliacao.pix_qrcode":    conciliacao.pix_qrcode.Conciliacao_Pix,
    "brinks.painel":             brinks.painel.BrinksPainel,
    "prossegur.painel":          prossegur.painel.ProssegurPainel,
}


def _abrir_modulo_conciliador(mod_info: dict, master: tk.Tk) -> None:
    nome_mod  = mod_info["modulo"]
    nome_func = mod_info["funcao"]

    try:
        func = _MAPA_MODULOS.get(nome_mod)
        if func is None:
            # Fallback: tenta importlib para modulos nao mapeados
            mod  = importlib.import_module(nome_mod)
            func = getattr(mod, nome_func)
        func(master)
    except ModuleNotFoundError as e:
        messagebox.showerror(
            "Módulo não encontrado",
            f"Erro: {e}")
    except AttributeError:
        messagebox.showerror(
            "Função não encontrada",
            f"A função '{nome_func}' não existe em '{nome_mod}.py'.")
    except Exception:
        messagebox.showerror(
            "Erro ao abrir módulo",
            f"Ocorreu um erro ao abrir '{mod_info['titulo']}':\n\n"
            + traceback.format_exc())


class _CardModulo(tk.Frame):
    def __init__(self, parent, mod_info: dict, master_root: tk.Tk, **kwargs):
        super().__init__(parent, **kwargs)
        self._mod      = mod_info
        self._win_root = master_root
        self._build()
        self._bind_hover()

    def _build(self):
        self.config(bg=T("PAINEL"), relief="flat", bd=0,
                    padx=16, pady=14, cursor="hand2")

        topo = tk.Frame(self, bg=T("PAINEL"))
        topo.pack(fill="x")

        tk.Label(topo, text=self._mod["icone"], font=("Segoe UI", 22),
                 bg=T("PAINEL"), fg=T(self._mod["cor_acento"])
                 ).pack(side="left", padx=(0, 10))

        titulo_frame = tk.Frame(topo, bg=T("PAINEL"))
        titulo_frame.pack(side="left", fill="x", expand=True)

        tk.Label(titulo_frame, text=self._mod["titulo"],
                 font=("Segoe UI", 11, "bold"),
                 bg=T("PAINEL"), fg=T("ACENTO"), anchor="w").pack(fill="x")

        tk.Label(titulo_frame, text=self._mod["subtitulo"],
                 font=("Segoe UI", 8),
                 bg=T("PAINEL"), fg=T("TEXTO_SEC"), anchor="w").pack(fill="x")

        sep = tk.Frame(self, bg=T(self._mod["cor_acento"]), height=2)
        sep.pack(fill="x", pady=(10, 0))

        self._todos = ([self, topo, titulo_frame, sep]
                       + list(topo.winfo_children())
                       + list(titulo_frame.winfo_children()))

    def _bind_hover(self):
        for w in self._todos:
            w.bind("<Enter>",    self._on_enter)
            w.bind("<Leave>",    self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, _e=None):
        for w in self._todos:
            try: w.config(bg=T("BORDA"))
            except Exception: pass

    def _on_leave(self, _e=None):
        for w in self._todos:
            try: w.config(bg=T("PAINEL"))
            except Exception: pass

    def _on_click(self, _e=None):
        _abrir_modulo_conciliador(self._mod, self._win_root)

    def reaplicar_tema(self):
        for child in self.winfo_children():
            child.destroy()
        self._build()
        self._bind_hover()
        self.config(bg=T("PAINEL"))


class _ConciliadorFrame:
    def __init__(self, root: tk.Tk):
        self._win    = root
        self._cards  = []
        self._style  = ttk.Style(root)

        self._win.title("")
        self._win.geometry("560x440")
        self._win.resizable(False, False)
        self._win.configure(bg=T("BG"))

        self._win.update_idletasks()
        w, h = self._win.winfo_width(), self._win.winfo_height()
        x = (self._win.winfo_screenwidth()  - w) // 2
        y = (self._win.winfo_screenheight() - h) // 2
        self._win.geometry(f"+{x}+{y}")

        self._build()
        registrar_callback(self._reaplicar_tema)
        aplicar_estilos_ttk(self._style)

    def _build(self):
        self._topbar = tk.Frame(self._win, bg=T("PAINEL"), pady=10)
        self._topbar.pack(fill="x")

        tk.Label(self._topbar, text="Conciliador",
                 font=("Segoe UI", 13, "bold"),
                 bg=T("PAINEL"), fg=T("ACENTO")).pack(side="left", padx=16)

        self._btn_tema = botao_tema(self._topbar)
        self._btn_tema.pack(side="right", padx=12)

        self._lbl_sub = tk.Label(self._win,
                                 text="Selecione uma conciliação para abrir",
                                 font=("Segoe UI", 9),
                                 bg=T("BG"), fg=T("TEXTO_SEC"))
        self._lbl_sub.pack(pady=(14, 6))

        self._frame_cards = tk.Frame(self._win, bg=T("BG"))
        self._frame_cards.pack(fill="both", expand=True, padx=24, pady=6)

        self._montar_cards()

        self._rodape = tk.Frame(self._win, bg=T("PAINEL"), pady=6)
        self._rodape.pack(fill="x", side="bottom")

        self._lbl_rodape = tk.Label(self._rodape, text=versao.texto(),
                                    font=("Segoe UI", 7),
                                    bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        self._lbl_rodape.pack()

    def _montar_cards(self):
        for child in self._frame_cards.winfo_children():
            child.destroy()
        self._cards.clear()
        for mod in _MODULOS_CONCILIADOR:
            card = _CardModulo(self._frame_cards, mod_info=mod,
                               master_root=self._win)
            card.pack(fill="x", pady=5)
            self._cards.append(card)

    def _reaplicar_tema(self):
        aplicar_estilos_ttk(self._style)
        self._win.configure(bg=T("BG"))
        self._topbar.config(bg=T("PAINEL"))
        for w in self._topbar.winfo_children():
            try: w.config(bg=T("PAINEL"), fg=T("ACENTO"))
            except Exception: pass
        try: self._btn_tema.config(bg=T("BORDA"), fg=T("TEXTO"))
        except Exception: pass
        self._lbl_sub.config(bg=T("BG"), fg=T("TEXTO_SEC"))
        self._frame_cards.config(bg=T("BG"))
        self._rodape.config(bg=T("PAINEL"))
        self._lbl_rodape.config(bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        for card in self._cards:
            card.reaplicar_tema()


# ══════════════════════════════════════════════════════════════════════════════
#  Factories para o MenuPrincipal
# ══════════════════════════════════════════════════════════════════════════════

def _open_cofre(root: tk.Tk, on_close):
    win = tk.Toplevel(root)
    win.title("")
    win.geometry("960x600")
    win.minsize(800, 520)
    win.resizable(True, True)
    win.configure(bg=theme.BG_APP)
    theme.apply_theme(win)

    app_frame = _CofreFrame(win)
    app_frame.pack(fill="both", expand=True)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_conciliador(root: tk.Tk, on_close):
    win = tk.Toplevel(root)
    win.title("Conciliador")
    _ConciliadorFrame(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win

def _open_importacao(root: tk.Tk, on_close):
    import core.app as imp_app

    win = imp_app.App(parent=None)   # <-- parent=None evita grab_set e transient
    win.title("Importações")
    theme.apply_theme(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win

def _open_livro_saida(root: tk.Tk, on_close):
    import livro_fiscal  # seu novo arquivo

    win = tk.Toplevel(root)
    win.title("Livro Fiscal - Saída")
    win.geometry("1100x720")
    win.minsize(900, 580)
    win.configure(bg=theme.BG_APP)

    theme.apply_theme(win)

    livro_fiscal._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_conciliador_contas(root: tk.Tk, on_close):
    from conciliador import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_notas(root: tk.Tk, on_close):
    from notas import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_adiantamentos(root: tk.Tk, on_close):
    from adiantamento import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_tef(root: tk.Tk, on_close):
    from conciliacao import tef_autorizacao as _tef

    win = tk.Toplevel(root)
    _tef._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_conferencia(root: tk.Tk, on_close):
    import conferencia_legado_protheus as _cf

    win = tk.Toplevel(root)
    _cf._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_energia(root: tk.Tk, on_close):
    import energia

    win = tk.Toplevel(root)
    energia._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_folha_pagamento(root: tk.Tk, on_close):
    import fopag_v2

    win = tk.Toplevel(root)
    fopag_v2._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_transmit(root: tk.Tk, on_close):
    import transmit

    win = tk.Toplevel(root)
    transmit._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_dpi(root: tk.Tk, on_close):
    import DPI

    win = tk.Toplevel(root)
    DPI.ConsultaDAIApp(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_gerador_txt(root: tk.Tk, on_close):
    from protheus import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_cte(root: tk.Tk, on_close):
    from cte import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_tarifas(root: tk.Tk, on_close):
    import tarifas

    win = tk.Toplevel(root)
    tarifas._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


def _open_obrigacoes(root: tk.Tk, on_close):
    from obrigacoes import painel as _pnl

    frame = _pnl.abrir_modulo(root, auth.usuario_publico(), on_close)
    if frame is None:       # o módulo já mostrou o motivo e fechou a própria janela
        raise RuntimeError("O Controle de Obrigações não pôde ser aberto (veja a mensagem anterior).")
    return frame.winfo_toplevel()


def _open_central_fiscal(root: tk.Tk, on_close):
    from fiscal import painel as _pnl

    win = tk.Toplevel(root)
    _pnl._build_content(win)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


MODULES = [
    {
        "emoji":    "🏦",
        "title":    "Cofre",
        "subtitle": "Painel do Cofre - Prossegur e Brinks",
        "factory":  _open_cofre,
    },
    {
        "emoji":    "🏪",
        "title":    "Vendas GOODCARD",
        "subtitle": "Conciliador para as vendas GOODCARD",
        "factory":  _open_lojas,
    },
    {
        "emoji":    "💳",
        "title":    "Conciliador GETNET e PIX",
        "subtitle": "Conciliação para as vendas GETNET e PIX QRS",
        "factory":  _open_conciliador,
    },
    {
        "emoji":    "📦",
        "title":    "Importações",
        "subtitle": "Geração dos modelos de INPUT para as importações",
        "factory":  _open_importacao,
    },  
    {
        "emoji": "📚",
        "title": "Livro Fiscal Saída",
        "subtitle": "Conferência de vendas e notas fiscais",
        "factory": _open_livro_saida,
    },
    {
        "emoji": "💳",
        "title": "Conciliador de Contas e Livro Entrada",
        "subtitle": "Contas 989/1677 e conferência do livro fiscal de entrada",
        "factory": _open_conciliador_contas,
    },
    {
        "emoji": "🛃",
        "title": "Consulta DAI",
        "subtitle": "Documentos de Importação - SEFAZ-AM (GDDI)",
        "factory": _open_dpi,
    },
    {
        "emoji": "📨",
        "title": "Notas Fiscais",
        "subtitle": "Notas pendentes (base x transmissão) e notas em atraso",
        "factory": _open_notas,
    },
    {
        "emoji": "🚚",
        "title": "Consulta CT-e",
        "subtitle": "Tomador, destinatário, valor e notas a partir da chave do CT-e",
        "factory": _open_cte,
    },
    {
        "emoji": "💰",
        "title": "Adiantamentos",
        "subtitle": "Balancete x Financeiro e Balancete x Adiantamento",
        "factory": _open_adiantamentos,
    },
    {
        "emoji": "⚖️",
        "title": "Conferência Legado x Protheus",
        "subtitle": "Compara totais por filial e CFOP entre Legado e Protheus",
        "factory": _open_conferencia,
    },
    {
        "emoji": "🤖",
        "title": "Robô de Busca",
        "subtitle": "Busca termos no Histórico dos Excel de uma pasta e exporta em CSV",
        "factory": _open_tarifas,
    },
    {
        "emoji": "📅",
        "title": "Controle de Obrigações",
        "subtitle": "Prazos, etapas e entregas fiscais por filial (SPED, DAM, apurações…)",
        "factory": _open_obrigacoes,
    },
    {
        "emoji": "🧾",
        "title": "Central Fiscal",
        "subtitle": "SF1, TES, contas de despesas, devolução, C170 e EFD",
        "factory": _open_central_fiscal,
    },
    {
        "emoji": "🔐",
        "title": "TEF Autorização",
        "subtitle": "Cruza Autorização (Transações) x AUTHO (Banco)",
        "factory": _open_tef,
    },
    {
        "emoji": "📝",
        "title": "Gerador TXT Protheus",
        "subtitle": "Gera o TXT de lançamentos contábeis para importação em lote",
        "factory": _open_gerador_txt,
    },
    {
        "emoji": "⚡",
        "title": "Conciliação de Energia",
        "subtitle": "Concilia Razão x Relatório de NF (PDF) por filial",
        "factory": _open_energia,
    },
    {
        "emoji":    "🧮",
        "title":    "Folha de Pagamento x Contábil",
        "subtitle": "Compara lançamentos por conta (COD_EVENTO) com a Folha Analítica em PDF",
        "factory":  _open_folha_pagamento,
    },
    {
        "emoji": "📤",
        "title": "Conversor NF-e",
        "subtitle": "Banco Geral → Modelo ICMS Unificado",
        "factory": _open_transmit,
    },

]


# ══════════════════════════════════════════════════════════════════════════════
#  MENU PRINCIPAL — janela raiz
# ══════════════════════════════════════════════════════════════════════════════

GRUPOS = [
    ("Conciliações",     ["Conciliador GETNET e PIX", "Vendas GOODCARD", "TEF Autorização",
                          "Conciliador de Contas e Livro Entrada", "Conciliação de Energia",
                          "Adiantamentos", "Robô de Busca"]),
    ("Fiscal",           ["Central Fiscal", "Controle de Obrigações","Livro Fiscal Saída", "Notas Fiscais", "Conferência Legado x Protheus",
                          "Consulta DAI",
                          "Consulta CT-e", "Conversor NF-e"]),
    ("Folha",            ["Folha de Pagamento x Contábil"]),
    ("Cofre e Importações", ["Cofre", "Importações", "Gerador TXT Protheus"]),
]

_CARD_MIN_W = 300   # largura mínima de um card; define quantas colunas cabem
_CARD_PAD   = 6

# Visual dos cards: mesmo desenho do ModuleCard do Prisma V22
_CARD_ICONE_W   = 56   # reserva para o ícone (para calcular a quebra de linha)
_CARD_SETA_W    = 36   # reserva para a seta ›


# ══════════════════════════════════════════════════════════════════════════════
#  MÓDULOS DENTRO DAS GUIAS
#  Cada função monta o módulo num HostFrame (core/shell.py), que faz o papel da janela.
#  O código dos módulos não muda: é a mesma chamada que as factories acima fazem na Toplevel.
#  Módulo sem entrada aqui continua abrindo em janela própria (ex.: Importações).
# ══════════════════════════════════════════════════════════════════════════════

def _embutir_obrigacoes(host):
    from obrigacoes import painel as _pnl
    _pnl.Control(host, auth.usuario_publico()).pack(fill="both", expand=True)


def _embutir_lojas(host):
    from core.lojas import LojasFrame
    LojasFrame(host).pack(fill="both", expand=True)


def _embutir(modulo: str, funcao: str = "_build_content"):
    def _montar(host):
        mod = importlib.import_module(modulo)
        getattr(mod, funcao)(host)
    return _montar


EMBUTIR = {
    "Cofre":                                 lambda h: _CofreFrame(h).pack(fill="both", expand=True),
    "Vendas GOODCARD":                       _embutir_lojas,
    "Conciliador GETNET e PIX":              lambda h: _ConciliadorFrame(h),
    "Livro Fiscal Saída":                    _embutir("livro_fiscal"),
    "Conciliador de Contas e Livro Entrada": _embutir("conciliador.painel"),
    "Consulta DAI":                          _embutir("DPI", "ConsultaDAIApp"),
    "Notas Fiscais":                         _embutir("notas.painel"),
    "Consulta CT-e":                         _embutir("cte.painel"),
    "Adiantamentos":                         _embutir("adiantamento.painel"),
    "Conferência Legado x Protheus":         _embutir("conferencia_legado_protheus"),
    "Robô de Busca":                         _embutir("tarifas"),
    "Controle de Obrigações":                _embutir_obrigacoes,
    "Central Fiscal":                        _embutir("fiscal.painel"),
    "TEF Autorização":                       _embutir("conciliacao.tef_autorizacao"),
    "Gerador TXT Protheus":                  _embutir("protheus.painel"),
    "Conciliação de Energia":                _embutir("energia"),
    "Folha de Pagamento x Contábil":         _embutir("fopag_v2"),
    "Conversor NF-e":                        _embutir("transmit"),
}


def grupos_modulos() -> list[tuple[str, list[str]]]:
    """[(seção, [títulos])] na ordem do menu — usado também no painel de funções."""
    titulos = {m["title"] for m in MODULES}
    grupos = [(nome, [t for t in ts if t in titulos]) for nome, ts in GRUPOS]
    usados = {t for _, ts in grupos for t in ts}
    resto = [m["title"] for m in MODULES if m["title"] not in usados]
    if resto:                                   # módulos novos sem grupo definido
        grupos.append(("Outros", resto))
    return [(nome, ts) for nome, ts in grupos if ts]


def _descendentes(w) -> list:
    out = [w]
    for f in w.winfo_children():
        out.extend(_descendentes(f))
    return out


# ══════════════════════════════════════════════════════════════════════════════
#  TELA: MENU (cards dos módulos) — MenuPrincipal/ModuleCard do Prisma
# ══════════════════════════════════════════════════════════════════════════════

class MenuTela(tk.Frame):
    def __init__(self, workspace: "Workspace"):
        super().__init__(workspace, bg=theme.BG_APP)
        self._ws = workspace
        self._cols = 0
        self._grids: list[tuple[tk.Frame, list[tk.Frame]]] = []
        self._card_labels: list[tuple[tk.Label, tk.Label]] = []
        # rolagem só com o mouse sobre o menu (os módulos usam bind_all para a deles)
        self._tag_rolar = f"rolar_menu_{id(self)}"
        self.bind_class(self._tag_rolar, "<MouseWheel>", self._on_wheel)

        self._canvas = tk.Canvas(self, bg=theme.BG_APP, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._body = tk.Frame(self._canvas, bg=theme.BG_APP)
        self._win_id = self._canvas.create_window((0, 0), window=self._body, anchor="nw")
        self._body.bind("<Configure>", lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", self._on_canvas_resize)

        topo = tk.Frame(self._body, bg=theme.BG_APP)
        topo.pack(fill="x", padx=28, pady=(26, 0))
        shell.rotulo_secao(topo, "Módulos", cor=theme.FG_MUTED).pack(side="left")
        self._secoes = tk.Frame(self._body, bg=theme.BG_APP)
        self._secoes.pack(fill="x", pady=(0, 24))
        self.montar(self._ws.shell.permitidos)

    def montar(self, permitidos: set[str] | None):
        """(Re)monta as seções só com os módulos permitidos (None = todos)."""
        for w in self._secoes.winfo_children():
            w.destroy()
        self._grids.clear()
        self._card_labels.clear()
        por_titulo = {m["title"]: m for m in MODULES}
        visiveis = 0
        for nome, titulos in grupos_modulos():
            mods = [por_titulo[t] for t in titulos if permitidos is None or t in permitidos]
            if mods:
                visiveis += len(mods)
                self._add_section(nome, mods)
        if not visiveis:
            tk.Label(self._secoes, text="Nenhum módulo liberado para a sua função. Fale com um administrador.",
                     bg=theme.BG_APP, fg=theme.FG_MUTED, font=("Segoe UI", 10)).pack(anchor="w", padx=28, pady=18)
        self._cols = 0                              # força o reposicionamento dos cards
        w = self._canvas.winfo_width()
        if w > 1:
            self._on_canvas_resize(type("E", (), {"width": w})())
        self._marcar_rolagem()

    def _marcar_rolagem(self):
        for w in _descendentes(self):
            tags = w.bindtags()
            if self._tag_rolar not in tags:
                w.bindtags((self._tag_rolar,) + tags)

    def _add_section(self, nome: str, mods: list):
        # título de seção: ciano, maiúsculas, negrito e espaçado (assinatura do Prisma)
        shell.rotulo_secao(self._secoes, nome).pack(anchor="w", padx=28, pady=(20, 6))
        grid = tk.Frame(self._secoes, bg=theme.BG_APP)
        grid.pack(fill="x", padx=22)
        self._grids.append((grid, [self._make_card(grid, m) for m in mods]))

    def _make_card(self, parent, mod: dict) -> tk.Frame:
        """Card horizontal (ícone · título/descrição · seta), como o ModuleCard do Prisma."""
        card = tk.Frame(parent, bg=theme.CARD_BG, cursor="hand2", bd=0,
                        highlightthickness=1, highlightbackground=theme.GRID_LINE,
                        highlightcolor=theme.GRID_LINE)
        inner = tk.Frame(card, bg=theme.CARD_BG, padx=16, pady=14)
        inner.pack(fill="both", expand=True)
        tk.Label(inner, text=mod["emoji"], bg=theme.CARD_BG, fg=theme.ACCENT_CYAN,
                 font=("Segoe UI Emoji", 20), width=2).pack(side="left")
        tk.Label(inner, text="›", bg=theme.CARD_BG, fg=theme.FG_DIM,
                 font=("Segoe UI", 18)).pack(side="right")
        textos = tk.Frame(inner, bg=theme.CARD_BG)
        textos.pack(side="left", fill="both", expand=True, padx=(10, 6))
        t = tk.Label(textos, text=mod["title"], bg=theme.CARD_BG, fg=theme.FG_TEXT,
                     font=("Segoe UI", 11, "bold"), justify="left", anchor="w")
        t.pack(anchor="w", fill="x")
        d = tk.Label(textos, text=mod["subtitle"], bg=theme.CARD_BG, fg=theme.FG_MUTED,
                     font=("Segoe UI", 9), justify="left", anchor="w")
        d.pack(anchor="w", fill="x", pady=(3, 0))
        self._card_labels.append((t, d))

        filhos = _descendentes(card)

        def _enter(_e):
            card.configure(highlightbackground=theme.ACCENT, highlightcolor=theme.ACCENT)
            for w in filhos:
                try: w.configure(bg=theme.SURFACE2)
                except Exception: pass

        def _leave(_e):
            card.configure(highlightbackground=theme.GRID_LINE, highlightcolor=theme.GRID_LINE)
            for w in filhos:
                try: w.configure(bg=theme.CARD_BG)
                except Exception: pass

        for w in filhos:
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)
            w.bind("<Button-1>", lambda e, m=mod: self._ws.mostrar(m["title"]))
        return card

    # ── responsividade ───────────────────────────────────────────────────────
    def _on_canvas_resize(self, event):
        self._canvas.itemconfigure(self._win_id, width=event.width)
        cols = max(1, (event.width - 44) // (_CARD_MIN_W + 2 * _CARD_PAD))
        if cols != self._cols:
            self._cols = cols
            self._relayout()
        wrap = max(120, (event.width - 44) // cols - 2 * _CARD_PAD - 32 - _CARD_ICONE_W - _CARD_SETA_W)
        for t, d in self._card_labels:
            t.configure(wraplength=wrap)
            d.configure(wraplength=wrap)

    def _relayout(self):
        for grid, cards in self._grids:
            for c in cards:
                c.grid_forget()
            for i in range(20):
                grid.columnconfigure(i, weight=0, uniform="")
            for i, card in enumerate(cards):
                r, c = divmod(i, self._cols)
                card.grid(row=r, column=c, padx=_CARD_PAD, pady=_CARD_PAD, sticky="nsew")
            for c in range(self._cols):
                grid.columnconfigure(c, weight=1, uniform="card")

    def _on_wheel(self, event):
        if self._body.winfo_reqheight() > self._canvas.winfo_height():
            self._canvas.yview_scroll(int(-event.delta / 120), "units")


# ══════════════════════════════════════════════════════════════════════════════
#  TELA: PERFIL — MenuPerfil do Prisma
# ══════════════════════════════════════════════════════════════════════════════

class PerfilTela(tk.Frame):
    def __init__(self, workspace: "Workspace"):
        super().__init__(workspace, bg=theme.BG_APP)
        self._ws = workspace
        col = tk.Frame(self, bg=theme.BG_APP)
        col.pack(anchor="n", pady=40)
        shell.rotulo_secao(col, "Perfil").pack(anchor="w", pady=(0, 10))

        cartao = tk.Frame(col, bg=theme.CARD_BG, highlightthickness=1, highlightbackground=theme.GRID_LINE)
        cartao.pack(fill="x", pady=(0, 16))
        dentro = tk.Frame(cartao, bg=theme.CARD_BG, padx=18, pady=16)
        dentro.pack(fill="x")
        u = auth.usuario_publico()
        av = shell.Avatar(dentro, lambda: None, tamanho=48)
        av.configure(bg=theme.CARD_BG, cursor="")
        av.set_usuario(u.get("nome", ""))
        av.pack(side="left")
        txt = tk.Frame(dentro, bg=theme.CARD_BG)
        txt.pack(side="left", padx=14)
        tk.Label(txt, text=u.get("nome", ""), bg=theme.CARD_BG, fg=theme.FG_TEXT,
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(txt, text=f"@{u.get('username', '')}", bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(txt, text=shell.espacar(u.get("role", "").upper()), bg=theme.CARD_BG, fg=theme.ACCENT,
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(4, 0))

        def acao(texto, cmd, estilo="primario"):
            b = shell.botao(col, texto, cmd, estilo, width=34)
            b.pack(fill="x", pady=4)
            return b

        sh = self._ws.shell
        if u.get("role") == "admin":
            acao("👥  Usuários e funções", sh.abrir_usuarios)
        acao("🔄  Trocar usuário", sh.trocar_usuario, "secundario")
        acao("← Voltar ao Menu", lambda: self._ws.mostrar("menu"), "secundario")
        tk.Frame(col, bg=theme.BG_APP, height=24).pack()
        acao("Sair do Sistema", sh.sair, "perigo")


# ══════════════════════════════════════════════════════════════════════════════
#  TELA: MÓDULO EMBUTIDO
# ══════════════════════════════════════════════════════════════════════════════

class ModuloTela(tk.Frame):
    """Barra fina (‹ Início · título · ✕ Fechar módulo) + o módulo montado num HostFrame."""

    def __init__(self, workspace: "Workspace", mod: dict):
        super().__init__(workspace, bg=theme.BG_APP)
        self._ws = workspace
        self.mod = mod
        self._fechando = False
        barra = tk.Frame(self, bg=theme.CARD_BG, height=34)
        barra.pack(fill="x")
        barra.pack_propagate(False)
        shell.botao(barra, "‹  Início", lambda: self._ws.mostrar("menu"), "flat",
                    bg_pai=theme.CARD_BG, padx=12, pady=4).pack(side="left", padx=(8, 4))
        tk.Label(barra, text=mod["emoji"], bg=theme.CARD_BG, fg=theme.ACCENT_CYAN,
                 font=("Segoe UI Emoji", 11)).pack(side="left", padx=(6, 4))
        tk.Label(barra, text=mod["title"], bg=theme.CARD_BG, fg=theme.FG_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(barra, text=f"  ·  {mod['subtitle']}", bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left")
        shell.botao(barra, "✕  Fechar módulo", self.fechar, "flat", bg_pai=theme.CARD_BG,
                    padx=12, pady=4).pack(side="right", padx=8)
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")
        self.host = shell.HostFrame(self)
        self.host.pack(fill="both", expand=True)
        # o próprio módulo pode se fechar (destroy na "janela"): volta ao menu
        self.host.bind("<Destroy>", self._host_destruido)
        EMBUTIR[mod["title"]](self.host)

    def _host_destruido(self, event):
        if event.widget is self.host and not self._fechando:
            self.after_idle(lambda: self._ws.descartar(self))

    def fechar(self):
        self._fechando = True
        self._ws.descartar(self)


# ══════════════════════════════════════════════════════════════════════════════
#  WORKSPACE — o conteúdo de uma guia (ui/workspace.py do Prisma)
# ══════════════════════════════════════════════════════════════════════════════

class Workspace(tk.Frame):
    def __init__(self, parent, shell_: "MenuPrincipal"):
        super().__init__(parent, bg=theme.BG_APP)
        self.shell = shell_
        self._telas: dict[str, tk.Frame] = {}
        self.rota = ""
        self.guia = None

    def titulo(self) -> str:
        if self.rota == "menu":
            return "🏠  Início"
        if self.rota == "perfil":
            return "👤  Perfil"
        mod = _MOD_POR_TITULO.get(self.rota)
        return f"{mod['emoji']}  {mod['title']}" if mod else self.rota

    def mostrar(self, rota: str) -> None:
        mod = _MOD_POR_TITULO.get(rota)
        if mod and rota not in EMBUTIR:          # só abre em janela própria
            self.shell.abrir_em_janela(mod)
            return
        tela = self._telas.get(rota)
        if tela is None or not tela.winfo_exists():
            try:
                self.shell.config(cursor="watch")
                self.update_idletasks()
                if rota == "menu":
                    tela = MenuTela(self)
                elif rota == "perfil":
                    tela = PerfilTela(self)
                else:
                    tela = ModuloTela(self, mod)
            except Exception:
                texto = traceback.format_exc()
                for w in self.winfo_children():          # restos de uma tela que falhou ao montar
                    if w not in self._telas.values():
                        try:
                            w.destroy()
                        except Exception:
                            pass
                messagebox.showerror("Erro ao abrir módulo",
                                     f"Ocorreu um erro ao abrir '{rota}':\n\n" + texto, parent=self.shell)
                return
            finally:
                self.shell.config(cursor="")
            self._telas[rota] = tela
        for t in self._telas.values():
            if t is not tela:
                t.pack_forget()
        tela.pack(fill="both", expand=True)
        self.rota = rota
        self.shell.titulo_da_guia(self)

    def descartar(self, tela: tk.Frame) -> None:
        """Fecha uma tela de módulo (o estado dela é perdido) e volta ao menu."""
        for rota, t in list(self._telas.items()):
            if t is tela:
                del self._telas[rota]
        try:
            tela.destroy()
        except Exception:
            pass
        if self.winfo_exists():
            self.mostrar("menu")

    def remontar_menu(self) -> None:
        menu = self._telas.get("menu")
        if menu is not None and menu.winfo_exists():
            menu.montar(self.shell.permitidos)
        perfil = self._telas.pop("perfil", None)          # perfil é refeito (nome/função)
        if perfil is not None:
            perfil.destroy()
            if self.rota == "perfil":
                self.mostrar("perfil")


# ══════════════════════════════════════════════════════════════════════════════
#  JANELA PRINCIPAL — Header + linha de acento + guias (ui/base_window.py do Prisma)
# ══════════════════════════════════════════════════════════════════════════════

class MenuPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self._pronto = False
        self.title("Contabilidade")
        self.palette = theme.apply_theme(self)
        self.configure(bg=theme.BG_APP)
        self.permitidos: set[str] | None = None
        self._workspaces: list[Workspace] = []
        self._atual: Workspace | None = None

        self.header = shell.Header(self, versao.obter(), lambda: self._ir("perfil"))
        self.header.pack(fill="x")
        shell.LinhaAcento(self).pack(fill="x")
        self._faixa = shell.FaixaGuias(self, self._selecionar, self.fechar_guia, lambda: self.nova_guia())
        self._faixa.pack(fill="x")
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")
        self._pilha = tk.Frame(self, bg=theme.BG_APP)
        self._pilha.pack(fill="both", expand=True)

        atalhos = (("<Control-t>", self.nova_guia), ("<Control-T>", self.nova_guia),
                   ("<Control-w>", self._fechar_atual), ("<Control-W>", self._fechar_atual),
                   ("<Control-Tab>", lambda: self._ciclar(1)),
                   ("<Control-Shift-Tab>", lambda: self._ciclar(-1)))
        for seq, fn in atalhos:
            self.bind(seq, lambda _e, f=fn: (f(), "break")[1])
        self._fit_to_screen()
        self._pronto = True

    # ── janela ───────────────────────────────────────────────────────────────
    def title(self, texto=None):
        # módulos embutidos às vezes chamam winfo_toplevel().title(...): vira complemento
        if texto is None or not getattr(self, "_pronto", False) or texto == "Contabilidade":
            return super().title(texto)
        return super().title(f"Contabilidade — {texto}" if texto else "Contabilidade")

    wm_title = title

    def _fit_to_screen(self):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1400, sw - 40), min(900, sh - 80)
        self.minsize(min(900, sw), min(560, sh))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{max((sh - h) // 2 - 20, 0)}")
        try:
            self.state("zoomed")                    # o Prisma abre maximizado
        except tk.TclError:
            pass

    # ── usuário ──────────────────────────────────────────────────────────────
    def set_usuario(self, user: dict | None):
        """Mostra quem está logado e monta os menus só com os módulos da função dele."""
        u = auth.usuario_publico(user)
        self.header.atualizar_usuario(u.get("nome", ""), u.get("role", ""))
        self.atualizar_modulos()
        if not self._workspaces:
            self.nova_guia()

    def atualizar_modulos(self):
        """Recarrega a função do usuário logado e remonta o menu de todas as guias."""
        try:
            self.permitidos = auth.modulos_permitidos()
        except Exception:
            traceback.print_exc()                   # tabelas de funções ausentes: mostra tudo
            self.permitidos = None
        for ws in self._workspaces:
            ws.remontar_menu()

    @staticmethod
    def grupos_modulos() -> list[tuple[str, list[str]]]:
        return grupos_modulos()

    def abrir_usuarios(self):
        usuarios_painel.abrir(self, grupos_modulos(), self.atualizar_modulos)

    def trocar_usuario(self):
        for ws in list(self._workspaces):
            self._remover(ws)
        self.withdraw()
        auth.set_current_user(None)
        try:
            usuario = login.exigir_login(self)
        except Exception:
            messagebox.showerror("Login", "Não foi possível verificar o usuário.\n\n" + traceback.format_exc(),
                                 parent=self)
            usuario = None
        if not usuario:
            self.destroy()
            return
        self.set_usuario(usuario)
        self.deiconify()
        self._fit_to_screen()

    def sair(self):
        if messagebox.askyesno("Sair", "Deseja sair do sistema?", parent=self):
            self.destroy()

    # ── guias ────────────────────────────────────────────────────────────────
    def nova_guia(self, rota: str = "menu") -> Workspace:
        ws = Workspace(self._pilha, self)
        ws.guia = self._faixa.adicionar("🏠  Início")
        self._workspaces.append(ws)
        self._selecionar(len(self._workspaces) - 1)
        ws.mostrar(rota)
        return ws

    def _selecionar(self, i: int) -> None:
        ws = self._workspaces[i]
        if self._atual is not None and self._atual is not ws:
            self._atual.pack_forget()
        ws.pack(fill="both", expand=True)
        self._atual = ws
        self._faixa.marcar_ativa(ws.guia)

    def titulo_da_guia(self, ws: Workspace) -> None:
        if ws.guia is not None:
            ws.guia.set_texto(ws.titulo())

    def _remover(self, ws: Workspace) -> None:
        self._workspaces.remove(ws)
        self._faixa.remover(ws.guia)
        if self._atual is ws:
            self._atual = None
        ws.destroy()

    def fechar_guia(self, i: int) -> None:
        self._remover(self._workspaces[i])
        if not self._workspaces:
            self.nova_guia()                        # sempre fica ao menos a guia Início
        elif self._atual is None:
            self._selecionar(min(i, len(self._workspaces) - 1))

    def _fechar_atual(self):
        if self._atual in self._workspaces:
            self.fechar_guia(self._workspaces.index(self._atual))

    def _ciclar(self, passo: int):
        if self._atual in self._workspaces and len(self._workspaces) > 1:
            self._selecionar((self._workspaces.index(self._atual) + passo) % len(self._workspaces))

    def _ir(self, rota: str):
        if self._atual is not None:
            self._atual.mostrar(rota)

    # ── módulos que só funcionam em janela própria ──────────────────────────
    def abrir_em_janela(self, mod: dict):
        def on_close(win: tk.Toplevel):
            win.destroy()

        try:
            win = mod["factory"](self, on_close)
            win.protocol("WM_DELETE_WINDOW", lambda w=win: on_close(w))
        except Exception:
            messagebox.showerror("Erro ao abrir módulo",
                                 f"Ocorreu um erro ao abrir '{mod['title']}':\n\n" + traceback.format_exc(),
                                 parent=self)


_MOD_POR_TITULO = {m["title"]: m for m in MODULES}


def _iniciar():
    """Abre a janela só depois do login; a função do usuário fica registrada na sessão (core.auth)."""
    app = MenuPrincipal()
    app.withdraw()
    try:
        usuario = login.exigir_login(app)
    except Exception:
        messagebox.showerror("Login", "Não foi possível verificar o usuário.\n\n" + traceback.format_exc(), parent=app)
        usuario = None
    if not usuario:
        app.destroy()
        return
    app.set_usuario(usuario)
    app.deiconify()
    app._fit_to_screen()
    app.mainloop()


if __name__ == "__main__":
    _iniciar()