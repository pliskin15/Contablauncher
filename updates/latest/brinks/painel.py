"""
brinks/painel.py
Painel de conciliação — Brinks
Grade visual interativa: eixo Y = lojas, eixo X = datas de depósito.
Filtros: status, regional (UF), mês/data.
"""

import json
import os
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import calendar
from core.memorando_service import get_memorandos_por_loja_dia
from core import ui_theme as theme
from core import database as db
import holidays

# ──────────────────────────────────────────────────────────────────────────── #
#  Carrega mapeamento loja → UF do JSON                                        #
# ──────────────────────────────────────────────────────────────────────────── #

import sys

def _base_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _load_loja_uf_map() -> dict[str, str]:
    json_path = os.path.join(_base_dir(), "codigos_cofres.json")
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        return {str(entry["loja"]): entry["uf"] for entry in data.get("lojas", [])}
    except Exception:
        return {}
LOJA_UF: dict[str, str] = _load_loja_uf_map()

# ──────────────────────────────────────────────────────────────────────────── #
#  Helpers de data                                                              #
# ──────────────────────────────────────────────────────────────────────────── #

FERIADOS_BR = holidays.Brazil(years=range(2024, 2030))

def _parse(s: str) -> datetime | None:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None

def _proximo_util(dt: datetime) -> datetime:
    dt = dt + timedelta(days=1)
    while dt.weekday() >= 5 or dt.date() in FERIADOS_BR:
        dt = dt + timedelta(days=1)
    return dt

def _credito_esperado(dt: datetime) -> datetime:
    return _proximo_util(dt)

def _fmt_brl(v: float) -> str:
    return (
        f"R$ {abs(v):,.2f}"
        .replace(",", "X").replace(".", ",").replace("X", ".")
    )

def _fmt_date(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

# ──────────────────────────────────────────────────────────────────────────── #
#  Motor de conciliação                                                         #
# ──────────────────────────────────────────────────────────────────────────── #

def _carregar_memorandos_para(depositos: list[dict]) -> dict[tuple[str, str], float]:
    meses: set[tuple[int, int]] = set()
    lojas: set[str] = set()
    for d in depositos:
        dt = _parse(d.get("data_corte", ""))
        if dt:
            meses.add((dt.year, dt.month))
            lojas.add(str(int(str(d["loja"]))))
    mem_idx: dict[tuple[str, str], float] = {}
    for (ano, mes) in meses:
        parcial = get_memorandos_por_loja_dia(ano, mes, list(lojas))
        mem_idx.update(parcial)
    return mem_idx


def conciliar() -> list[dict]:
    depositos = db.get_resumo_depositos_brinks()
    creditos  = db.get_creditos_brinks()

    def _iguais(a, b, tol=0.01):
        return abs(float(a or 0) - float(b or 0)) <= tol

    # ── memorandos Supabase ────────────────────────────────────────────────
    mem_idx = _carregar_memorandos_para(depositos)

    # ── índice de créditos: (data_credito, loja) → total ──────────────────
    cred_idx: dict[tuple, float] = {}
    for c in creditos:
        key = (c["data"], c["loja"])
        cred_idx[key] = cred_idx.get(key, 0.0) + c["debito"]

    # ── índice de depósitos: (data_deposito DD/MM/YYYY, loja) → total ─────
    dep_idx: dict[tuple, float] = {}
    for d in depositos:
        loja_norm = str(int(str(d["loja"])))
        dt_corte  = _parse(d["data_corte"])
        data_fmt  = _fmt_date(dt_corte) if dt_corte else d["data_corte"]
        key = (data_fmt, loja_norm)
        dep_idx[key] = dep_idx.get(key, 0.0) + d["total"]

    # ── Grupo crédito: agrupa depósitos pelo dia de crédito esperado ───────
    grupo_idx: dict[tuple, dict] = {}
    for (data_dep_str, loja), total_dep in dep_idx.items():
        dt_dep = _parse(data_dep_str)
        if dt_dep is None:
            continue
        data_cred_str = _fmt_date(_credito_esperado(dt_dep))
        key_cred = (data_cred_str, loja)
        if key_cred not in grupo_idx:
            grupo_idx[key_cred] = {"datas_deposito": [], "total_deposito": 0.0}
        grupo_idx[key_cred]["datas_deposito"].append(data_dep_str)
        grupo_idx[key_cred]["total_deposito"] += total_dep

    linhas = []
    processados_cred: set = set()

    for (data_cred_str, loja), grupo in sorted(grupo_idx.items()):
        total_dep = grupo["total_deposito"]
        datas_dep = sorted(grupo["datas_deposito"], key=lambda s: _parse(s) or datetime.max)

        # ── Grupo 1: Depósito × Memorando (cada dia individualmente) ──────
        # Soma os memorandos de todos os dias do grupo
        total_memo = 0.0
        tem_memo   = False
        for data_dep_str in datas_dep:
            memo_val = mem_idx.get((loja, data_dep_str))
            if memo_val is not None:
                total_memo += float(memo_val or 0)
                tem_memo = True

        # ── Grupo 2: Depósito × Crédito (agrupado pela data de crédito) ───
        key_cred   = (data_cred_str, loja)
        total_cred = cred_idx.get(key_cred, 0.0)
        processados_cred.add(key_cred)

        # Diferenças individuais por grupo
        dif_dep_memo = total_dep - total_memo   # positivo = depósito maior
        dif_dep_cred = total_dep - total_cred   # positivo = depósito maior

        # ── Status: verde só se AMBOS os grupos batem ─────────────────────
        g1_ok = tem_memo and _iguais(total_dep, total_memo)
        g2_ok = total_cred > 0 and _iguais(total_dep, total_cred)

        if total_cred == 0.0 and not tem_memo:
            status = "sem_credito"
        elif total_cred == 0.0:
            status = "sem_credito"
        elif not tem_memo:
            status = "divergente"
        elif g1_ok and g2_ok:
            status = "ok"
            dif_dep_memo = 0.0
            dif_dep_cred = 0.0
        else:
            status = "divergente"

        # Gera uma linha por data de depósito do grupo
        for idx_d, data_dep_str in enumerate(datas_dep):
            memo_val = mem_idx.get((loja, data_dep_str))
            linhas.append({
                "data_deposito":         data_dep_str,
                "data_credito_esperada": data_cred_str,
                "loja":                  loja,
                "total_deposito":        dep_idx[(data_dep_str, loja)],
                "total_deposito_grupo":  total_dep,
                "total_credito":         total_cred if idx_d == 0 else 0.0,
                "memo_remessa":          memo_val,
                "memo_remessa_grupo":    total_memo,
                # diferenças separadas por grupo
                "dif_dep_memo":          dif_dep_memo if idx_d == 0 else 0.0,
                "dif_dep_cred":          dif_dep_cred if idx_d == 0 else 0.0,
                "status":                status,
                "grupo_size":            len(datas_dep),
                "observacao":            "",
            })

    # créditos sem depósito correspondente
    for (data_cred_str, loja), total_cred in cred_idx.items():
        if (data_cred_str, loja) not in processados_cred:
            linhas.append({
                "data_deposito":         "—",
                "data_credito_esperada": data_cred_str,
                "loja":                  loja,
                "total_deposito":        0.0,
                "total_deposito_grupo":  0.0,
                "total_credito":         total_cred,
                "memo_remessa":          None,
                "memo_remessa_grupo":    0.0,
                "dif_dep_memo":          0.0,
                "dif_dep_cred":          total_cred,
                "status":                "sem_deposito",
                "grupo_size":            1,
                "observacao":            "",
            })

    # ── Aplica ajustes manuais salvos no SQLite ────────────────────────────
    ajustes = db.get_ajustes("brinks")
    for linha in linhas:
        key = (linha["loja"], linha["data_deposito"])
        if key in ajustes:
            aj = ajustes[key]
            linha["total_deposito"]        = aj["total_deposito"]
            linha["total_credito"]         = aj["total_credito"]
            linha["memo_remessa"]          = aj["memo_remessa"]
            linha["dif_dep_memo"]          = aj["dif_dep_memo"] or 0.0
            linha["dif_dep_cred"]          = aj["dif_dep_cred"] or 0.0
            linha["status"]                = aj["status"]
            linha["observacao"]            = aj["observacao"] or ""
            if aj.get("data_credito_esperada"):
                linha["data_credito_esperada"] = aj["data_credito_esperada"]

    return linhas


# ──────────────────────────────────────────────────────────────────────────── #
#  Constantes visuais                                                           #
# ──────────────────────────────────────────────────────────────────────────── #

STATUS_LABEL = {
    "ok":           "✔  OK",
    "divergente":   "⚠  Divergente",
    "sem_credito":  "✘  Sem crédito",
    "sem_deposito": "?  Sem depósito",
}
STATUS_COLOR = {
    "ok":           theme.ACCENT_GREEN,
    "divergente":   theme.ACCENT_ORANGE,
    "sem_credito":  theme.ACCENT_RED,
    "sem_deposito": "#9b9b9b",
}

CHIP_W      = 30
CHIP_H      = 20
CHIP_PAD_X  = 7
CHIP_PAD_Y  = 5
CHIP_VAZIO  = theme.SURFACE2
HEADER_X    = 32
HEADER_Y    = 115


# ──────────────────────────────────────────────────────────────────────────── #
#  Janela de detalhe                                                            #
# ──────────────────────────────────────────────────────────────────────────── #

class DetalheCell(tk.Toplevel):
    def __init__(self, parent, linha: dict):
        super().__init__(parent)
        self.title(f"Loja {linha['loja']}  ·  {linha['data_deposito']}")
        self.geometry("460x540")
        self.resizable(False, False)
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)

        self._dep        = tk.DoubleVar(value=linha.get("total_deposito", 0.0))
        self._cred       = tk.DoubleVar(value=linha.get("total_credito", 0.0))
        self._memo       = tk.DoubleVar(value=linha.get("memo_remessa") or 0.0)
        self._data_cred  = tk.StringVar(value=linha.get("data_credito_esperada", ""))
        self._linha      = linha

        self._build(linha)
        self._center(parent)

    def _fmt(self, v: float) -> str:
        return f"R$ {abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _recalcular(self, *_):
        dep  = self._dep.get()
        cred = self._cred.get()
        memo = self._memo.get()

        def _iguais(a, b, tol=0.01):
            return abs(float(a or 0) - float(b or 0)) <= tol

        g1_ok = memo > 0 and _iguais(dep, memo)
        g2_ok = cred > 0 and _iguais(dep, cred)

        if cred == 0.0:
            status = "sem_credito"
        elif memo == 0.0:
            status = "divergente"
        elif g1_ok and g2_ok:
            status = "ok"
        else:
            status = "divergente"

        dif_memo = 0.0 if (status == "ok") else abs(dep - memo)
        dif_cred = 0.0 if (status == "ok") else abs(dep - cred)

        def _fmt_dif(v):
            return "Zerado" if abs(v) < 0.02 else self._fmt(v)

        self._lbl_dif_memo.configure(text=_fmt_dif(dif_memo))
        self._lbl_dif_cred.configure(text=_fmt_dif(dif_cred))
        self._lbl_status.configure(
            text=STATUS_LABEL[status],
            fg=STATUS_COLOR[status],
        )
        self._cur_status   = status
        self._cur_dif_memo = dif_memo
        self._cur_dif_cred = dif_cred

    def _make_edit_row(self, body, label: str, var: tk.DoubleVar, bg: str):
        row = tk.Frame(body, bg=bg)
        row.pack(fill="x")
        tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=22, pady=7).pack(side="left", padx=10)

        right = tk.Frame(row, bg=bg)
        right.pack(side="right", padx=10)

        lbl = tk.Label(right, text=self._fmt(var.get()),
                       bg=bg, fg=theme.FG_TEXT,
                       font=("Segoe UI Semibold", 10))
        lbl.pack(side="left")

        pen = tk.Label(right, text=" ✏", bg=bg, fg=theme.FG_MUTED,
                       font=("Segoe UI", 9), cursor="hand2")
        pen.pack(side="left")

        def _start_edit(e=None):
            lbl.pack_forget()
            pen.pack_forget()
            entry = tk.Entry(right, textvariable=var, width=12,
                             font=("Segoe UI Semibold", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
            entry.pack(side="left")
            entry.focus_set()
            entry.select_range(0, "end")

            def _finish(e=None):
                try:
                    raw = entry.get().strip().replace("R$", "").strip()
                    if "," in raw:
                        raw = raw.replace(".", "").replace(",", ".")
                    var.set(float(raw))
                except ValueError:
                    pass
                lbl.configure(text=self._fmt(var.get()))
                entry.destroy()
                lbl.pack(side="left")
                pen.pack(side="left")
                self._recalcular()

            entry.bind("<Return>", _finish)
            entry.bind("<FocusOut>", _finish)

        pen.bind("<Button-1>", _start_edit)
        lbl.bind("<Double-1>", _start_edit)

    def _make_date_edit_row(self, body, label: str, var: tk.StringVar, bg: str):
        """Linha editável para data (string DD/MM/YYYY)."""
        row = tk.Frame(body, bg=bg)
        row.pack(fill="x")
        tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=22, pady=7).pack(side="left", padx=10)

        right = tk.Frame(row, bg=bg)
        right.pack(side="right", padx=10)

        lbl = tk.Label(right, text=var.get(),
                       bg=bg, fg=theme.FG_TEXT,
                       font=("Segoe UI Semibold", 10))
        lbl.pack(side="left")

        pen = tk.Label(right, text=" ✏", bg=bg, fg=theme.FG_MUTED,
                       font=("Segoe UI", 9), cursor="hand2")
        pen.pack(side="left")

        def _start_edit(e=None):
            lbl.pack_forget()
            pen.pack_forget()
            entry = tk.Entry(right, textvariable=var, width=12,
                             font=("Segoe UI Semibold", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
            entry.pack(side="left")
            entry.focus_set()
            entry.select_range(0, "end")

            def _finish(e=None):
                # valida formato DD/MM/YYYY
                txt = entry.get().strip()
                try:
                    datetime.strptime(txt, "%d/%m/%Y")
                    var.set(txt)
                except ValueError:
                    pass  # mantém o valor anterior
                lbl.configure(text=var.get())
                entry.destroy()
                lbl.pack(side="left")
                pen.pack(side="left")

            entry.bind("<Return>", _finish)
            entry.bind("<FocusOut>", _finish)

        pen.bind("<Button-1>", _start_edit)
        lbl.bind("<Double-1>", _start_edit)

    def _build(self, l: dict):
        cor = STATUS_COLOR[l["status"]]

        header = tk.Frame(self, bg=cor)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"  Loja {l['loja']}   ·   {l['data_deposito']}",
            bg=cor, fg="white",
            font=("Segoe UI Semibold", 11), pady=10, anchor="w",
        ).pack(fill="x", padx=12)

        body = tk.Frame(self, bg=theme.CARD_BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        grupo_size      = l.get("grupo_size", 1)
        total_dep_grupo = l.get("total_deposito_grupo", l["total_deposito"])

        def _static_row(label, valor, idx):
            bg = theme.CARD_BG if idx % 2 == 0 else theme.SURFACE2
            row = tk.Frame(body, bg=bg)
            row.pack(fill="x")
            tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                     font=("Segoe UI", 10), anchor="w",
                     width=22, pady=7).pack(side="left", padx=10)
            tk.Label(row, text=valor, bg=bg, fg=theme.FG_TEXT,
                     font=("Segoe UI Semibold", 10),
                     anchor="e").pack(side="right", padx=10)

        _static_row("Data depósito", l["data_deposito"], 0)

        # Data crédito esperada — editável
        bg1 = theme.SURFACE2
        self._make_date_edit_row(body, "Data crédito esperada", self._data_cred, bg1)

        row_idx = 2

        if grupo_size > 1:
            _static_row(f"Depósito ({l['data_deposito']})",
                        self._fmt(l["total_deposito"]), row_idx)
            row_idx += 1
            _static_row(f"Total grupo ({grupo_size} dias)",
                        self._fmt(total_dep_grupo), row_idx)
            row_idx += 1

        # ── Grupo 1: Depósito × Memorando ─────────────────────────────────
        sep1_bg = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        sep1 = tk.Frame(body, bg=sep1_bg)
        sep1.pack(fill="x")
        tk.Label(sep1, text="— Depósito × Memorando —",
                 bg=sep1_bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8, "italic"),
                 pady=3).pack(side="left", padx=10)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        dep_label = "Total depósito" if grupo_size == 1 else f"Depósito ({l['data_deposito']})"
        self._make_edit_row(body, dep_label, self._dep, bg_e)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Memorando (remessa)", self._memo, bg_e)
        row_idx += 1

        bg_d = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_dm = tk.Frame(body, bg=bg_d)
        row_dm.pack(fill="x")
        tk.Label(row_dm, text="Diferença dep×memo", bg=bg_d, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=22, pady=7).pack(side="left", padx=10)
        dif_memo_init = l.get("dif_dep_memo", 0.0) or 0.0
        self._lbl_dif_memo = tk.Label(row_dm,
            text="Zerado" if abs(dif_memo_init) < 0.02 else self._fmt(dif_memo_init),
            bg=bg_d, fg=theme.FG_TEXT,
            font=("Segoe UI Semibold", 10))
        self._lbl_dif_memo.pack(side="right", padx=10)
        row_idx += 1

        # ── Grupo 2: Depósito × Crédito ───────────────────────────────────
        sep2_bg = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        sep2 = tk.Frame(body, bg=sep2_bg)
        sep2.pack(fill="x")
        tk.Label(sep2, text="— Depósito × Crédito —",
                 bg=sep2_bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8, "italic"),
                 pady=3).pack(side="left", padx=10)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Total crédito", self._cred, bg_e)
        row_idx += 1

        bg_d2 = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_dc = tk.Frame(body, bg=bg_d2)
        row_dc.pack(fill="x")
        tk.Label(row_dc, text="Diferença dep×crédito", bg=bg_d2, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=22, pady=7).pack(side="left", padx=10)
        dif_cred_init = l.get("dif_dep_cred", 0.0) or 0.0
        self._lbl_dif_cred = tk.Label(row_dc,
            text="Zerado" if abs(dif_cred_init) < 0.02 else self._fmt(dif_cred_init),
            bg=bg_d2, fg=theme.FG_TEXT,
            font=("Segoe UI Semibold", 10))
        self._lbl_dif_cred.pack(side="right", padx=10)
        row_idx += 1

        # ── Status geral ───────────────────────────────────────────────────
        bg_s = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_st = tk.Frame(body, bg=bg_s)
        row_st.pack(fill="x")
        tk.Label(row_st, text="Status", bg=bg_s, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=22, pady=7).pack(side="left", padx=10)
        self._lbl_status = tk.Label(row_st,
                                     text=STATUS_LABEL[l["status"]],
                                     bg=bg_s, fg=STATUS_COLOR[l["status"]],
                                     font=("Segoe UI Semibold", 10))
        self._lbl_status.pack(side="right", padx=10)
        self._cur_status   = l["status"]
        self._cur_dif_memo = dif_memo_init
        self._cur_dif_cred = dif_cred_init

        # ── Observações ───────────────────────────────────────────────────
        obs_frame = tk.Frame(self, bg=theme.BG_APP)
        obs_frame.pack(fill="x", padx=14, pady=(6, 0))
        tk.Label(obs_frame, text="Observações:", bg=theme.BG_APP,
                 fg=theme.FG_MUTED, font=("Segoe UI", 9)).pack(anchor="w")

        self._obs_var = tk.StringVar(value=l.get("observacao", "") or "")
        obs_entry = tk.Entry(obs_frame, textvariable=self._obs_var,
                             font=("Segoe UI", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
        obs_entry.pack(fill="x", pady=4, ipady=5)

        self._lbl_chars = tk.Label(obs_frame, text="0/150",
                                    bg=theme.BG_APP, fg=theme.FG_MUTED,
                                    font=("Segoe UI", 8))
        self._lbl_chars.pack(anchor="e")

        def _on_obs(*_):
            txt = self._obs_var.get()
            if len(txt) > 150:
                self._obs_var.set(txt[:150])
            self._lbl_chars.configure(text=f"{len(self._obs_var.get())}/150")

        self._obs_var.trace_add("write", _on_obs)
        _on_obs()

        # ── Botão salvar ──────────────────────────────────────────────────
        tk.Button(
            self,
            text="Salvar alterações",
            command=self._salvar,
            bg=theme.PRIMARY, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=8,
            font=("Segoe UI Semibold", 10),
        ).pack(fill="x", padx=14, pady=10)

    def _salvar(self):
        self._linha["total_deposito"]        = self._dep.get()
        self._linha["total_credito"]         = self._cred.get()
        self._linha["memo_remessa"]          = self._memo.get() or None
        self._linha["dif_dep_memo"]          = self._cur_dif_memo
        self._linha["dif_dep_cred"]          = self._cur_dif_cred
        self._linha["status"]                = self._cur_status
        self._linha["observacao"]            = self._obs_var.get().strip()
        self._linha["data_credito_esperada"] = self._data_cred.get()

        db.salvar_ajuste("brinks", self._linha["loja"], self._linha["data_deposito"], {
            "total_deposito":        self._linha["total_deposito"],
            "total_credito":         self._linha["total_credito"],
            "memo_remessa":          self._linha["memo_remessa"],
            "dif_dep_memo":          self._cur_dif_memo,
            "dif_dep_cred":          self._cur_dif_cred,
            "status":                self._cur_status,
            "observacao":            self._linha["observacao"],
            "data_credito_esperada": self._data_cred.get(),
        })
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w}x{h}+{pw - w // 2}+{ph - h // 2}")


# ──────────────────────────────────────────────────────────────────────────── #
#  Frame principal                                                              #
# ──────────────────────────────────────────────────────────────────────────── #

class BrinksPainel(ttk.Frame):

    def __init__(self, master, **kwargs):
        super().__init__(master, style="App.TFrame", **kwargs)
        db.inicializar()
        self._linhas: list[dict] = []
        self._build_shell()
        self._atualizar()

    def _build_shell(self):
        toolbar1 = tk.Frame(self, bg=theme.CARD_BG, pady=6)
        toolbar1.pack(fill="x")

        tk.Button(
            toolbar1, text="  Atualizar",
            command=self._atualizar,
            bg=theme.PRIMARY, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=6,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left", padx=10)

        tk.Label(toolbar1, text="Status:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        self._filtro_var = tk.StringVar(value="todos")
        for label, valor in [
            ("Todos", "todos"), ("✔ OK", "ok"),
            ("⚠ Divergente", "divergente"), ("✘ Sem crédito", "sem_credito"),
        ]:
            tk.Radiobutton(
                toolbar1, text=label,
                variable=self._filtro_var, value=valor,
                command=self._aplicar_filtro,
                bg=theme.CARD_BG, fg=theme.FG_TEXT,
                selectcolor=theme.PRIMARY,
                activebackground=theme.CARD_BG,
                font=("Segoe UI", 9), relief="flat", bd=0,
            ).pack(side="left", padx=4)

        self._lbl_resumo = tk.Label(
            toolbar1, text="",
            bg=theme.CARD_BG, fg=theme.FG_MUTED,
            font=("Segoe UI", 9),
        )
        self._lbl_resumo.pack(side="right", padx=16)

        toolbar2 = tk.Frame(self, bg=theme.CARD_BG, pady=4)
        toolbar2.pack(fill="x")

        tk.Label(toolbar2, text="Regional:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        ufs_disponiveis = sorted(set(LOJA_UF.values()))
        self._uf_var = tk.StringVar(value="Todas")
        self._cmb_uf = ttk.Combobox(
            toolbar2, textvariable=self._uf_var,
            values=["Todas"] + ufs_disponiveis,
            state="readonly", width=8, font=("Segoe UI", 9),
        )
        self._cmb_uf.pack(side="left", padx=4)
        self._cmb_uf.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Label(toolbar2, text="Mês:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(16, 4))

        self._mes_var = tk.StringVar(value="Todos")
        self._cmb_mes = ttk.Combobox(
            toolbar2, textvariable=self._mes_var,
            values=["Todos"] + [f"{n:02d} – {MESES_PT[n]}" for n in range(1, 13)],
            state="readonly", width=16, font=("Segoe UI", 9),
        )
        self._cmb_mes.pack(side="left", padx=4)
        self._cmb_mes.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Label(toolbar2, text="Ano:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        self._ano_var = tk.StringVar(value="Todos")
        self._cmb_ano = ttk.Combobox(
            toolbar2, textvariable=self._ano_var,
            values=["Todos"], state="readonly", width=8, font=("Segoe UI", 9),
        )
        self._cmb_ano.pack(side="left", padx=4)
        self._cmb_ano.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")

        outer = ttk.Frame(self, style="App.TFrame")
        outer.pack(fill="both", expand=True)

        self._vsb = ttk.Scrollbar(outer, orient="vertical", style="Vertical.TScrollbar")
        self._hsb = ttk.Scrollbar(outer, orient="horizontal")
        self._canvas = tk.Canvas(
            outer, bg=theme.BG_APP, highlightthickness=0,
            yscrollcommand=self._vsb.set,
            xscrollcommand=self._hsb.set,
        )
        self._vsb.configure(command=self._canvas.yview)
        self._hsb.configure(command=self._canvas.xview)
        self._vsb.pack(side="right", fill="y")
        self._hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self._canvas.bind("<Shift-MouseWheel>",
            lambda e: self._canvas.xview_scroll(-1 * (e.delta // 120), "units"))

        leg = tk.Frame(self, bg=theme.BG_APP, pady=5)
        leg.pack(fill="x", side="bottom")
        for status, cor in STATUS_COLOR.items():
            tk.Label(leg, text="●", fg=cor, bg=theme.BG_APP,
                     font=("Segoe UI", 11)).pack(side="left", padx=(12, 2))
            tk.Label(leg, text=STATUS_LABEL[status],
                     fg=theme.FG_MUTED, bg=theme.BG_APP,
                     font=("Segoe UI", 9)).pack(side="left", padx=(0, 10))

        tk.Label(leg, text="Clique em um quadrado para ver o detalhe",
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8)).pack(side="right", padx=12)

    def _atualizar(self):
        self._linhas = conciliar()
        self._popular_anos()
        self._aplicar_filtro()

    def _popular_anos(self):
        anos = set()
        for l in self._linhas:
            dt = _parse(l["data_deposito"])
            if dt:
                anos.add(str(dt.year))
        opcoes = ["Todos"] + sorted(anos, reverse=True)
        self._cmb_ano.configure(values=opcoes)
        if self._ano_var.get() not in opcoes:
            self._ano_var.set("Todos")

    def _aplicar_filtro(self):
        f       = self._filtro_var.get()
        uf      = self._uf_var.get()
        mes_sel = self._mes_var.get()
        ano_sel = self._ano_var.get()

        mes_num = None
        if mes_sel != "Todos":
            try:
                mes_num = int(mes_sel.split("–")[0].strip())
            except (ValueError, IndexError):
                mes_num = None

        ano_num = None if ano_sel == "Todos" else int(ano_sel)

        exibir = []
        for l in self._linhas:
            if f != "todos" and l["status"] != f:
                continue
            if uf != "Todas":
                if LOJA_UF.get(str(l["loja"]), "") != uf:
                    continue
            if mes_num is not None or ano_num is not None:
                dt = _parse(l["data_deposito"])
                if dt is None:
                    continue
                if mes_num is not None and dt.month != mes_num:
                    continue
                if ano_num is not None and dt.year != ano_num:
                    continue
            exibir.append(l)

        datas_extras = self._gerar_datas_mes(mes_num, ano_num) if mes_num else None
        self._desenhar_grade(exibir, datas_extras=datas_extras)
        self._atualizar_resumo()

    def _gerar_datas_mes(self, mes: int, ano: int | None) -> list[str]:
        if ano is None:
            ano = datetime.now().year
        _, num_dias = calendar.monthrange(ano, mes)
        return [f"{d:02d}/{mes:02d}/{ano}" for d in range(1, num_dias + 1)]

    def _atualizar_resumo(self):
        total = len(self._linhas)
        ok  = sum(1 for l in self._linhas if l["status"] == "ok")
        div = sum(1 for l in self._linhas if l["status"] == "divergente")
        sc  = sum(1 for l in self._linhas if l["status"] == "sem_credito")
        self._lbl_resumo.configure(
            text=f"{total} registros  ·  ✔ {ok}  ⚠ {div}  ✘ {sc}"
        )

    def _desenhar_grade(self, linhas: list[dict], datas_extras: list[str] | None = None):
        c = self._canvas
        c.delete("all")

        if not linhas and not datas_extras:
            c.create_text(200, 80, text="Nenhum dado para exibir.",
                          fill=theme.FG_MUTED, font=("Segoe UI", 11))
            c.configure(scrollregion=(0, 0, 400, 160))
            return

        def _sort_date(s):
            try:    return datetime.strptime(s, "%d/%m/%Y")
            except: return datetime.max

        datas_dados = {l["data_deposito"] for l in linhas if l["data_deposito"] != "—"}
        if datas_extras:
            datas_todas = sorted(datas_dados | set(datas_extras), key=_sort_date)
        else:
            datas_todas = sorted(datas_dados, key=_sort_date)

        lojas = sorted({l["loja"] for l in linhas},
                       key=lambda x: int(x) if str(x).isdigit() else x)

        if not datas_todas or not lojas:
            c.create_text(200, 80, text="Nenhum dado para exibir.",
                          fill=theme.FG_MUTED, font=("Segoe UI", 11))
            c.configure(scrollregion=(0, 0, 400, 160))
            return

        idx = {(l["data_deposito"], l["loja"]): l for l in linhas}

        ROW_H   = CHIP_H + CHIP_PAD_Y * 2
        COL_W   = CHIP_W + CHIP_PAD_X * 2
        total_w = HEADER_Y + len(datas_todas) * COL_W + 20
        total_h = HEADER_X + len(lojas) * ROW_H + 20

        for j, data in enumerate(datas_todas):
            x = HEADER_Y + j * COL_W + COL_W // 2
            c.create_text(x, HEADER_X // 2, text=data[:5],
                          fill=theme.FG_MUTED, font=("Segoe UI", 8))

        c.create_line(0, HEADER_X, total_w, HEADER_X, fill=theme.GRID_LINE, width=1)

        for i, loja in enumerate(lojas):
            y_top  = HEADER_X + i * ROW_H
            bg_row = theme.CARD_BG if i % 2 == 0 else theme.SURFACE2

            c.create_rectangle(0, y_top, total_w, y_top + ROW_H, fill=bg_row, outline="")

            uf_loja    = LOJA_UF.get(str(loja), "")
            label_loja = f"L{loja}" + (f" ({uf_loja})" if uf_loja else "")
            c.create_text(6, y_top + ROW_H // 2, text=label_loja,
                          fill=theme.FG_MUTED, font=("Segoe UI", 9), anchor="w")

            for j, data in enumerate(datas_todas):
                linha = idx.get((data, loja))
                x0 = HEADER_Y + j * COL_W + CHIP_PAD_X
                y0 = y_top + CHIP_PAD_Y
                x1 = x0 + CHIP_W
                y1 = y0 + CHIP_H

                if linha is None:
                    c.create_rectangle(x0, y0, x1, y1, fill=CHIP_VAZIO, outline=theme.GRID_LINE)
                else:
                    cor   = STATUS_COLOR[linha["status"]]
                    hover = self._darken(cor)
                    rid   = c.create_rectangle(x0, y0, x1, y1, fill=cor, outline=cor, tags=("chip",))

                    def _enter(e, r=rid, h=hover): c.itemconfig(r, fill=h, outline=h)
                    def _leave(e, r=rid, co=cor):  c.itemconfig(r, fill=co, outline=co)
                    def _click(e, l=linha):         DetalheCell(self.winfo_toplevel(), l)

                    c.tag_bind(rid, "<Enter>",    _enter)
                    c.tag_bind(rid, "<Leave>",    _leave)
                    c.tag_bind(rid, "<Button-1>", _click)
                    c.tag_bind(rid, "<ButtonRelease-1>",
                               lambda e, r=rid, co=cor: c.itemconfig(r, fill=co, outline=co))

        c.configure(scrollregion=(0, 0, total_w, total_h))
        c.configure(cursor="hand2")

    @staticmethod
    def _darken(hex_color: str) -> str:
        h = hex_color.lstrip("#")
        r, g, b = (int(h[i:i+2], 16) for i in (0, 2, 4))
        return f"#{max(0,r-35):02x}{max(0,g-35):02x}{max(0,b-35):02x}"
