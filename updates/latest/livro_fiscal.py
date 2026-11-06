"""
livro_fiscal.py
Módulo autônomo com as duas atividades do Livro Fiscal:
  - Livro Fiscal × Relatórios  (iniciar_conciliacao)
  - Livro Fiscal × Contabilidade (iniciar_comparador_diario)

Pode ser usado standalone (if __name__ == "__main__") ou
integrado ao menu_principal.py via _build_content(root).
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import pdfplumber
import re
import os
from datetime import date, datetime

import core.theme

# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

CFOPS_VALIDOS = {
    "5.102", "5.403", "5.405", "5.646", "5.655", "5.656",
    "6.108", "6.403", "6.404", "7.102",
}
CFOPS_EXCLUIDOS = {"5.409"}

# Fontes
FONTE = ("Segoe UI", 10)
MONO  = ("Consolas", 10)

# ── Atalhos de cor (lidos do theme para evitar repetição) ──────────────────
def _C(key):
    """Atalho: retorna cor fixa do theme (constantes)."""
    return {
        "BG":      core.theme.BG_APP,
        "BG2":     core.theme.CARD_BG,
        "FG":      core.theme.FG_TEXT,
        "MUTED":   core.theme.FG_MUTED,
        "PRIMARY": core.theme.PRIMARY,
        "PRIMARY_H": core.theme.PRIMARY_HOVER,
        "GREEN":   core.theme.ACCENT_GREEN,
        "RED":     core.theme.ACCENT_RED,
        "ORANGE":  core.theme.ACCENT_ORANGE,
        "GRID":    core.theme.GRID_LINE,
    }[key]

# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS DE PARSING  (sem alterações)
# ══════════════════════════════════════════════════════════════════════════════

def _eh_moeda(s):
    if not s:
        return False
    s = str(s).strip()
    return bool(
        re.fullmatch(r"[0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2}", s)
        or re.fullmatch(r"[0-9]+[.,][0-9]{2}", s)
    )

def parse_valor(valor_str):
    if not valor_str:
        return 0.0
    v = str(valor_str).replace(".", "").replace(",", ".")
    try:
        return float(v)
    except Exception:
        return 0.0

def _fmt(v):
    try:
        return (f"R$ {float(v):,.2f}"
                .replace(",", "X").replace(".", ",").replace("X", "."))
    except Exception:
        return "R$ 0,00"

def extrair_data(tokens):
    for t in tokens:
        if re.match(r"\d{2}[/\-]\d{2}[/\-]\d{4}", t) or re.match(r"\d{4}[/\-]\d{2}[/\-]\d{2}", t):
            return t
    return None

def extrair_cfop(tokens):
    """Extrai o CFOP real ignorando colunas de NÚMERO, DIA, SÉRIE."""
    if "ICMS" in tokens:
        idx = tokens.index("ICMS")
        for t in tokens[idx:idx + 6]:
            if t in CFOPS_VALIDOS or t in CFOPS_EXCLUIDOS:
                return t
    for t in tokens[-6:]:
        if t in CFOPS_VALIDOS or t in CFOPS_EXCLUIDOS:
            return t
    return None

def extrair_valor_linha_livro(tokens):
    """Retorna o valor contábil imediatamente antes do CFOP."""
    cfop_tok = extrair_cfop(tokens)
    if cfop_tok is None:
        return None

    cfop_idx = None
    if "ICMS" in tokens:
        icms_pos = len(tokens) - 1 - tokens[::-1].index("ICMS")
        start = max(0, icms_pos - 8)
        for i in range(icms_pos, start - 1, -1):
            if tokens[i] == cfop_tok:
                cfop_idx = i
                break

    if cfop_idx is None:
        for i in range(len(tokens) - 1, -1, -1):
            if tokens[i] == cfop_tok:
                cfop_idx = i
                break

    if cfop_idx is None:
        return None

    for j in range(cfop_idx - 1, -1, -1):
        if _eh_moeda(tokens[j]):
            return tokens[j]
    return None

def canonicalizar(serie, numero):
    def clean(x):
        x = re.sub(r"[^\d-]", "", str(x))
        x = x.split("-")[0]
        return x
    try:
        s = str(int(clean(serie)))
        n = str(int(clean(numero)))
        return (s, n)
    except Exception:
        return None

# ══════════════════════════════════════════════════════════════════════════════
#  EXTRATORES DE PDF  (sem alterações)
# ══════════════════════════════════════════════════════════════════════════════

def extrair_livro_como_lista(pdf_path):
    registros = []
    if not os.path.exists(pdf_path):
        return registros
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.splitlines():
                tokens = linha.split()
                cfop = extrair_cfop(tokens)
                if cfop is None or cfop in CFOPS_EXCLUIDOS:
                    continue
                tokens_upper = [t.upper() for t in tokens]
                if any(t in tokens_upper for t in ["NFCE", "NFE", "NFS", "MFE"]):
                    for i, t in enumerate(tokens_upper):
                        if t in ["NFCE", "NFE", "NFS", "MFE"] and i + 2 < len(tokens):
                            serie, numero = tokens[i + 1], tokens[i + 2]
                            valor = extrair_valor_linha_livro(tokens)
                            if not valor:
                                break
                            registros.append({
                                "serie":       serie,
                                "numero":      numero,
                                "fiscal":      t,
                                "valor_float": parse_valor(valor),
                                "data":        extrair_data(tokens),
                            })
                            break
    return registros


def extrair_cupons(pdf_path):
    documentos = {}
    if not os.path.exists(pdf_path):
        return documentos
    CANCEL_KEYS = ("CANCEL", "CANCELADO", "CANCELADA", "CANC.", "CANC")
    TOTAL_KEYS  = ("TOTAL", "TOTAIS")
    VENDA_KEYS  = ("VENDA PIX", "VENDA A CARTAO", "VENDA DINHEIRO", "DESCRICAO TOTAIS")

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.splitlines():
                tokens = linha.split()
                if any(t in CFOPS_EXCLUIDOS for t in tokens):
                    continue
                if not any(t in CFOPS_VALIDOS for t in tokens):
                    continue
                up = linha.upper()
                if any(key in up for key in CANCEL_KEYS):
                    continue
                if any(key in up for key in TOTAL_KEYS):
                    continue
                if any(key in up for key in VENDA_KEYS):
                    continue
                if "NORMAL" not in up:
                    continue
                numeros = re.findall(r"\d+", linha)
                if len(numeros) < 2:
                    continue
                serie  = numeros[0]
                numero = numeros[1]
                par = canonicalizar(serie, numero)
                if not par:
                    continue
                valores = [t for t in tokens if _eh_moeda(t)]
                if not valores:
                    continue
                documentos[par] = {"tipo": "Cupom", "valor": valores[-1]}
    return documentos


def extrair_notas(pdf_path):
    documentos = {}
    if not os.path.exists(pdf_path):
        return documentos

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            for linha in texto.splitlines():
                up = linha.upper().strip()
                if not up:
                    continue
                if up.startswith("VENDA") or up.startswith("CODPOS") or up.startswith("TOTAL"):
                    continue
                tokens_raw = linha.split()
                if any(t in CFOPS_EXCLUIDOS for t in tokens_raw):
                    continue
                if not any(t in CFOPS_VALIDOS for t in tokens_raw):
                    continue

                tokens = [t.strip().strip(".,;:/\\()[]") for t in tokens_raw]
                tokens_upper = [t.upper() for t in tokens]
                if len(tokens) < 2:
                    continue

                tokens_norm = []
                for t in tokens:
                    m = re.match(r"^([A-Za-z\-]+[A-Za-z0-9]*)(\d{1,3}(?:\.\d{3})*,\d{2})$", t)
                    if m:
                        tokens_norm.append(m.group(1))
                        tokens_norm.append(m.group(2))
                    else:
                        tokens_norm.append(t)
                tokens = tokens_norm

                recombined = []
                i = 0
                while i < len(tokens):
                    t = tokens[i]
                    if i + 1 < len(tokens):
                        nxt = tokens[i + 1]
                        if re.fullmatch(r"\d{1,3},\d{2}", nxt) and re.fullmatch(r"\d{1,3}", t):
                            merged = (t + nxt).replace(".", "")
                            recombined.append(merged)
                            i += 2
                            continue
                    recombined.append(t)
                    i += 1
                tokens = recombined
                tokens_upper = [t.upper() for t in tokens]

                cond = None
                for t in tokens_upper:
                    if re.fullmatch(r'GOOD[A-Z]*', t):
                        cond = "GOOD"
                        break

                nota  = tokens[0] if len(tokens) > 0 else None
                serie = tokens[1] if len(tokens) > 1 else None
                if not nota or not serie or not re.fullmatch(r"\d{6}", str(nota)) or not re.fullmatch(r"\d{3}", str(serie)):
                    nums = [t for t in tokens if re.fullmatch(r"\d+", t)]
                    if len(nums) >= 2:
                        nota  = nums[0]
                        serie = nums[1]
                    else:
                        continue

                valor_liquido = None

                def eh_moeda(s):
                    return (re.fullmatch(r"[0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2}", s)
                            or re.fullmatch(r"[0-9]+[.,][0-9]{2}", s))

                forma_idx = None
                for i, t in enumerate(tokens):
                    if re.fullmatch(r"[A-Z]{2,}\-?[A-Z0-9]*", t):
                        forma_idx = i
                        break
                if forma_idx is not None:
                    for j in range(forma_idx + 1, len(tokens)):
                        if eh_moeda(tokens[j]) and parse_valor(tokens[j]) > 0:
                            valor_liquido = tokens[j]
                            break

                if not valor_liquido:
                    cfop_idx = None
                    for i, t in enumerate(tokens):
                        if t in CFOPS_VALIDOS:
                            cfop_idx = i
                            break
                    if cfop_idx is not None:
                        valores_antes = [
                            (parse_valor(tokens[j]), tokens[j])
                            for j in range(cfop_idx - 1, -1, -1)
                            if eh_moeda(tokens[j]) and parse_valor(tokens[j]) > 0
                        ]
                        if valores_antes:
                            valores_antes.sort(reverse=True)
                            valor_liquido = valores_antes[0][1]

                if not valor_liquido:
                    for j in range(len(tokens) - 1, -1, -1):
                        if eh_moeda(tokens[j]) and parse_valor(tokens[j]) > 0:
                            valor_liquido = tokens[j]
                            break

                if not valor_liquido:
                    continue

                par = canonicalizar(serie, nota)
                if not par:
                    continue

                documentos[par] = {
                    "tipo":  "Nota",
                    "valor": valor_liquido,
                    "data":  extrair_data(tokens),
                    "cond":  cond,
                }

    return documentos


def extrair_prazo(pdf_path):
    documentos = {}
    if not os.path.exists(pdf_path):
        return documentos

    pad = re.compile(
        r'(?:^|\s)'
        r'(\d{1,2})\s+'
        r'(\d{4,6})\s+'
        r'(\d{1,3})\s+'
        r'.*?'
        r'([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})'
        r'(?:\s+[Ss])?\b'
    )

    def _canon(serie, nota):
        return (str(serie).strip(), str(nota).strip())

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            for linha in texto.splitlines():
                u = (linha or "").upper().strip()
                eh_total_resumo = (
                    u.startswith("TOTAL") or
                    "TOTAL FILIAL" in u or
                    "TOTAL GERAL" in u or
                    u.startswith("CLIENTE TOTAL")
                )
                if eh_total_resumo:
                    continue
                m = pad.search(linha)
                if not m:
                    continue
                filial, nota, serie, valor = m.groups()
                par = _canon(serie, nota)
                documentos[par] = {"tipo": "Prazo", "valor": valor, "data": None}

    return documentos


# ══════════════════════════════════════════════════════════════════════════════
#  WIDGETS AUXILIARES  (modernizados com theme.py)
# ══════════════════════════════════════════════════════════════════════════════

def _btn(parent, text, cmd, danger=False, secondary=False):
    """Botão moderno usando as cores do theme."""
    if danger:
        bg   = _C("RED")
        abg  = "#c0392b"
    elif secondary:
        bg   = _C("BG2")
        abg  = _C("GRID")
    else:
        bg   = _C("PRIMARY")
        abg  = _C("PRIMARY_H")

    b = tk.Button(
        parent, text=text, command=cmd,
        font=("Segoe UI Semibold", 10),
        bg=bg, fg=_C("FG"),
        activebackground=abg, activeforeground=_C("FG"),
        relief="flat", bd=0,
        padx=14, pady=7,
        cursor="hand2",
    )
    # Efeito hover
    b.bind("<Enter>", lambda e: b.config(bg=abg))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


def _entry(parent, width=55):
    """Entry estilizado no tema escuro."""
    e = tk.Entry(
        parent, width=width,
        font=FONTE,
        bg=_C("BG2"), fg=_C("FG"),
        insertbackground=_C("FG"),
        relief="flat",
        bd=0,
        highlightthickness=1,
        highlightbackground=_C("GRID"),
        highlightcolor=_C("PRIMARY"),
    )
    return e


def _label(parent, text, muted=False, title=False):
    """Label estilizado no tema."""
    font = ("Segoe UI", 14, "bold") if title else FONTE
    fg   = _C("MUTED") if muted else _C("FG")
    return tk.Label(parent, text=text, font=font, bg=_C("BG"), fg=fg)


def _section_header(parent, text):
    """Cabeçalho de seção com linha decorativa."""
    f = tk.Frame(parent, bg=_C("BG"))
    tk.Label(f, text=text, font=("Segoe UI Semibold", 11),
             bg=_C("BG"), fg=_C("PRIMARY")).pack(side="left")
    tk.Frame(f, bg=_C("GRID"), height=1).pack(side="left", fill="x", expand=True, padx=(8, 0))
    return f


def _file_row(parent, label_text, row, picker_fn):
    """
    Linha de seleção de arquivo: Label | Entry | Botão.
    Retorna o widget Entry.
    """
    pad_lbl = tk.Label(
        parent, text=label_text, font=("Segoe UI", 10),
        bg=_C("BG"), fg=_C("MUTED"), anchor="w", width=16,
    )
    pad_lbl.grid(row=row, column=0, sticky="w", pady=4, padx=(0, 6))

    e = _entry(parent, width=52)
    e.grid(row=row, column=1, padx=(0, 6), sticky="ew")

    btn = _btn(parent, "📂 Selecionar", lambda ent=e: picker_fn(ent), secondary=True)
    btn.grid(row=row, column=2)
    return e


def _selecionar_pdf(entry):
    p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if p:
        entry.delete(0, tk.END)
        entry.insert(0, p)

def _selecionar_excel(entry):
    p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if p:
        entry.delete(0, tk.END)
        entry.insert(0, p)


def _saida_text(parent, height=28):
    """Área de texto de saída estilizada."""
    frame = tk.Frame(parent, bg=_C("GRID"), bd=0)
    frame.pack(padx=12, pady=(0, 8), fill="both", expand=True)

    inner = tk.Frame(frame, bg=_C("BG2"), bd=0)
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    txt = scrolledtext.ScrolledText(
        inner, height=height,
        bg=_C("BG2"), fg=_C("FG"),
        font=MONO,
        insertbackground=_C("FG"),
        relief="flat", bd=0,
        padx=10, pady=10,
        selectbackground=_C("PRIMARY"),
        selectforeground=_C("FG"),
    )
    txt.pack(fill="both", expand=True)

    # Tags de cor para o texto de saída
    txt.tag_configure("header",    foreground=_C("PRIMARY"),  font=("Segoe UI Semibold", 11))
    txt.tag_configure("ok",        foreground=_C("GREEN"))
    txt.tag_configure("warn",      foreground=_C("ORANGE"))
    txt.tag_configure("err",       foreground=_C("RED"))
    txt.tag_configure("muted",     foreground=_C("MUTED"))
    txt.tag_configure("separator", foreground=_C("GRID"))

    return txt


# ══════════════════════════════════════════════════════════════════════════════
#  TELA 1 — Livro Fiscal × Relatórios
# ══════════════════════════════════════════════════════════════════════════════

def _build_relatorios(parent, on_voltar=None):
    """Monta a UI de Livro Fiscal × Relatórios dentro de `parent`."""

    for w in parent.winfo_children():
        w.destroy()

    parent.configure(bg=_C("BG"))

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    header = tk.Frame(parent, bg=_C("PRIMARY"), pady=10)
    header.pack(fill="x")
    tk.Label(header, text="📄  Livro Fiscal × Relatórios",
             font=("Segoe UI Semibold", 13), bg=_C("PRIMARY"), fg="white").pack(padx=16)

    # ── Card de entradas ─────────────────────────────────────────────────────
    card = tk.Frame(parent, bg=_C("BG2"), bd=0,
                    highlightthickness=1, highlightbackground=_C("GRID"))
    card.pack(padx=12, pady=10, fill="x")

    inner = tk.Frame(card, bg=_C("BG2"))
    inner.pack(padx=14, pady=12, fill="x")
    inner.columnconfigure(1, weight=1)

    _section_header(inner, "Arquivos de Entrada").grid(
        row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))

    labels_fns = [
        ("Livro Fiscal",    _selecionar_pdf),
        ("Cupons",          _selecionar_pdf),
        ("Notas à Vista",   _selecionar_pdf),
        ("Notas a Prazo",   _selecionar_pdf),
    ]
    entries = []
    for i, (lbl, fn) in enumerate(labels_fns, start=1):
        e = _file_row(inner, lbl, i, fn)
        entries.append(e)

    entry_livro, entry_cupons, entry_notas, entry_prazo = entries

    # ── Área de saída ────────────────────────────────────────────────────────
    saida = _saida_text(parent, height=26)

    # ── Lógica de execução ──
    def executar():
        try:
            livro  = entry_livro.get().strip()
            cupons = entry_cupons.get().strip()
            notas  = entry_notas.get().strip()
            prazo  = entry_prazo.get().strip()
            if not all([livro, cupons, notas, prazo]):
                messagebox.showwarning("Aviso", "Selecione todos os arquivos antes de executar.")
                return

            registros_livro = extrair_livro_como_lista(livro)
            docs_cupons = extrair_cupons(cupons)
            docs_notas  = extrair_notas(notas)
            docs_prazo  = extrair_prazo(prazo)

            docs_outros = {**docs_cupons, **docs_notas, **docs_prazo}
            docs_outros = {k: v for k, v in docs_outros.items() if isinstance(k, tuple) and len(k) >= 2}

            def normalizar_numero(n):
                n = str(n).replace(".", "").replace(",", "").strip()
                return n.lstrip("0") or "0"

            def agrupar_lista(registros):
                agrupado = {}
                mapa_series = {}
                for r in registros:
                    numero = normalizar_numero(r.get("numero"))
                    serie  = str(r.get("serie")).strip()
                    valor  = r.get("valor_float", 0.0)
                    agrupado[numero] = agrupado.get(numero, 0.0) + valor
                    if numero not in mapa_series:
                        mapa_series[numero] = serie
                return agrupado, mapa_series

            def agrupar_dict(docs):
                agrupado = {}
                for key, meta in docs.items():
                    if isinstance(key, tuple) and len(key) >= 2:
                        numero = normalizar_numero(key[1])
                    else:
                        continue
                    agrupado[numero] = agrupado.get(numero, 0.0) + parse_valor(meta.get("valor"))
                return agrupado

            livro_agregado, mapa_series = agrupar_lista(registros_livro)
            outros_agregado = agrupar_dict(docs_outros)

            set_livro  = set(livro_agregado.keys())
            set_outros = set(outros_agregado.keys())
            so_no_livro = set_livro - set_outros
            em_ambos    = set_livro & set_outros

            saida.delete("1.0", tk.END)

            # Resumo
            saida.insert(tk.END, "RESUMO DA CONCILIAÇÃO\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            saida.insert(tk.END, f"  Somente no Livro : {len(so_no_livro)}\n")
            saida.insert(tk.END, f"  Em ambos         : {len(em_ambos)}\n\n")

            # Divergências
            saida.insert(tk.END, "DIVERGÊNCIAS (apenas no Livro)\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            for numero in sorted(so_no_livro)[:50]:
                valor = livro_agregado.get(numero, 0.0)
                serie = mapa_series.get(numero, "-")
                saida.insert(tk.END, f"  Série {serie}  Nº {numero}\n", "warn")
                saida.insert(tk.END, f"  Valor: {_fmt(valor)}\n\n", "muted")

            # Diferenças de valor
            saida.insert(tk.END, "DIFERENÇAS DE VALOR\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            total_difs = 0.0
            for numero in sorted(em_ambos):
                vl = livro_agregado[numero]
                vo = outros_agregado[numero]
                if abs(vl - vo) > 0.01:
                    diff = vl - vo
                    total_difs += abs(diff)
                    serie = mapa_series.get(numero, "-")
                    saida.insert(tk.END, f"  Série {serie}  Nº {numero}\n", "warn")
                    saida.insert(tk.END, f"  Livro     : {_fmt(vl)}\n")
                    saida.insert(tk.END, f"  Relatório : {_fmt(vo)}\n")
                    saida.insert(tk.END, f"  Diferença : {_fmt(diff)}\n\n", "err")
            saida.insert(tk.END, f"  Total das diferenças: {_fmt(total_difs)}\n\n", "err" if total_difs > 0 else "ok")

            # Comparativo por tipo
            total_nfce_livro = sum(r["valor_float"] for r in registros_livro if r["fiscal"] == "NFCE")
            total_nfs_livro  = sum(r["valor_float"] for r in registros_livro if r["fiscal"] == "NFS")
            cupons_agrup = agrupar_dict(docs_cupons)
            total_cupons = sum(cupons_agrup.values())
            total_vista  = sum(parse_valor(m.get("valor")) for m in docs_notas.values() if m.get("valor"))
            total_prazo_ = sum(parse_valor(m.get("valor")) for m in docs_prazo.values() if m.get("valor"))
            total_notas  = total_vista + total_prazo_

            saida.insert(tk.END, "COMPARATIVO POR TIPO\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            saida.insert(tk.END, "  NFCE\n", "muted")
            saida.insert(tk.END, f"  Livro     : {_fmt(total_nfce_livro)}\n")
            saida.insert(tk.END, f"  Relatório : {_fmt(total_cupons)}\n")
            dif_nfce = total_nfce_livro - total_cupons
            saida.insert(tk.END, f"  Diferença : {_fmt(dif_nfce)}\n\n",
                         "ok" if abs(dif_nfce) < 0.01 else "err")

            saida.insert(tk.END, "  NFS\n", "muted")
            saida.insert(tk.END, f"  Livro     : {_fmt(total_nfs_livro)}\n")
            saida.insert(tk.END, f"  Relatório : {_fmt(total_notas)}\n")
            dif_nfs = total_nfs_livro - total_notas
            saida.insert(tk.END, f"  Diferença : {_fmt(dif_nfs)}\n\n",
                         "ok" if abs(dif_nfs) < 0.01 else "err")

            # Comparativo geral
            total_livro_g  = total_nfce_livro + total_nfs_livro
            total_outros_g = total_cupons + total_notas

            saida.insert(tk.END, "COMPARATIVO GERAL\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            saida.insert(tk.END, f"  Total Livro      : {_fmt(total_livro_g)}\n")
            saida.insert(tk.END, f"  Total Relatórios : {_fmt(total_outros_g)}\n")
            dif_g = total_livro_g - total_outros_g
            saida.insert(tk.END, f"  Diferença        : {_fmt(dif_g)}\n\n",
                         "ok" if abs(dif_g) < 0.01 else "err")

            # Notas GoodCard
            notas_good = []
            for key, meta in docs_notas.items():
                c = meta.get("cond")
                if isinstance(key, tuple) and c is not None and c.startswith("GOOD"):
                    serie, numero = key[:2]
                    valor = parse_valor(meta.get("valor")) if meta.get("valor") else 0.0
                    notas_good.append((str(serie), str(numero), valor))

            saida.insert(tk.END, "NOTAS GOODCARD (à vista)\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            if notas_good:
                def _norm(n):
                    return re.sub(r"[^0-9]", "", str(n)).lstrip("0") or "0"
                for serie, numero, valor in sorted(notas_good, key=lambda x: (x[0], _norm(x[1]))):
                    saida.insert(tk.END,
                        f"  Série {serie}  Nº {numero}  {_fmt(valor)}\n", "warn")
            else:
                saida.insert(tk.END, "  Nenhuma nota com COND. = GOOD encontrada.\n", "ok")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # ── Barra de botões ──────────────────────────────────────────────────────
    btn_bar = tk.Frame(parent, bg=_C("BG"), pady=6)
    btn_bar.pack(fill="x", padx=12)

    _btn(btn_bar, "▶  Executar", executar).pack(side="left", padx=(0, 6))
    _btn(btn_bar, "🗑  Limpar",
         lambda: saida.delete("1.0", tk.END), danger=True).pack(side="left")
    if on_voltar:
        _btn(btn_bar, "← Voltar", on_voltar, secondary=True).pack(side="right")


# ══════════════════════════════════════════════════════════════════════════════
#  TELA 2 — Livro Fiscal × Contabilidade
# ══════════════════════════════════════════════════════════════════════════════

def _build_contabilidade(parent, on_voltar=None):
    """Monta a UI de Livro Fiscal × Contabilidade dentro de `parent`."""

    for w in parent.winfo_children():
        w.destroy()

    parent.configure(bg=_C("BG"))

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    header = tk.Frame(parent, bg=_C("PRIMARY"), pady=10)
    header.pack(fill="x")
    tk.Label(header, text="📊  Livro Fiscal × Contabilidade",
             font=("Segoe UI Semibold", 13), bg=_C("PRIMARY"), fg="white").pack(padx=16)

    # ── Card de entradas ─────────────────────────────────────────────────────
    card = tk.Frame(parent, bg=_C("BG2"), bd=0,
                    highlightthickness=1, highlightbackground=_C("GRID"))
    card.pack(padx=12, pady=10, fill="x")

    inner = tk.Frame(card, bg=_C("BG2"))
    inner.pack(padx=14, pady=12, fill="x")
    inner.columnconfigure(1, weight=1)

    _section_header(inner, "Arquivos de Entrada").grid(
        row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))

    entry_pdf   = _file_row(inner, "Livro Fiscal",   1, _selecionar_pdf)
    entry_excel = _file_row(inner, "Contabilidade",  2, _selecionar_excel)

    # ── Área de saída ────────────────────────────────────────────────────────
    saida = _saida_text(parent, height=26)

    # ── Lógica de execução ──
    def executar():
        def _nf_por_1575(doc_str):
            try:
                m = re.search(r'(\d+)\s*-\s*(\d{3})\s*-?', str(doc_str))
                if not m:
                    return None
                return "NFCE" if int(m.group(2)) < 25 else "NFS"
            except Exception:
                return None

        try:
            caminho_pdf   = entry_pdf.get().strip()
            caminho_excel = entry_excel.get().strip()
            if not caminho_pdf or not caminho_excel:
                messagebox.showwarning("Aviso", "Selecione os dois arquivos antes de executar.")
                return

            # ---- Livro Fiscal (PDF) ----
            registros_livro = extrair_livro_como_lista(caminho_pdf)
            livro_por_dia = {}
            for r in registros_livro:
                data_raw = r.get("data", "")
                try:
                    data_fmt = pd.to_datetime(data_raw, dayfirst=True).strftime("%d/%m")
                except Exception:
                    data_fmt = None
                if data_fmt:
                    tipo = "NFCE" if r["fiscal"] == "NFCE" else "NFS"
                    livro_por_dia.setdefault(data_fmt, {"NFCE": 0.0, "NFS": 0.0})
                    livro_por_dia[data_fmt][tipo] += r["valor_float"]

            # ---- Contabilidade (Excel) ----
            df = pd.read_excel(caminho_excel, header=None)
            contabilidade_por_dia = {}
            for _, row in df.iterrows():
                valor = pd.to_numeric(row[9], errors="coerce")
                if pd.isna(valor):
                    continue
                linha_texto = " ".join(str(x) for x in row.values).upper()
                if any(cfop in linha_texto for cfop in CFOPS_EXCLUIDOS):
                    continue

                doc_raw     = row[4]
                texto_valor = str(row[7]).upper()

                match = pd.Series(texto_valor).str.extract(r'(\d{3,4})')
                conta_num = int(match.iloc[0, 0]) if not match.isna().iloc[0, 0] else None

                if conta_num == 1575:
                    tipo_doc = _nf_por_1575(doc_raw)
                    if tipo_doc is None:
                        try:
                            tipo_doc = "NFCE" if 1 <= int(doc_raw) <= 100 else "NFS"
                        except Exception:
                            tipo_doc = "NFS"
                elif "COLIGADAS" in texto_valor or conta_num in (1022, 430):
                    tipo_doc = "NFS"
                elif conta_num == 989 or "DEPÓSITO" in texto_valor:
                    tipo_doc = "NFCE"
                else:
                    if conta_num is not None and conta_num in (1253, 1179):
                        tipo_doc = "NFS"
                    else:
                        try:
                            tipo_doc = "NFCE" if 1 <= int(doc_raw) <= 100 else "NFS"
                        except Exception:
                            tipo_doc = "NFS"

                data_raw = str(row[0]).strip()
                if isinstance(row[0], (pd.Timestamp, datetime)):
                    data_fmt = row[0].strftime("%d/%m")
                else:
                    m = re.search(r'(\d{2}[/-]\d{2})(?:[/-](\d{4}))?', data_raw)
                    if m:
                        dia_mes = m.group(1).replace("-", "/")
                        ano = m.group(2) if m.group(2) else str(date.today().year)
                        try:
                            data_fmt = pd.to_datetime(f"{dia_mes}/{ano}", dayfirst=True).strftime("%d/%m")
                        except Exception:
                            data_fmt = None
                    else:
                        data_fmt = None

                if data_fmt:
                    contabilidade_por_dia.setdefault(data_fmt, {"NFCE": 0.0, "NFS": 0.0})
                    contabilidade_por_dia[data_fmt][tipo_doc] += valor

            # ---- Impressão ----
            todos_os_dias = sorted(
                set(livro_por_dia) | set(contabilidade_por_dia),
                key=lambda d: datetime.strptime(d, "%d/%m"),
            )

            saida.delete("1.0", tk.END)

            saida.insert(tk.END, "LIVRO × CONTABILIDADE — DIÁRIO\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")

            for dia in todos_os_dias:
                livro_nfce = livro_por_dia.get(dia, {}).get("NFCE", 0.0)
                livro_nfs  = livro_por_dia.get(dia, {}).get("NFS",  0.0)
                cont_nfce  = contabilidade_por_dia.get(dia, {}).get("NFCE", 0.0)
                cont_nfs   = contabilidade_por_dia.get(dia, {}).get("NFS",  0.0)

                saida.insert(tk.END, f"\n  📅 Dia {dia}\n", "header")
                saida.insert(tk.END, f"  Livro NFCE   : {_fmt(livro_nfce)}\n")
                saida.insert(tk.END, f"  Contab. NFCE : {_fmt(cont_nfce)}\n",
                             "ok" if abs(livro_nfce - cont_nfce) < 0.01 else "err")
                saida.insert(tk.END, f"  Livro NFS    : {_fmt(livro_nfs)}\n")
                saida.insert(tk.END, f"  Contab. NFS  : {_fmt(cont_nfs)}\n",
                             "ok" if abs(livro_nfs - cont_nfs) < 0.01 else "err")

            # Somatória
            tot_l_nfce = sum(livro_por_dia.get(d, {}).get("NFCE", 0.0) for d in todos_os_dias)
            tot_l_nfs  = sum(livro_por_dia.get(d, {}).get("NFS",  0.0) for d in todos_os_dias)
            tot_c_nfce = sum(contabilidade_por_dia.get(d, {}).get("NFCE", 0.0) for d in todos_os_dias)
            tot_c_nfs  = sum(contabilidade_por_dia.get(d, {}).get("NFS",  0.0) for d in todos_os_dias)

            saida.insert(tk.END, "\n")
            saida.insert(tk.END, "SOMATÓRIA DO PERÍODO\n", "header")
            saida.insert(tk.END, "─" * 44 + "\n", "separator")
            saida.insert(tk.END, f"  Livro   NFCE total : {_fmt(tot_l_nfce)}\n")
            saida.insert(tk.END, f"  Contab. NFCE total : {_fmt(tot_c_nfce)}\n")
            dif_nfce = tot_l_nfce - tot_c_nfce
            saida.insert(tk.END, f"  Dif. NFCE          : {_fmt(dif_nfce)}\n\n",
                         "ok" if abs(dif_nfce) < 0.01 else "err")

            saida.insert(tk.END, f"  Livro   NFS  total : {_fmt(tot_l_nfs)}\n")
            saida.insert(tk.END, f"  Contab. NFS  total : {_fmt(tot_c_nfs)}\n")
            dif_nfs = tot_l_nfs - tot_c_nfs
            saida.insert(tk.END, f"  Dif. NFS           : {_fmt(dif_nfs)}\n\n",
                         "ok" if abs(dif_nfs) < 0.01 else "err")

            saida.insert(tk.END, f"  Livro   TOTAL      : {_fmt(tot_l_nfce + tot_l_nfs)}\n")
            saida.insert(tk.END, f"  Contab. TOTAL      : {_fmt(tot_c_nfce + tot_c_nfs)}\n")
            saida.insert(tk.END, "─" * 44 + "\n\n", "separator")

            # Divergências
            divergencias = []
            for dia in todos_os_dias:
                livro_nfce = livro_por_dia.get(dia, {}).get("NFCE", 0.0)
                livro_nfs  = livro_por_dia.get(dia, {}).get("NFS",  0.0)
                cont_nfce  = contabilidade_por_dia.get(dia, {}).get("NFCE", 0.0)
                cont_nfs   = contabilidade_por_dia.get(dia, {}).get("NFS",  0.0)
                if abs(livro_nfce - cont_nfce) > 0.01:
                    divergencias.append((dia, "NFCE", livro_nfce, cont_nfce))
                if abs(livro_nfs - cont_nfs) > 0.01:
                    divergencias.append((dia, "NFS",  livro_nfs,  cont_nfs))

            if divergencias:
                saida.insert(tk.END, "DIVERGÊNCIAS ENCONTRADAS\n", "header")
                saida.insert(tk.END, "─" * 44 + "\n", "separator")
                for dia, tipo, vl, vc in divergencias:
                    saida.insert(tk.END,
                        f"  Dia {dia}  {tipo}\n"
                        f"  Livro: {_fmt(vl)}   Contab: {_fmt(vc)}   Dif: {_fmt(vl - vc)}\n\n",
                        "err")
            else:
                saida.insert(tk.END, "  ✅ Nenhuma divergência encontrada.\n", "ok")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # ── Barra de botões ──────────────────────────────────────────────────────
    btn_bar = tk.Frame(parent, bg=_C("BG"), pady=6)
    btn_bar.pack(fill="x", padx=12)

    _btn(btn_bar, "▶  Executar", executar).pack(side="left", padx=(0, 6))
    _btn(btn_bar, "🗑  Limpar",
         lambda: saida.delete("1.0", tk.END), danger=True).pack(side="left")
    if on_voltar:
        _btn(btn_bar, "← Voltar", on_voltar, secondary=True).pack(side="right")


# ══════════════════════════════════════════════════════════════════════════════
#  TELA PRINCIPAL — submenu com abas (Notebook estilizado)
# ══════════════════════════════════════════════════════════════════════════════

def _estilizar_notebook(root):
    """Aplica estilo moderno ao Notebook usando as cores do theme."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("LF.TNotebook",
                    background=_C("BG"),
                    borderwidth=0,
                    tabmargins=[0, 0, 0, 0])
    style.configure("LF.TNotebook.Tab",
                    background=_C("BG2"),
                    foreground=_C("MUTED"),
                    font=("Segoe UI Semibold", 10),
                    padding=[18, 8],
                    borderwidth=0)
    style.map("LF.TNotebook.Tab",
              background=[("selected", _C("PRIMARY")), ("active", _C("PRIMARY_H"))],
              foreground=[("selected", "white"),       ("active", "white")])


def _build_content(root):
    """
    Monta toda a UI do módulo Livro Fiscal dentro de `root`
    (pode ser Tk, Toplevel ou qualquer Frame pai).
    """
    root.configure(bg=_C("BG"))
    _estilizar_notebook(root)

    nb = ttk.Notebook(root, style="LF.TNotebook")
    nb.pack(fill="both", expand=True, padx=0, pady=0)

    aba1 = tk.Frame(nb, bg=_C("BG"))
    aba2 = tk.Frame(nb, bg=_C("BG"))
    nb.add(aba1, text="  📄 Livro × Relatórios  ")
    nb.add(aba2, text="  📊 Livro × Contabilidade  ")

    _build_relatorios(aba1)
    _build_contabilidade(aba2)


# ══════════════════════════════════════════════════════════════════════════════
#  STANDALONE
# ══════════════════════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    root.title("Livro Fiscal")
    root.geometry("1100x740")
    root.minsize(900, 580)
    core.theme.apply_theme(root)   # aplica paleta base do sistema
    _build_content(root)
    root.mainloop()


if __name__ == "__main__":
    main()