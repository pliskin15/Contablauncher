import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import pandas as pd
import pdfplumber
import re
import os

# ─────────────────────────────────────────────
#  CFOPs desta aplicação
# ─────────────────────────────────────────────
CFOPS_ALVO = {
    "1.202","1.209","1.210","1.411","1.553","1.661","1.662",
    "2.201","2.202","2.203","2.209","2.210","2.411","2.553",
    "2.661","2.918","2.921","3.202",
}
CONTA_CONTAB = 446

# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────
def _eh_moeda(s):
    s = str(s).strip()
    return bool(
        re.fullmatch(r"[0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2}", s)
        or re.fullmatch(r"[0-9]+[.,][0-9]{2}", s)
    )

def _parse_valor(v):
    if not v: return 0.0
    return float(str(v).replace(".", "").replace(",", "."))

def _fmt(v):
    try:
        return (f"R$ {float(v):,.2f}"
                .replace(",","X").replace(".",",").replace("X","."))
    except:
        return "R$ 0,00"

def _extrair_data(tokens):
    for t in tokens:
        if re.match(r"\d{2}[/\-]\d{2}[/\-]\d{4}", t):
            return t
    return None

# ─────────────────────────────────────────────
#  Extração do CFOP  (layout Entradas: valor → CFOP → ICMS)
# ─────────────────────────────────────────────
def _extrair_cfop(tokens):
    if "ICMS" in tokens:
        idx = tokens.index("ICMS")
        for i in range(max(0, idx - 4), idx):
            if tokens[i] in CFOPS_ALVO:
                return tokens[i]
    for t in tokens:
        if t in CFOPS_ALVO:
            return t
    return None

def _valor_contabil_livro(tokens, cfop_tok):
    for i, t in enumerate(tokens):
        if t == cfop_tok and i > 0:
            for j in range(i - 1, -1, -1):
                if _eh_moeda(tokens[j]):
                    return tokens[j]
    return None

# ─────────────────────────────────────────────
#  Leitura do Livro (PDF)
#  Retorna:
#    por_dia   → {dd/mm: {cfop: float}}
#    registros → {dd/mm: [ {esp, ser, numero, emitente, cfop, valor} ]}
# ─────────────────────────────────────────────
def ler_livro(pdf_path):
    por_dia   = {}
    registros = {}
    if not os.path.exists(pdf_path):
        return por_dia, registros

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            for linha in texto.splitlines():
                tokens = linha.split()
                cfop = _extrair_cfop(tokens)
                if cfop is None:
                    continue
                valor_str = _valor_contabil_livro(tokens, cfop)
                if valor_str is None:
                    continue
                data_raw = _extrair_data(tokens)
                if data_raw is None:
                    continue
                try:
                    data_fmt = pd.to_datetime(data_raw, dayfirst=True).strftime("%d/%m")
                except:
                    continue

                por_dia.setdefault(data_fmt, {})
                por_dia[data_fmt][cfop] = por_dia[data_fmt].get(cfop, 0.0) + _parse_valor(valor_str)

                esp    = tokens[1] if len(tokens) > 1 else ""
                ser    = tokens[2] if len(tokens) > 2 else ""
                numero = tokens[3] if len(tokens) > 3 else ""

                # emitente: tokens após ICMS + código, retirando inscrição estadual (só dígitos) no fim
                emitente = ""
                if "ICMS" in tokens:
                    icms_i      = tokens.index("ICMS")
                    nome_tokens = tokens[icms_i + 2:]
                    if nome_tokens and re.fullmatch(r"\d+", nome_tokens[-1]):
                        nome_tokens = nome_tokens[:-1]
                    emitente = " ".join(nome_tokens)

                registros.setdefault(data_fmt, []).append({
                    "esp":      esp,
                    "ser":      ser,
                    "numero":   numero,
                    "emitente": emitente,
                    "cfop":     cfop,
                    "valor":    _parse_valor(valor_str),
                })

    return por_dia, registros


# ─────────────────────────────────────────────
#  Leitura da Contabilidade (Excel, conta 446)
#  Retorna:
#    por_dia   → {dd/mm: float}
#    registros → {dd/mm: [ {seq, doc, hist, conta_partida, valor} ]}
# ─────────────────────────────────────────────
def _extrair_conta_num(texto):
    if not isinstance(texto, str):
        texto = "" if pd.isna(texto) else str(texto)
    m = re.search(r'\b(\d{3,4})\b', texto.upper())
    if m:
        try: return int(m.group(1))
        except: pass
    return None

def ler_contabilidade(excel_path):
    por_dia = {}
    registros = {}

    if not excel_path or not os.path.exists(excel_path):
        return por_dia, registros

    try:
        # Escolhe engine conforme extensão (.xls antigo exige xlrd)
        ext = os.path.splitext(excel_path)[1].lower()
        engine = "xlrd" if ext == ".xls" else None

        # Lê TODAS as abas
        all_sheets = pd.read_excel(excel_path, header=None, sheet_name=None,
                                   engine=engine)

        # Junta tudo
        df = pd.concat(all_sheets.values(), ignore_index=True)

    except Exception as e:
        messagebox.showerror("Erro",
            f"Não foi possível ler o Excel:\n{e}\n\n"
            f"Dica: arquivos .xls precisam da biblioteca 'xlrd'.\n"
            f"Instale com:  pip install xlrd")
        return por_dia, registros

    COL_DATA, COL_SEQ, COL_DOC, COL_HIST, COL_CONTA, COL_DEB = 0, 1, 4, 5, 7, 8
    dentro_bloco = False

    for _, row in df.iterrows():
        texto_bloco = (
            ("" if pd.isna(row[COL_CONTA]) else str(row[COL_CONTA])) + " " +
            ("" if pd.isna(row[COL_HIST])  else str(row[COL_HIST]))
        ).upper()

        if str(CONTA_CONTAB) in texto_bloco or "DEVOLU" in texto_bloco:
            nb = _extrair_conta_num(texto_bloco)
            if nb == CONTA_CONTAB:
                dentro_bloco = True
                continue

        if not dentro_bloco:
            continue

        cp_txt = "" if pd.isna(row[COL_CONTA]) else str(row[COL_CONTA]).upper()
        if "TOTAL DEB" in cp_txt or "TOTAL CRED" in cp_txt:
            dentro_bloco = False
            continue

        valor = pd.to_numeric(row[COL_DEB], errors="coerce")
        if pd.isna(valor) or float(valor) == 0.0:
            continue

        data_raw = row[COL_DATA]
        if isinstance(data_raw, (pd.Timestamp, datetime)):
            data_fmt = data_raw.strftime("%d/%m")
        else:
            m = re.search(r"(\d{2}[/-]\d{2})(?:[/-](\d{4}))?", str(data_raw))
            if not m:
                continue
            dm  = m.group(1).replace("-","/")
            ano = m.group(2) if m.group(2) else str(datetime.today().year)
            try:
                data_fmt = pd.to_datetime(f"{dm}/{ano}", dayfirst=True).strftime("%d/%m")
            except:
                continue

        por_dia[data_fmt] = por_dia.get(data_fmt, 0.0) + float(valor)

        registros.setdefault(data_fmt, []).append({
            "seq":           "" if pd.isna(row[COL_SEQ])  else str(row[COL_SEQ]).strip(),
            "doc":           "" if pd.isna(row[COL_DOC])  else str(row[COL_DOC]).strip(),
            "hist":          "" if pd.isna(row[COL_HIST]) else str(row[COL_HIST]).strip(),
            "conta_partida": "" if pd.isna(row[COL_CONTA])else str(row[COL_CONTA]).strip(),
            "valor":         float(valor),
        })

    return por_dia, registros


# ─────────────────────────────────────────────
#  Helpers de divergência
# ─────────────────────────────────────────────

def _regs_sem_par(regs_origem, valores_destino, tol=0.02):
    """
    Retorna os registros de `regs_origem` cujo valor não tem correspondente
    (dentro da tolerância `tol`) na lista `valores_destino`.
    Usa matching one-to-one: cada valor de destino só "cobre" um registro.
    """
    disponiveis = list(valores_destino)  # cópia para consumir
    sem_par = []
    for r in regs_origem:
        v = r["valor"]
        encontrou = False
        for i, vd in enumerate(disponiveis):
            if abs(v - vd) <= tol:
                disponiveis.pop(i)
                encontrou = True
                break
        if not encontrou:
            sem_par.append(r)
    return sem_par


def _buscar_candidatos(regs, alvo, tol=0.02):
    """
    Tenta encontrar o subconjunto de `regs` cuja soma ≈ `alvo`.
    Retorna a lista de registros se encontrar, [] caso contrário.
    Limita a busca para evitar explosão combinatória (≤ 20 itens).
    """
    if not regs or alvo <= 0:
        return []

    n = len(regs)
    if n > 20:
        return []  # muitos itens, pula busca exaustiva

    from itertools import combinations
    valores = [r["valor"] for r in regs]

    for tamanho in range(1, n + 1):
        for indices in combinations(range(n), tamanho):
            if abs(sum(valores[i] for i in indices) - alvo) <= tol:
                return [regs[i] for i in indices]
    return []


# ─────────────────────────────────────────────
#  Interface
# ─────────────────────────────────────────────
from core import theme as _t     # identidade visual do Prisma (core/theme.py)

BG      = _t.BG_APP
BG2     = _t.CARD_BG
BG3     = _t.SURFACE2
BORDA   = _t.GRID_LINE
AZUL    = _t.ACCENT_BLUE
AZUL2   = _t.PRIMARY
VERDE   = _t.ACCENT_GREEN
VERM    = _t.ACCENT_RED
AMARELO = _t.ACCENT_ORANGE
TEXTO   = _t.FG_TEXT
SUBTXT  = _t.FG_MUTED
MONO    = "Consolas"
LABEL_F = ("Segoe UI", 9)
BTN_F   = ("Segoe UI", 9, "bold")
MONO_F  = (MONO, 9)


def _build_content(root):
    hdr = tk.Frame(root, bg=BG, pady=14, padx=24)
    hdr.pack(fill="x")
    tk.Label(hdr, text="LIVRO DE ENTRADAS", font=("Segoe UI",11,"bold"),
             bg=BG, fg=AZUL).pack(side="left")
    tk.Label(hdr, text="  ×  ", font=("Segoe UI",11), bg=BG, fg=SUBTXT).pack(side="left")
    tk.Label(hdr, text="CONTABILIDADE", font=("Segoe UI",11,"bold"),
             bg=BG, fg=AMARELO).pack(side="left")
    tk.Label(hdr, text=f"  —  conta {CONTA_CONTAB}",
             font=("Segoe UI",9), bg=BG, fg=SUBTXT).pack(side="left")
    tk.Label(hdr, text="CFOPs: " + "  ".join(sorted(CFOPS_ALVO)),
             font=("Segoe UI",7), bg=BG, fg=SUBTXT).pack(side="right")

    tk.Frame(root, bg=BORDA, height=1).pack(fill="x")

    top = tk.Frame(root, bg=BG2, padx=20, pady=16)
    top.pack(fill="x")

    def _entry_row(label, color, row):
        tk.Label(top, text=label, font=LABEL_F, bg=BG2, fg=color,
                 width=22, anchor="w").grid(row=row, column=0, padx=(0,8), pady=5)
        e = tk.Entry(top, bg=BG3, fg=TEXTO, insertbackground=TEXTO,
                     relief="flat", font=MONO_F, bd=0,
                     highlightthickness=1, highlightbackground=BORDA, highlightcolor=AZUL)
        e.grid(row=row, column=1, sticky="ew", padx=(0,8), ipady=5)
        top.columnconfigure(1, weight=1)
        return e

    def _btn(txt, cmd, col=AZUL):
        return tk.Button(top, text=txt, command=cmd, bg=col, fg="#fff",
                         font=BTN_F, relief="flat", bd=0,
                         activebackground=AZUL2, activeforeground="#fff",
                         padx=12, pady=6, cursor="hand2")

    entry_pdf   = _entry_row("📘 Livro (PDF)", AZUL, 0)
    entry_excel = _entry_row("📊 Contabilidade (Excel)", AMARELO, 1)

    def _sel_pdf():
        p = filedialog.askopenfilename(filetypes=[("PDF","*.pdf")])
        if p: entry_pdf.delete(0,"end"); entry_pdf.insert(0, p)

    def _sel_excel():
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if p: entry_excel.delete(0,"end"); entry_excel.insert(0, p)

    _btn("Selecionar", _sel_pdf).grid(row=0, column=2, padx=4)
    _btn("Selecionar", _sel_excel, col=_t.PRIMARY).grid(row=1, column=2, padx=4)

    btn_exec = _btn("▶  Executar análise", None, col=VERDE)
    btn_exec.grid(row=0, column=3, rowspan=2, padx=(16,0), pady=4, sticky="ns")

    tk.Frame(root, bg=BORDA, height=1).pack(fill="x")

    frame_out = tk.Frame(root, bg=BG)
    frame_out.pack(fill="both", expand=True)

    saida = tk.Text(frame_out, bg=BG2, fg=TEXTO, font=MONO_F,
                    insertbackground=TEXTO, relief="flat", bd=0,
                    padx=20, pady=16, wrap="none",
                    selectbackground=AZUL2, selectforeground="#fff")
    vsb = tk.Scrollbar(frame_out, orient="vertical",  command=saida.yview, bg=BG3)
    hsb = tk.Scrollbar(frame_out, orient="horizontal", command=saida.xview, bg=BG3)
    saida.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right",  fill="y")
    hsb.pack(side="bottom", fill="x")
    saida.pack(fill="both", expand=True)

    saida.tag_configure("titulo",    foreground=AZUL,    font=(MONO, 9, "bold"))
    saida.tag_configure("dia",       foreground=AMARELO, font=(MONO, 9, "bold"))
    saida.tag_configure("livro",     foreground="#60a5fa")
    saida.tag_configure("contab",    foreground="#fbbf24")
    saida.tag_configure("ok",        foreground=VERDE)
    saida.tag_configure("dif",       foreground=VERM,    font=(MONO, 9, "bold"))
    saida.tag_configure("sep",       foreground=SUBTXT)
    saida.tag_configure("cfop_hdr",  foreground="#a78bfa", font=(MONO, 8, "bold"))
    saida.tag_configure("cfop_ln",   foreground="#c4b5fd", font=(MONO, 8))
    saida.tag_configure("div_hdr",   foreground=VERM,    font=(MONO, 8, "bold"))
    saida.tag_configure("div_livro", foreground="#93c5fd", font=(MONO, 8))
    saida.tag_configure("div_cont",  foreground="#fcd34d", font=(MONO, 8))

    def _ins(txt, tag=""):
        saida.insert("end", txt, tag)

    def executar():
        pdf_path   = entry_pdf.get().strip()
        excel_path = entry_excel.get().strip()

        if not pdf_path:
            messagebox.showwarning("Aviso", "Selecione o PDF do livro fiscal.")
            return

        saida.config(state="normal")
        saida.delete("1.0","end")
        _ins("⏳ Lendo arquivos…\n", "sep")
        root.update_idletasks()

        livro_por_dia, livro_regs = ler_livro(pdf_path)

        if excel_path:
            cont_por_dia, cont_regs = ler_contabilidade(excel_path)
            if excel_path and not cont_por_dia:
                _ins("⚠️  Excel carregado mas nenhum lançamento da conta 446 foi encontrado.\n", "dif")
                _ins("   Verifique se o arquivo é o razão correto e se a conta 446 está presente.\n\n", "sep")
        else:
            cont_por_dia, cont_regs = {}, {}

        saida.delete("1.0","end")

        todos_os_dias = sorted(
            set(livro_por_dia) | set(cont_por_dia),
            key=lambda d: datetime.strptime(d, "%d/%m")
        )

        if not todos_os_dias:
            _ins("⚠️  Nenhum dado encontrado. Verifique os arquivos.\n", "dif")
            return

        _ins("📘 LIVRO DE ENTRADAS  ×  CONTABILIDADE (conta 446)\n", "titulo")
        _ins(f"CFOPs: {', '.join(sorted(CFOPS_ALVO))}\n", "sep")
        _ins("─" * 90 + "\n\n", "sep")

        _ins("DETALHE DIÁRIO\n", "titulo")
        _ins("─" * 90 + "\n", "sep")

        divergencias = []

        for dia in todos_os_dias:
            cfops_dia   = livro_por_dia.get(dia, {})
            total_livro = sum(cfops_dia.values())
            total_cont  = cont_por_dia.get(dia, 0.0)
            dif         = total_livro - total_cont

            _ins(f"\n• {dia}\n", "dia")
            _ins(f"  {'Livro total':30s}  {_fmt(total_livro)}\n", "livro")
            _ins(f"  {'Contabilidade (446)':30s}  {_fmt(total_cont)}\n", "contab")

            if abs(dif) < 0.01:
                _ins(f"  {'Diferença':30s}  {_fmt(0.0)}  ✓\n", "ok")
            else:
                _ins(f"  {'Diferença':30s}  {_fmt(dif)}  ◀ DIVERGÊNCIA\n", "dif")
                divergencias.append((dia, total_livro, total_cont, dif))

            if cfops_dia:
                _ins("  CFOPs do dia:\n", "cfop_hdr")
                for cfop, val in sorted(cfops_dia.items()):
                    _ins(f"    {cfop}  →  {_fmt(val)}\n", "cfop_ln")

        # ── Somatória do período ──
        _ins("\n" + "═" * 90 + "\n", "sep")
        _ins("📊 SOMATÓRIA DO PERÍODO\n", "titulo")
        _ins("─" * 90 + "\n", "sep")

        tot_livro = sum(sum(v.values()) for v in livro_por_dia.values())
        tot_cont  = sum(cont_por_dia.values())
        tot_dif   = tot_livro - tot_cont

        _ins(f"\n  {'Livro (total período)':38s}  {_fmt(tot_livro)}\n", "livro")
        _ins(f"  {'Contabilidade 446 (total período)':38s}  {_fmt(tot_cont)}\n", "contab")
        tag_t = "ok" if abs(tot_dif) < 0.01 else "dif"
        suf   = "  ✓ OK" if abs(tot_dif) < 0.01 else "  ◀ DIVERGÊNCIA"
        _ins(f"  {'Diferença':38s}  {_fmt(tot_dif)}{suf}\n", tag_t)

        _ins("\n  Breakdown por CFOP (período completo):\n", "cfop_hdr")
        totais_cfop: dict = {}
        for d_cfops in livro_por_dia.values():
            for cfop, val in d_cfops.items():
                totais_cfop[cfop] = totais_cfop.get(cfop, 0.0) + val
        for cfop, val in sorted(totais_cfop.items()):
            _ins(f"    {cfop}  →  {_fmt(val)}\n", "cfop_ln")

        # ── Divergências detalhadas ──
        _ins("\n" + "═" * 90 + "\n", "sep")

        if not divergencias:
            _ins("✅ Nenhuma divergência encontrada.\n", "ok")
        else:
            _ins(f"🚨 {len(divergencias)} DIVERGÊNCIA(S) — DETALHAMENTO\n", "dif")

            for dia, vl, vc, dif in divergencias:
                _ins("─" * 90 + "\n", "sep")
                _ins(f"  Dia {dia}  |  Livro: {_fmt(vl)}  |  Contab: {_fmt(vc)}  |  Dif: {_fmt(dif)}\n", "dif")

                if dif > 0:
                    # ── Livro > Contabilidade ──
                    # Tenta identificar lançamentos do livro sem par na contabilidade
                    regs_livro = livro_regs.get(dia, [])
                    regs_cont  = cont_regs.get(dia, [])

                    # Monta lista de valores da contabilidade para matching
                    cont_valores = sorted([r["valor"] for r in regs_cont])

                    # Tentativa de subset-sum: quais notas do livro somam a diferença
                    candidatos = _buscar_candidatos(regs_livro, abs(dif))

                    if regs_livro:
                        _ins(f"\n  ↳ Livro maior (+{_fmt(abs(dif))}) — lançamentos do livro sem correspondência na contabilidade:\n", "div_hdr")
                        _ins(
                            f"    {'ESP':<5} {'SER':<5} {'NÚMERO':<12} {'CFOP':<7} {'VALOR':>14}  EMITENTE\n",
                            "div_hdr"
                        )
                        _ins("    " + "·" * 82 + "\n", "sep")

                        # Se encontrou subconjunto exato, destaca só eles; senão mostra todos
                        exibir = candidatos if candidatos else regs_livro
                        sem_par = _regs_sem_par(regs_livro, cont_valores)
                        if sem_par:
                            exibir = sem_par

                        for r in exibir:
                            _ins(
                                f"    {r['esp']:<5} {r['ser']:<5} {r['numero']:<12} "
                                f"{r['cfop']:<7} {_fmt(r['valor']):>14}  {r['emitente']}\n",
                                "div_livro"
                            )
                        _ins(
                            f"    {'':5} {'':5} {'TOTAL':>12}  {'':7} {_fmt(sum(r['valor'] for r in exibir)):>14}",
                            "div_hdr"
                        )
                        if not candidatos and not sem_par:
                            _ins("  (todos os lançamentos do dia exibidos)", "sep")
                        _ins("\n", "")

                else:
                    # ── Contabilidade > Livro ──
                    regs_cont  = cont_regs.get(dia, [])
                    regs_livro = livro_regs.get(dia, [])

                    # Monta lista de valores do livro para matching
                    livro_valores = sorted([r["valor"] for r in regs_livro])

                    candidatos = _buscar_candidatos(regs_cont, abs(dif))

                    if regs_cont:
                        _ins(f"\n  ↳ Contabilidade maior (+{_fmt(abs(dif))}) — lançamentos da contabilidade sem correspondência no livro:\n", "div_hdr")
                        _ins(
                            f"    {'SEQ':<10} {'DOC':<12} {'CONTA PART.':<16} {'VALOR':>14}  HISTÓRICO\n",
                            "div_hdr"
                        )
                        _ins("    " + "·" * 82 + "\n", "sep")

                        sem_par = _regs_sem_par(regs_cont, livro_valores)
                        exibir  = candidatos if candidatos else (sem_par if sem_par else regs_cont)

                        for r in exibir:
                            _ins(
                                f"    {r['seq']:<10} {r['doc']:<12} {r['conta_partida']:<16} "
                                f"{_fmt(r['valor']):>14}  {r['hist']}\n",
                                "div_cont"
                            )
                        _ins(
                            f"    {'':10} {'':12} {'TOTAL':>16} {_fmt(sum(r['valor'] for r in exibir)):>14}",
                            "div_hdr"
                        )
                        if not candidatos and not sem_par:
                            _ins("  (todos os lançamentos do dia exibidos)", "sep")
                        _ins("\n", "")

        saida.config(state="disabled")

    btn_exec.config(command=executar)

    tk.Frame(root, bg=BORDA, height=1).pack(fill="x")
    tk.Label(root, text="Selecione os arquivos e clique em Executar.",
             font=("Segoe UI",8), bg=BG, fg=SUBTXT, anchor="w", padx=16, pady=4).pack(fill="x")


def criar_janela():
    root = tk.Tk()
    root.title("Livro de Entradas × Contabilidade  (conta 446)")
    root.configure(bg=BG)
    root.geometry("1200x820")
    root.minsize(900, 600)
    _build_content(root)
    root.mainloop()


if __name__ == "__main__":
    criar_janela()