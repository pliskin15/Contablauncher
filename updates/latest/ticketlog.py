"""Relatório Ticket Log — concilia, dia a dia, o relatório da Ticket Log (CNPJ GOODCARD)
com o Razão da contabilidade e aponta o que está na Ticket Log e não foi achado no Razão.

Regras
  Ticket Log : só as abas cujo nome tem exclusivamente números; data-base = coluna
               PAGAMENTO; valor = coluna VALOR BRUTO.
  Razão      : (Razão Modelo I) só a coluna CRÉDITO das linhas cuja CONTA PARTIDA seja
               "xxxx - N" com N em CONTAS_PARTIDA (12, 11 e 9; "- 1253" ou "- 91" não entram).
"""
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime

import pandas as pd

TOL = 0.01          # diferença (R$) tolerada para considerar o dia conciliado
CONTAS_PARTIDA = (12, 11, 9)       # sufixos "xxxx - N" aceitos na CONTA PARTIDA do Razão
_RE_CP = re.compile(r"^\s*\d+\s*-\s*(?:%s)\s*$" % "|".join(str(c) for c in CONTAS_PARTIDA))
_ROT_CP = "xxxx-" + "/".join(str(c) for c in CONTAS_PARTIDA)


# ══════════════════════════════════════════════════════════════════════════════
#  LEITURA
# ══════════════════════════════════════════════════════════════════════════════
def _norm(txt) -> str:
    s = unicodedata.normalize("NFKD", str(txt))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper()


def _num(v):
    """Converte valor de célula em float (aceita 1.234,56 / 1234.56 / R$). None se inválido."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").replace(" ", "").strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        r = float(s)
    except ValueError:
        return None
    return -r if neg else r


def _data(v):
    """Converte célula em date. Aceita datetime, serial do Excel e texto dd/mm/aaaa."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        if 20000 < v < 80000:                       # serial do Excel
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
        return None
    s = str(v).strip()[:19]
    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _excel(path):
    engine = "xlrd" if str(path).lower().endswith(".xls") else None
    return pd.ExcelFile(path, **({"engine": engine} if engine else {}))


def aba_numerica(nome: str) -> bool:
    return bool(re.fullmatch(r"\d+", str(nome).strip()))


def parse_ticket(path):
    """Retorna (DataFrame[aba, linha, data, valor], avisos)."""
    xl = _excel(path)
    rows, avisos = [], []
    abas = [s for s in xl.sheet_names if aba_numerica(s)]
    if not abas:
        raise ValueError("Nenhuma aba com nome somente numérico foi encontrada no relatório da Ticket Log.")

    for aba in abas:
        raw = xl.parse(aba, header=None, dtype=object)
        hdr = ic_data = ic_val = ic_reemb = None
        for i in range(min(len(raw), 40)):
            nomes = [_norm(v) if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
            c_val = [j for j, n in enumerate(nomes) if n == "VALOR BRUTO"] or \
                    [j for j, n in enumerate(nomes) if "VALOR BRUTO" in n]
            c_dat = [j for j, n in enumerate(nomes) if n == "PAGAMENTO"] or \
                    [j for j, n in enumerate(nomes) if "PAGAMENTO" in n]
            c_re = [j for j, n in enumerate(nomes) if n == "VALOR DE REEMBOLSO"] or \
                   [j for j, n in enumerate(nomes) if "REEMBOLSO" in n]
            if c_val and c_dat:
                hdr, ic_data, ic_val = i, c_dat[0], c_val[0]
                ic_reemb = c_re[0] if c_re else None
                ic_ext = {}
                for chave, rot in (("estab", "ESTABELECIMENTO"), ("lote", "LOTE"), ("cliente", "CLIENTE"),
                                   ("nota", "NOTA FISCAL"), ("status", "STATUS PAGAMENTO"),
                                   ("desconto", "VALOR DESCONTADO")):
                    ic_ext[chave] = nomes.index(rot) if rot in nomes else None
                break
        if hdr is None:
            avisos.append(f"Aba {aba}: colunas PAGAMENTO / VALOR BRUTO não encontradas — ignorada.")
            continue
        for i in range(hdr + 1, len(raw)):
            cel = raw.iat[i, ic_data]
            d = _data(cel)
            v = _num(raw.iat[i, ic_val])
            if d is None or v is None:
                continue
            real = isinstance(cel, (pd.Timestamp, datetime, date))     # data "de verdade" do Excel
            rv = _num(raw.iat[i, ic_reemb]) if ic_reemb is not None else None
            rec = {"aba": str(aba).strip(), "linha": i + 1, "data": d, "valor": round(v, 2),
                   "reembolso": None if rv is None else round(rv, 2), "_real": real}
            for chave, j in ic_ext.items():
                x = raw.iat[i, j] if j is not None else None
                if chave == "desconto":
                    rec[chave] = _num(x)
                else:
                    rec[chave] = "" if x is None or (isinstance(x, float) and pd.isna(x)) else \
                        (str(int(x)) if isinstance(x, float) and x == int(x) else str(x).strip())
            rows.append(rec)
    df = pd.DataFrame(rows, columns=["aba", "linha", "data", "valor", "reembolso", "estab", "lote", "cliente",
                                     "nota", "status", "desconto", "_real"])
    if df.empty:
        raise ValueError("Nenhuma linha válida (PAGAMENTO + VALOR BRUTO) nas abas numéricas da Ticket Log.")

    # O relatório mistura texto dd/mm/aaaa com datas reais do Excel. Quando o dia era <= 12, o
    # Excel interpretou como mm/dd e gravou data real com dia e mês INVERTIDOS. Sinal: todas as
    # datas reais têm dia <= 12 (nunca > 12) e existem datas em texto. Nesse caso desinverte.
    # A decisão é por aba: aba com datas em texto e datas reais só com dia <= 12 = invertidas.
    # Aba só com datas reais (ou com alguma de dia > 12) é considerada correta.
    n_corr = 0
    for aba_, g in df.groupby("aba"):
        reais = g[g["_real"]]
        if len(reais) and not g["_real"].all() and all(d.day <= 12 for d in reais["data"]):
            df.loc[reais.index, "data"] = [date(d.year, d.day, d.month) for d in reais["data"]]
            n_corr += len(reais)
    if n_corr:
        avisos.append(f"{n_corr} data(s) de PAGAMENTO estavam com dia/mês invertidos no Excel e foram corrigidas.")
    return df.drop(columns="_real"), avisos


def parse_razao(path):
    """Retorna DataFrame[filial, linha, data, doc, historico, conta_partida, valor] só com
    CRÉDITO das contrapartidas 'xxxx - 12', 'xxxx - 11' e 'xxxx - 9'."""
    xl = _excel(path)
    rows = []
    for aba in xl.sheet_names:
        if "RESUMO" in aba.upper():
            continue
        raw = xl.parse(aba, header=None, dtype=object)
        filial, idx = None, None
        for i in range(len(raw)):
            vals = raw.iloc[i].tolist()
            first = _norm(vals[0]) if pd.notna(vals[0]) else ""
            if first == "FILIAL":
                filial = str(vals[1]).strip() if len(vals) > 1 and pd.notna(vals[1]) else None
                idx = None
                continue
            if first == "DATA":
                nomes = [_norm(v) if pd.notna(v) else "" for v in vals]
                try:
                    idx = {
                        "cp": nomes.index("CONTA PARTIDA"),
                        "cred": nomes.index("CREDITO"),
                        "hist": nomes.index("HISTORICO") if "HISTORICO" in nomes else None,
                        "doc": nomes.index("DOC. NRO.") if "DOC. NRO." in nomes else None,
                    }
                except ValueError:
                    idx = None
                continue
            if idx is None:
                continue
            d = _data(vals[0])
            if d is None:                          # totais / rodapé
                continue
            cp = str(vals[idx["cp"]]).strip() if pd.notna(vals[idx["cp"]]) else ""
            if not _RE_CP.match(cp):
                continue
            cred = _num(vals[idx["cred"]])
            if not cred:
                continue
            rows.append({
                "filial": filial or str(aba).strip(),
                "linha": i + 1,
                "data": d,
                "doc": str(vals[idx["doc"]]).strip() if idx["doc"] is not None and pd.notna(vals[idx["doc"]]) else "",
                "historico": str(vals[idx["hist"]]).strip() if idx["hist"] is not None and pd.notna(vals[idx["hist"]]) else "",
                "conta_partida": cp,
                "valor": round(cred, 2),
            })
    df = pd.DataFrame(rows, columns=["filial", "linha", "data", "doc", "historico", "conta_partida", "valor"])
    if df.empty:
        raise ValueError(f"Nenhum crédito com contrapartida '{_ROT_CP}' no Razão. "
                         "Verifique se o arquivo está no formato Razão Modelo I.")
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  CONCILIAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
def conciliar(ticket: pd.DataFrame, razao: pd.DataFrame, tol: float = TOL):
    """Retorna (dias, ticket_sem_razao, razao_sem_ticket).

    dias: lista de dicts por data (ordenada) com totais, diferença e status.
    ticket_sem_razao / razao_sem_ticket: DataFrames com as linhas sem par de mesmo valor
    no mesmo dia (só para dias com diferença)."""
    t_by = {d: g for d, g in ticket.groupby("data")}
    r_by = {d: g for d, g in razao.groupby("data")}
    dias, t_pend, r_pend = [], [], []

    for d in sorted(set(t_by) | set(r_by)):
        tg = t_by.get(d)
        rg = r_by.get(d)
        t_tot = round(float(tg["valor"].sum()), 2) if tg is not None else 0.0
        r_tot = round(float(rg["valor"].sum()), 2) if rg is not None else 0.0
        dif = round(t_tot - r_tot, 2)
        if tg is None:
            status = "Só no Razão"
        elif rg is None:
            status = "Só na Ticket"
        elif abs(dif) <= tol:
            status = "OK"
        elif dif > 0:
            status = "Falta no Razão"
        else:
            status = "Razão a maior"
        dias.append({
            "data": d, "ticket": t_tot, "razao": r_tot, "dif": dif, "status": status,
            "qtd_ticket": 0 if tg is None else len(tg), "qtd_razao": 0 if rg is None else len(rg),
        })
        if status == "OK":
            continue

        # pareamento linha a linha (mesmo valor, mesmo dia) para apontar o que sobra
        livres = defaultdict(list)
        if rg is not None:
            for rec in rg.to_dict("records"):
                livres[round(rec["valor"], 2)].append(rec)
        if tg is not None:
            for rec in tg.to_dict("records"):
                lista = livres.get(round(rec["valor"], 2))
                if lista:
                    lista.pop()
                else:
                    t_pend.append(rec)
        for lista in livres.values():
            r_pend.extend(lista)

    cols_t = ["aba", "linha", "data", "valor"]
    cols_r = ["filial", "linha", "data", "doc", "historico", "conta_partida", "valor"]
    return (dias,
            pd.DataFrame(t_pend, columns=cols_t).sort_values(["data", "valor"]).reset_index(drop=True),
            pd.DataFrame(r_pend, columns=cols_r).sort_values(["data", "valor"]).reset_index(drop=True))


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORTAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
def exportar_excel(dias, t_pend, r_pend, ticket, razao, path, rotulo="Valor Bruto"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    FMT = "#,##0.00"
    COR = {"OK": "C6EFCE", "Falta no Razão": "FFC7CE", "Só na Ticket": "FFC7CE",
           "Razão a maior": "FFEB9C", "Só no Razão": "FFEB9C"}

    def aba(ws, titulo, headers, widths, linhas, fmts=None, cor_status_col=None):
        ws.title = titulo
        for c, (h, w) in enumerate(zip(headers, widths), 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            cell.fill = PatternFill("solid", fgColor="2F3655")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            ws.column_dimensions[get_column_letter(c)].width = w
        for r, linha in enumerate(linhas, 2):
            fill = None
            if cor_status_col is not None:
                fill = PatternFill("solid", fgColor=COR.get(linha[cor_status_col], "FFFFFF"))
            for c, v in enumerate(linha, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(name="Arial", size=10)
                cell.border = border
                if fill:
                    cell.fill = fill
                if fmts and fmts.get(c):
                    cell.number_format = fmts[c]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(linhas), 1) + 1}"

    ws = wb.active
    aba(ws, "Diário", ["Data", f"Ticket Log ({rotulo})", f"Razão (Créd. {_ROT_CP})", "Diferença", "Status", "Qtd Ticket", "Qtd Razão"],
        [14, 20, 22, 16, 18, 12, 12],
        [[x["data"], x["ticket"], x["razao"], x["dif"], x["status"], x["qtd_ticket"], x["qtd_razao"]] for x in dias],
        {1: "DD/MM/YYYY", 2: FMT, 3: FMT, 4: FMT}, cor_status_col=4)
    n = len(dias) + 2
    ws.cell(row=n, column=1, value="TOTAL").font = Font(bold=True, name="Arial", size=10)
    for col in (2, 3, 4):
        L = get_column_letter(col)
        c = ws.cell(row=n, column=col, value=f"=SUM({L}2:{L}{n - 1})")
        c.font = Font(bold=True, name="Arial", size=10)
        c.number_format = FMT

    aba(wb.create_sheet(), "Ticket sem Razão", ["Data (PAGAMENTO)", rotulo, "Aba", "Linha na planilha"],
        [18, 16, 10, 18],
        [[r["data"], r["valor"], r["aba"], r["linha"]] for r in t_pend.to_dict("records")],
        {1: "DD/MM/YYYY", 2: FMT})
    aba(wb.create_sheet(), "Razão sem Ticket", ["Data", "Crédito", "Filial", "Doc.", "Conta partida", "Histórico", "Linha na planilha"],
        [14, 16, 10, 14, 16, 60, 18],
        [[r["data"], r["valor"], r["filial"], r["doc"], r["conta_partida"], r["historico"], r["linha"]]
         for r in r_pend.to_dict("records")],
        {1: "DD/MM/YYYY", 2: FMT})
    aba(wb.create_sheet(), "Base Ticket", ["Data (PAGAMENTO)", rotulo, "Aba", "Linha"], [18, 16, 10, 10],
        [[r["data"], r["valor"], r["aba"], r["linha"]] for r in ticket.sort_values(["data", "aba", "linha"]).to_dict("records")],
        {1: "DD/MM/YYYY", 2: FMT})
    aba(wb.create_sheet(), "Base Razão", ["Data", "Crédito", "Filial", "Doc.", "Conta partida", "Histórico", "Linha"],
        [14, 16, 10, 14, 16, 60, 10],
        [[r["data"], r["valor"], r["filial"], r["doc"], r["conta_partida"], r["historico"], r["linha"]]
         for r in razao.sort_values(["data", "linha"]).to_dict("records")],
        {1: "DD/MM/YYYY", 2: FMT})
    wb.save(path)


# ══════════════════════════════════════════════════════════════════════════════
#  TELA
# ══════════════════════════════════════════════════════════════════════════════
def _brl(v) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _dmy(d) -> str:
    return d.strftime("%d/%m/%Y")


def _build_content(win):
    """Monta a tela dentro de um Toplevel/HostFrame (padrão dos demais módulos)."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    from core import theme as _t

    BG, CARD, ACCENT = _t.BG_APP, _t.CARD_BG, _t.ACCENT_BLUE
    GREEN, RED, YELLOW = _t.ACCENT_GREEN, _t.ACCENT_RED, _t.ACCENT_ORANGE
    TXT, TXT2, FIELD, HOVER = _t.FG_TEXT, _t.FG_MUTED, _t.SURFACE2, _t.GRID_LINE
    FONT, FONT_B, FONT_SM = ("Segoe UI", 10), ("Segoe UI", 10, "bold"), ("Segoe UI", 9)

    try:
        win.title("Relatório Ticket Log — PMZ")
        win.geometry("1200x800")
        win.minsize(960, 640)
    except Exception:
        pass
    try:
        win.configure(bg=BG)
    except Exception:
        pass

    st = {"dias": [], "t_pend": None, "r_pend": None, "ticket": None, "razao": None,
          "ticket_full": None, "razao_full": None, "avisos": []}
    v_ticket, v_razao = tk.StringVar(), tk.StringVar()
    v_filtro = tk.StringVar(value="Divergentes")
    v_ano, v_mes = tk.StringVar(value="Todos"), tk.StringVar(value="Todos")
    v_valor = tk.StringVar(value="Valor de Reembolso")
    v_status = tk.StringVar(value="Somente Pago")

    root = tk.Frame(win, bg=BG)
    root.pack(fill="both", expand=True)

    hdr = tk.Frame(root, bg=ACCENT, height=54)
    hdr.pack(fill="x")
    tk.Label(hdr, text="🎫  Relatório Ticket Log × Razão (dia a dia)", bg=ACCENT, fg="white",
             font=("Segoe UI", 14, "bold"), padx=20).pack(side="left", pady=12)

    ff = tk.Frame(root, bg=BG, pady=10)
    ff.pack(fill="x", padx=20)

    def linha_arquivo(rot, var, cmd, r):
        tk.Label(ff, text=rot, bg=BG, fg=TXT2, font=FONT_SM, width=30, anchor="e").grid(
            row=r, column=0, padx=(0, 6), pady=4, sticky="e")
        tk.Entry(ff, textvariable=var, width=62, bg=CARD, fg=TXT, insertbackground=TXT,
                 relief="flat", font=FONT, state="readonly").grid(row=r, column=1, padx=(0, 6))
        tk.Button(ff, text="📂 Abrir", command=cmd, bg=CARD, fg=ACCENT, font=FONT_SM, relief="flat",
                  activebackground=HOVER, cursor="hand2", padx=8).grid(row=r, column=2)

    def pick(var, titulo):
        p = filedialog.askopenfilename(title=titulo, filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if p:
            var.set(p)

    linha_arquivo("Relatório Ticket Log (CNPJ GOODCARD):", v_ticket,
                  lambda: pick(v_ticket, "Selecionar relatório Ticket Log"), 0)
    linha_arquivo("Razão (Modelo I):", v_razao, lambda: pick(v_razao, "Selecionar Razão"), 1)
    tk.Label(ff, text="Valor da Ticket a comparar:", bg=BG, fg=TXT2, font=FONT_SM, width=30, anchor="e").grid(
        row=2, column=0, padx=(0, 6), pady=4, sticky="e")
    cb_valor = ttk.Combobox(ff, textvariable=v_valor, values=["Valor de Reembolso", "Valor Bruto"],
                            width=24, state="readonly")
    cb_valor.grid(row=2, column=1, sticky="w")
    tk.Label(ff, text="Status do pagamento (Ticket):", bg=BG, fg=TXT2, font=FONT_SM, width=30, anchor="e").grid(
        row=3, column=0, padx=(0, 6), pady=4, sticky="e")
    cb_status = ttk.Combobox(ff, textvariable=v_status, values=["Somente Pago", "Todos"],
                             width=24, state="readonly")
    cb_status.grid(row=3, column=1, sticky="w")

    btns = tk.Frame(root, bg=BG)
    btns.pack(pady=(6, 10))
    btn_proc = tk.Button(btns, text="▶  Conciliar", bg=ACCENT, fg="white", font=FONT_B, relief="flat",
                         padx=20, pady=8, activebackground=_t.PRIMARY_HOVER, cursor="hand2")
    btn_proc.pack(side="left", padx=6)
    btn_exp = tk.Button(btns, text="📥  Exportar Excel", bg=GREEN, fg="white", font=FONT_B, relief="flat",
                        padx=20, pady=8, activebackground="#00a862", cursor="hand2", state="disabled")
    btn_exp.pack(side="left", padx=6)

    resumo = tk.Frame(root, bg=BG)
    resumo.pack(fill="x", padx=20, pady=(0, 8))

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background=CARD, foreground=TXT, fieldbackground=CARD, rowheight=24, font=FONT)
    style.configure("Treeview.Heading", background=BG, foreground=_t.ACCENT_CYAN,
                    font=("Segoe UI", 8, "bold"), relief="flat", padding=(6, 6))
    style.map("Treeview", background=[("selected", _t.SEL_BG)], foreground=[("selected", TXT)])

    # ── tabela diária ──
    card1 = tk.Frame(root, bg=CARD)
    card1.pack(fill="both", expand=True, padx=20, pady=(0, 8))
    fr = tk.Frame(card1, bg=CARD)
    fr.pack(fill="x", padx=12, pady=(8, 2))
    tk.Label(fr, text="Dia a dia — Ticket Log × Razão", bg=CARD, fg=TXT, font=FONT_B).pack(side="left")
    for rot in ("Todos", "Divergentes", "OK"):
        tk.Radiobutton(fr, text=rot, variable=v_filtro, value=rot, bg=CARD, fg=TXT2, selectcolor=CARD,
                       activebackground=CARD, font=FONT_SM, command=lambda: preencher()).pack(side="right", padx=4)
    tk.Label(fr, text="Mostrar:", bg=CARD, fg=TXT2, font=FONT_SM).pack(side="right", padx=(12, 0))
    MESES = ["Todos", "01 - Jan", "02 - Fev", "03 - Mar", "04 - Abr", "05 - Mai", "06 - Jun",
             "07 - Jul", "08 - Ago", "09 - Set", "10 - Out", "11 - Nov", "12 - Dez"]
    cb_mes = ttk.Combobox(fr, textvariable=v_mes, values=MESES, width=10, state="readonly")
    cb_mes.pack(side="right", padx=(2, 0))
    tk.Label(fr, text="Mês:", bg=CARD, fg=TXT2, font=FONT_SM).pack(side="right", padx=(10, 0))
    cb_ano = ttk.Combobox(fr, textvariable=v_ano, values=["Todos"], width=7, state="readonly")
    cb_ano.pack(side="right", padx=(2, 0))
    tk.Label(fr, text="Ano:", bg=CARD, fg=TXT2, font=FONT_SM).pack(side="right", padx=(10, 0))

    cols1 = ("data", "ticket", "razao", "dif", "status", "qt", "qr")
    tv = ttk.Treeview(card1, columns=cols1, show="headings", selectmode="browse", height=10)
    for c, (t, w) in {"data": ("Data", 100), "ticket": ("Ticket Log (Bruto)", 150), "razao": (f"Razão ({_ROT_CP})", 150),
                      "dif": ("Diferença", 130), "status": ("Status", 130), "qt": ("Qtd Ticket", 80),
                      "qr": ("Qtd Razão", 80)}.items():
        tv.heading(c, text=t)
        tv.column(c, width=w, anchor="center", minwidth=60)
    tv.tag_configure("ok", foreground=GREEN)
    tv.tag_configure("falta", foreground=RED)
    tv.tag_configure("maior", foreground=YELLOW)
    sb = ttk.Scrollbar(card1, orient="vertical", command=tv.yview)
    tv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tv.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    # ── detalhe do dia ──
    card2 = tk.Frame(root, bg=CARD)
    card2.pack(fill="both", expand=True, padx=20, pady=(0, 8))
    lbl_det = tk.Label(card2, text="Selecione um dia para ver as linhas sem par.", bg=CARD, fg=TXT2,
                       font=FONT_B, anchor="w")
    lbl_det.pack(fill="x", padx=12, pady=(8, 2))
    cols2 = ("origem", "valor", "ref", "info")
    td = ttk.Treeview(card2, columns=cols2, show="headings", selectmode="browse", height=7)
    for c, (t, w, a) in {"origem": ("Origem", 190, "w"), "valor": ("Valor", 130, "center"),
                         "ref": ("Aba / Linha", 130, "center"), "info": ("Histórico / Doc.", 560, "w")}.items():
        td.heading(c, text=t)
        td.column(c, width=w, anchor=a, minwidth=60)
    td.tag_configure("t", foreground=RED)
    td.tag_configure("r", foreground=YELLOW)
    sb2 = ttk.Scrollbar(card2, orient="vertical", command=td.yview)
    td.configure(yscrollcommand=sb2.set)
    sb2.pack(side="right", fill="y")
    td.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    barra = tk.Label(root, text="Selecione os dois arquivos e clique em Conciliar.", bg=BG, fg=TXT2,
                     font=FONT_SM, anchor="w", padx=14)
    barra.pack(fill="x", side="bottom")

    TAG = {"OK": "ok", "Falta no Razão": "falta", "Só na Ticket": "falta"}

    def preencher():
        tv.delete(*tv.get_children())
        f = v_filtro.get()
        for x in st["dias"]:
            if f == "OK" and x["status"] != "OK":
                continue
            if f == "Divergentes" and x["status"] == "OK":
                continue
            tv.insert("", "end", iid=x["data"].isoformat(), tags=(TAG.get(x["status"], "maior"),), values=(
                _dmy(x["data"]), _brl(x["ticket"]), _brl(x["razao"]), _brl(x["dif"]), x["status"],
                x["qtd_ticket"], x["qtd_razao"]))

    def detalhe(_e=None):
        td.delete(*td.get_children())
        sel = tv.selection()
        if not sel or st["t_pend"] is None:
            return
        d = date.fromisoformat(sel[0])
        tp = st["t_pend"][st["t_pend"]["data"] == d]
        rp = st["r_pend"][st["r_pend"]["data"] == d]
        lbl_det.config(text=f"{_dmy(d)} — {len(tp)} linha(s) da Ticket sem par no Razão · "
                            f"{len(rp)} lançamento(s) do Razão sem par na Ticket")
        for r in tp.to_dict("records"):
            td.insert("", "end", tags=("t",), values=("Ticket sem Razão", _brl(r["valor"]),
                                                      f"{r['aba']} / {r['linha']}", ""))
        for r in rp.to_dict("records"):
            td.insert("", "end", tags=("r",), values=("Razão sem Ticket", _brl(r["valor"]),
                                                      f"{r['filial']} / {r['linha']}",
                                                      f"{r['doc']}  {r['historico']}".strip()))

    tv.bind("<<TreeviewSelect>>", detalhe)

    def abrir_dia(_e=None):
        """Janela com TODOS os lançamentos do dia selecionado: Ticket (esq.) e Razão (dir.)."""
        sel = tv.selection()
        if not sel or st["ticket"] is None:
            return
        d = date.fromisoformat(sel[0])
        tk_d = st["ticket"][st["ticket"]["data"] == d].sort_values(["aba", "linha"])
        rz_d = st["razao"][st["razao"]["data"] == d].sort_values("linha")
        sem_t = {(r["aba"], r["linha"]) for r in st["t_pend"].to_dict("records") if r["data"] == d}
        sem_r = {(r["filial"], r["linha"]) for r in st["r_pend"].to_dict("records") if r["data"] == d}
        rot = st.get("rotulo", "Valor Bruto")

        w = tk.Toplevel(win)
        w.title(f"Lançamentos de {_dmy(d)} — Ticket Log × Razão")
        w.geometry("1250x640")
        w.configure(bg=BG)
        tk.Label(w, text=f"{_dmy(d)}   ·   Ticket ({rot}): {_brl(float(tk_d['valor'].sum()))}   ·   "
                         f"Razão: {_brl(float(rz_d['valor'].sum()))}   ·   "
                         f"Diferença: {_brl(float(tk_d['valor'].sum() - rz_d['valor'].sum()))}",
                 bg=BG, fg=TXT, font=("Segoe UI", 11, "bold"), anchor="w", padx=14, pady=8).pack(fill="x")
        tk.Label(w, text="Linhas em vermelho = sem lançamento de mesmo valor no outro lado, no mesmo dia.",
                 bg=BG, fg=TXT2, font=FONT_SM, anchor="w", padx=14).pack(fill="x")

        def tabela(parent, titulo, cols, larg, linhas, tag_fn):
            box = tk.Frame(parent, bg=CARD)
            box.pack(fill="both", expand=True, padx=14, pady=6)
            tk.Label(box, text=titulo, bg=CARD, fg=TXT, font=FONT_B, anchor="w").pack(fill="x", padx=10, pady=(6, 2))
            t_ = ttk.Treeview(box, columns=[c[0] for c in cols], show="headings", height=8)
            for (k, tx), wd in zip(cols, larg):
                t_.heading(k, text=tx)
                t_.column(k, width=wd, anchor="w" if k in ("cliente", "hist") else "center", minwidth=50)
            t_.tag_configure("sem", foreground=RED)
            vs = ttk.Scrollbar(box, orient="vertical", command=t_.yview)
            hs = ttk.Scrollbar(box, orient="horizontal", command=t_.xview)
            t_.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
            vs.pack(side="right", fill="y")
            hs.pack(side="bottom", fill="x")
            t_.pack(fill="both", expand=True, padx=(10, 0), pady=(0, 6))
            for vals, tag in linhas:
                t_.insert("", "end", values=vals, tags=(tag,))

        l_t = [((r["aba"], r["linha"], r["lote"], r["status"], r["cliente"], r["nota"],
                 _brl(r["bruto"]) if pd.notna(r["bruto"]) else "", _brl(r["valor"])),
                "sem" if (r["aba"], r["linha"]) in sem_t else "") for r in tk_d.to_dict("records")]
        tabela(w, f"Relatório Ticket Log — {len(tk_d)} linha(s)",
               [("aba", "Aba"), ("linha", "Linha"), ("lote", "Lote"), ("status", "Status"),
                ("cliente", "Cliente"), ("nota", "Nota Fiscal"), ("bruto", "Valor Bruto"), ("valor", rot)],
               [50, 55, 100, 80, 330, 100, 110, 130], l_t, None)
        l_r = [((r["filial"], r["linha"], r["doc"], r["conta_partida"], r["historico"], _brl(r["valor"])),
                "sem" if (r["filial"], r["linha"]) in sem_r else "") for r in rz_d.to_dict("records")]
        tabela(w, f"Razão (crédito, contrapartida {_ROT_CP}) —{len(rz_d)} lançamento(s)",
               [("filial", "Filial"), ("linha", "Linha"), ("doc", "Doc."), ("cp", "Conta partida"),
                ("hist", "Histórico"), ("valor", "Crédito")],
               [60, 55, 110, 110, 560, 130], l_r, None)

    tv.bind("<Double-1>", abrir_dia)
    tk.Button(fr, text="🔍 Abrir lançamentos do dia (ou duplo clique)", command=abrir_dia, bg=CARD, fg=ACCENT,
              font=FONT_SM, relief="flat", activebackground=HOVER, cursor="hand2").pack(side="left", padx=14)

    def cartao(i, rot, val, cor):
        c = tk.Frame(resumo, bg=CARD, padx=14, pady=8)
        c.grid(row=0, column=i, padx=5, sticky="nsew")
        resumo.columnconfigure(i, weight=1)
        tk.Label(c, text=rot, bg=CARD, fg=TXT2, font=FONT_SM).pack()
        tk.Label(c, text=val, bg=CARD, fg=cor, font=FONT_B).pack()

    def processar():
        if not v_ticket.get() or not v_razao.get():
            messagebox.showwarning("Atenção", "Selecione os dois arquivos.")
            return
        barra.config(text="⏳ Processando…")
        win.update_idletasks()
        try:
            ticket, avisos = parse_ticket(v_ticket.get())
            razao = parse_razao(v_razao.get())
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            barra.config(text="Erro no processamento.")
            return
        st.update(ticket_full=ticket, razao_full=razao, avisos=avisos)
        anos = sorted({d.year for d in ticket["data"]} | {d.year for d in razao["data"]})
        cb_ano.config(values=["Todos"] + [str(a) for a in anos])
        v_ano.set("Todos")
        v_mes.set("Todos")
        recalcular()
        btn_exp.config(state="normal")
        if avisos:
            messagebox.showinfo("Avisos da leitura", "\n".join(avisos[:15]))

    def recalcular(_e=None):
        """Aplica o filtro de ano/mês nos dois lados e refaz a conciliação."""
        if st["ticket_full"] is None:
            return
        ano = int(v_ano.get()) if v_ano.get() != "Todos" else None
        mes = int(v_mes.get()[:2]) if v_mes.get() != "Todos" else None

        def filtra(df):
            m = pd.Series(True, index=df.index)
            if ano:
                m &= df["data"].map(lambda d: d.year == ano)
            if mes:
                m &= df["data"].map(lambda d: d.month == mes)
            return df[m]

        rot = v_valor.get()
        base = st["ticket_full"].assign(bruto=lambda d: d["valor"])
        if v_status.get() == "Somente Pago" and (base["status"] != "").any():
            base = base[base["status"].str.upper() == "PAGO"]
        if rot == "Valor de Reembolso":
            if base["reembolso"].notna().any():
                base = base[base["reembolso"].notna()].assign(valor=lambda d: d["reembolso"])
            else:
                rot = "Valor Bruto"
                v_valor.set(rot)
                messagebox.showinfo("Aviso", "Coluna 'Valor de Reembolso' não encontrada — usando Valor Bruto.")
        st["rotulo"] = rot
        tv.heading("ticket", text=f"Ticket Log ({rot.replace('Valor ', '').replace('de ', '')})")
        ticket, razao = filtra(base), filtra(st["razao_full"])
        dias, t_pend, r_pend = conciliar(ticket, razao)
        st.update(dias=dias, t_pend=t_pend, r_pend=r_pend, ticket=ticket, razao=razao)
        for w in resumo.winfo_children():
            w.destroy()
        tt, tr = float(ticket["valor"].sum()), float(razao["valor"].sum())
        n_div = sum(1 for x in dias if x["status"] != "OK")
        cartao(0, f"Total Ticket Log ({rot})", _brl(tt), ACCENT)
        cartao(1, f"Total Razão ({_ROT_CP})", _brl(tr), ACCENT)
        cartao(2, "Diferença", _brl(tt - tr), GREEN if abs(tt - tr) <= TOL else RED)
        cartao(3, "Dias divergentes", f"{n_div} / {len(dias)}", RED if n_div else GREEN)
        cartao(4, "Ticket sem Razão", f"{len(t_pend)} linhas · {_brl(float(t_pend['valor'].sum()))}",
               RED if len(t_pend) else GREEN)
        preencher()
        detalhe()
        per = " ".join(p for p in (v_mes.get() if mes else "", str(ano) if ano else "") if p) or "todo o período"
        barra.config(text=f"✅ {per}: {len(dias)} dias, {len(ticket)} linhas Ticket, {len(razao)} lançamentos Razão."
                          + (f"  ⚠ {len(st['avisos'])} aviso(s)." if st["avisos"] else ""))

    cb_valor.bind("<<ComboboxSelected>>", recalcular)
    cb_status.bind("<<ComboboxSelected>>", recalcular)
    cb_ano.bind("<<ComboboxSelected>>", recalcular)
    cb_mes.bind("<<ComboboxSelected>>", recalcular)

    def exportar():
        p = filedialog.asksaveasfilename(title="Salvar relatório", defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")], initialfile="relatorio_ticketlog.xlsx")
        if not p:
            return
        try:
            exportar_excel(st["dias"], st["t_pend"], st["r_pend"], st["ticket"], st["razao"], p,
                           st.get("rotulo", "Valor Bruto"))
            barra.config(text=f"✅ Relatório salvo em {p}")
            messagebox.showinfo("Exportado", f"Arquivo salvo com sucesso:\n{p}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    btn_proc.config(command=processar)
    btn_exp.config(command=exportar)
    return root


if __name__ == "__main__":
    import tkinter as tk
    _r = tk.Tk()
    _build_content(_r)
    _r.mainloop()
