import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from collections import defaultdict
import threading
import os

# ── Paleta de cores ──────────────────────────────────────────────────────────
from core import theme as _t     # identidade visual do Prisma (core/theme.py)

BG          = _t.BG_APP
SURFACE     = _t.CARD_BG
SURFACE2    = _t.SURFACE2
ACCENT      = _t.ACCENT_BLUE
ACCENT2     = _t.PRIMARY
TEXT        = _t.FG_TEXT
TEXT_MUTED  = _t.FG_MUTED
SUCCESS     = _t.ACCENT_GREEN
WARNING     = _t.ACCENT_ORANGE
DANGER      = _t.ACCENT_RED
ROW_ODD     = _t.CARD_BG
ROW_EVEN    = _t.SURFACE2
ROW_SEL     = _t.SEL_BG

FONT        = ("Segoe UI", 10)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)


# ── Lógica de negócio ─────────────────────────────────────────────────────────
def parse_balancete(ws):
    """Retorna dict  (emp, cod) -> saldo  do sheet Balancete."""
    data = {}
    skip_header = True
    for row in ws.iter_rows(values_only=True):
        if skip_header:
            if row[2] == "FORNECEDOR_EMP-FORNECEDOR_COD":
                skip_header = False
            continue
        chave_raw = row[2]
        saldo     = row[3]
        if chave_raw is None or saldo is None:
            continue
        try:
            partes = str(chave_raw).strip().split("-")
            if len(partes) < 2:
                continue
            emp = int(partes[0])
            cod = int(partes[1])
            # soma caso haja duplicatas no balancete
            data[(emp, cod)] = data.get((emp, cod), 0) + float(saldo)
        except (ValueError, TypeError):
            continue
    return data


def parse_financeiro(ws):
    """Retorna dict  (emp, cod) -> saldo_total  do sheet Financeiro (soma de parcelas)."""
    data = defaultdict(float)
    skip_header = True
    for row in ws.iter_rows(values_only=True):
        if skip_header:
            if row[1] == "EMPRESA_COD":
                skip_header = False
            continue
        try:
            emp   = int(row[2])   # FORNECEDOR_EMP
            cod   = int(row[3])   # FORNECEDOR_COD
            saldo = row[6]        # SALDO ATUAL
            if saldo is None:
                continue
            data[(emp, cod)] += float(saldo)
        except (TypeError, ValueError, IndexError):
            continue
    return data


def build_nome_map(ws):
    """Monta (emp,cod)->nome  a partir do Financeiro."""
    nomes = {}
    skip_header = True
    for row in ws.iter_rows(values_only=True):
        if skip_header:
            if row[1] == "EMPRESA_COD":
                skip_header = False
            continue
        try:
            emp  = int(row[2])
            cod  = int(row[3])
            nome = str(row[4]).strip() if row[4] else ""
            if (emp, cod) not in nomes and nome:
                nomes[(emp, cod)] = nome
        except (TypeError, ValueError, IndexError):
            continue
    return nomes


def comparar(bal, fin, nomes, tolerancia=0.05):
    """Retorna lista de dicts com divergências e matches."""
    resultados = []
    todas_chaves = set(bal) | set(fin)
    for (emp, cod) in sorted(todas_chaves):
        v_bal = bal.get((emp, cod))
        v_fin = fin.get((emp, cod))
        nome  = nomes.get((emp, cod), "")

        if v_bal is None:
            status = "Só no Financeiro"
            dif    = v_fin
        elif v_fin is None:
            status = "Só no Balancete"
            dif    = v_bal
        else:
            dif = v_bal - v_fin
            if abs(dif) <= tolerancia:
                status = "OK"
            else:
                status = "Divergente"

        resultados.append({
            "emp":    emp,
            "cod":    cod,
            "nome":   nome,
            "bal":    v_bal,
            "fin":    v_fin,
            "dif":    dif,
            "status": status,
        })
    return resultados


# ── Conciliação Débito x Crédito (por valor + Centro de Custo) ────────────────
def parse_centro_custo(valor):
    """Extrai (codigo, nome) de uma string 'COD - NOME' da coluna CENTRO DE CUSTO."""
    if valor is None:
        return None, ""
    s = str(valor).strip()
    partes = s.split("-", 1)
    if len(partes) == 2:
        cod  = partes[0].strip()
        nome = partes[1].strip()
    else:
        cod, nome = s, ""
    return cod, nome


def parse_lancamentos_conciliacao(ws):
    """
    Lê a planilha de lançamentos (colunas DATA, ..., HISTORICO,
    CENTRO DE CUSTO, CONTA PARTIDA, DEBITO, CREDITO) e retorna uma
    lista de dicts, um por lançamento (só linhas com débito ou crédito).
    """
    lancamentos = []
    header_ok = False
    for row in ws.iter_rows(values_only=True):
        if not header_ok:
            if row[0] == "DATA" and row[6] == "CENTRO DE CUSTO":
                header_ok = True
            continue
        if row[6] is None:
            continue
        try:
            data       = row[0]
            historico  = row[5]
            centro_raw = row[6]
            debito     = row[8]
            credito    = row[9]
        except IndexError:
            continue

        cod, nome = parse_centro_custo(centro_raw)
        try:
            debito  = float(debito)  if debito  not in (None, "") else 0.0
            credito = float(credito) if credito not in (None, "") else 0.0
        except (TypeError, ValueError):
            continue
        if debito == 0 and credito == 0:
            continue

        lancamentos.append({
            "data":        data,
            "historico":   historico,
            "centro_cod":  cod,
            "centro_nome": nome,
            "debito":      debito,
            "credito":     credito,
        })
    return lancamentos


def conciliar_debito_credito(lancamentos, tolerancia=0.01):
    """
    Concilia lançamentos de DÉBITO com os de CRÉDITO.

    Critérios:
      - VALOR: um débito só concilia com um crédito de mesmo valor
        (dentro da tolerância).
      - CENTRO DE CUSTO: o "casamento" só é feito dentro do mesmo
        código de centro de custo — é ele quem garante que o débito
        está sendo baixado contra o crédito do fornecedor correto
        (dois fornecedores podem ter o mesmo nome, mas nunca o mesmo
        código de centro de custo).

    Cada débito consome, no máximo, um crédito de mesmo valor no
    mesmo centro de custo (e vice-versa). Sobras de ambos os lados
    são consideradas NÃO conciliadas.

    Retorna (conciliados, nao_conciliados):
      - conciliados: lista de pares {centro_cod, centro_nome, valor,
        data_debito, historico_debito, data_credito, historico_credito}
      - nao_conciliados: lançamentos que sobraram, cada um com uma
        chave extra "tipo" ("DÉBITO" ou "CRÉDITO")
    """
    por_centro = defaultdict(lambda: {"deb": [], "cred": []})
    for lanc in lancamentos:
        grupo = por_centro[lanc["centro_cod"]]
        if lanc["debito"] > 0:
            grupo["deb"].append(lanc)
        if lanc["credito"] > 0:
            grupo["cred"].append(lanc)

    conciliados = []
    nao_conciliados = []

    for centro_cod, grupo in por_centro.items():
        debitos     = grupo["deb"]
        creditos    = grupo["cred"]
        cred_usado  = [False] * len(creditos)

        for deb in debitos:
            match_idx = None
            for i, cred in enumerate(creditos):
                if cred_usado[i]:
                    continue
                if abs(deb["debito"] - cred["credito"]) <= tolerancia:
                    match_idx = i
                    break
            if match_idx is not None:
                cred_usado[match_idx] = True
                cred = creditos[match_idx]
                conciliados.append({
                    "centro_cod":       centro_cod,
                    "centro_nome":      deb["centro_nome"] or cred["centro_nome"],
                    "valor":            deb["debito"],
                    "data_debito":      deb["data"],
                    "historico_debito": deb["historico"],
                    "data_credito":     cred["data"],
                    "historico_credito": cred["historico"],
                })
            else:
                nao_conciliados.append({**deb, "tipo": "DÉBITO"})

        for i, cred in enumerate(creditos):
            if not cred_usado[i]:
                nao_conciliados.append({**cred, "tipo": "CRÉDITO"})

    conciliados.sort(key=lambda r: (str(r["centro_cod"])))
    nao_conciliados.sort(key=lambda r: (str(r["centro_cod"]), r["tipo"]))
    return conciliados, nao_conciliados


def exportar_conciliacao_xlsx(conciliados, nao_conciliados, path):
    """Exporta o resultado da conciliação para um .xlsx com 2 abas:
    'Conciliados' e 'Não conciliados'."""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    header_fill = PatternFill("solid", fgColor="313147")
    header_font = Font(color="FFFFFF", bold=True)

    # ── Aba: Conciliados ──
    ws1 = wb.active
    ws1.title = "Conciliados"
    headers1 = ["Cód. CC", "Fornecedor (Centro de Custo)", "Valor (R$)",
                "Data Débito", "Histórico Débito",
                "Data Crédito", "Histórico Crédito"]
    ws1.append(headers1)
    for cell in ws1[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fill_ok = PatternFill("solid", fgColor="E1F5EC")
    for r in conciliados:
        ws1.append([r["centro_cod"], r["centro_nome"], r["valor"],
                   r["data_debito"], r["historico_debito"],
                   r["data_credito"], r["historico_credito"]])
        for cell in ws1[ws1.max_row]:
            cell.fill = fill_ok
        ws1.cell(row=ws1.max_row, column=3).number_format = "#,##0.00"

    widths1 = [10, 40, 14, 12, 45, 12, 45]
    for i, w in enumerate(widths1, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = "A2"

    # ── Aba: Não conciliados ──
    ws2 = wb.create_sheet("Não conciliados")
    headers2 = ["Tipo", "Data", "Cód. CC", "Fornecedor (Centro de Custo)",
                "Valor (R$)", "Histórico"]
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fill_deb  = PatternFill("solid", fgColor="FDE2E2")
    fill_cred = PatternFill("solid", fgColor="E1F5EC")
    for r in nao_conciliados:
        valor = r["debito"] if r["tipo"] == "DÉBITO" else r["credito"]
        ws2.append([r["tipo"], r["data"], r["centro_cod"], r["centro_nome"],
                   valor, r["historico"]])
        row_fill = fill_deb if r["tipo"] == "DÉBITO" else fill_cred
        for cell in ws2[ws2.max_row]:
            cell.fill = row_fill
        ws2.cell(row=ws2.max_row, column=5).number_format = "#,##0.00"

    widths2 = [10, 12, 10, 40, 14, 55]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    wb.save(path)


# ── Aplicativo ────────────────────────────────────────────────────────────────
class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master, bg=BG)
        self._resultados_completos = []
        self._filtro_status = "Todos"
        self._filepath = tk.StringVar()
        self._tolerancia = tk.DoubleVar(value=0.05)
        self._busca = tk.StringVar()
        self._busca.trace_add("write", lambda *_: self._aplicar_filtro())
        self._build_ui()

    # ── Construção da UI ──────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_topbar()
        self._build_summary()
        self._build_toolbar()
        self._build_table()
        self._build_statusbar()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=ACCENT2, height=56)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        tk.Label(bar, text="⚖", font=("Segoe UI", 18), bg=ACCENT2,
                 fg="white").pack(side="left", padx=(18, 6), pady=8)
        tk.Label(bar, text="Comparador  Balancete × Financeiro",
                 font=FONT_TITLE, bg=ACCENT2, fg="white").pack(side="left")

        # Arquivo
        frm = tk.Frame(bar, bg=ACCENT2)
        frm.pack(side="right", padx=14)
        tk.Entry(frm, textvariable=self._filepath, font=FONT_SMALL,
                 width=42, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left", padx=(0, 6))
        self._btn_abrir = tk.Button(frm, text="📂  Abrir XLSX",
                                    font=FONT_BOLD, bg=ACCENT, fg="white",
                                    activebackground=ACCENT2, activeforeground="white",
                                    relief="flat", bd=0, padx=12, pady=6,
                                    cursor="hand2", command=self._abrir_arquivo)
        self._btn_abrir.pack(side="left")

    def _build_summary(self):
        self._frm_summary = tk.Frame(self, bg=BG)
        self._frm_summary.pack(fill="x", padx=16, pady=(10, 0))
        self._cards = {}
        specs = [
            ("total",       "Total",          TEXT_MUTED, "📋"),
            ("ok",          "OK",             SUCCESS,    "✅"),
            ("divergente",  "Divergentes",    DANGER,     "⚠"),
            ("so_bal",      "Só Balancete",   WARNING,    "📊"),
            ("so_fin",      "Só Financeiro",  ACCENT,     "💰"),
        ]
        for key, label, cor, ico in specs:
            c = self._make_card(self._frm_summary, label, "—", cor, ico)
            c.pack(side="left", expand=True, fill="x", padx=5)
            self._cards[key] = c

    def _make_card(self, parent, label, valor, cor, ico):
        f = tk.Frame(parent, bg=SURFACE, bd=0, relief="flat",
                     highlightthickness=1, highlightbackground=SURFACE2)
        tk.Label(f, text=ico + "  " + label, font=FONT_SMALL,
                 bg=SURFACE, fg=TEXT_MUTED).pack(anchor="w", padx=12, pady=(8, 0))
        lbl = tk.Label(f, text=valor, font=("Segoe UI", 17, "bold"),
                       bg=SURFACE, fg=cor)
        lbl.pack(anchor="w", padx=12, pady=(0, 8))
        f._value_label = lbl
        return f

    def _update_cards(self, res):
        total     = len(res)
        ok        = sum(1 for r in res if r["status"] == "OK")
        div       = sum(1 for r in res if r["status"] == "Divergente")
        so_bal    = sum(1 for r in res if r["status"] == "Só no Balancete")
        so_fin    = sum(1 for r in res if r["status"] == "Só no Financeiro")
        for key, val in [("total", total), ("ok", ok), ("divergente", div),
                         ("so_bal", so_bal), ("so_fin", so_fin)]:
            self._cards[key]._value_label.config(text=str(val))

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=SURFACE, pady=6)
        bar.pack(fill="x", padx=16, pady=8)

        # Busca
        tk.Label(bar, text="🔍", font=("Segoe UI", 11), bg=SURFACE,
                 fg=TEXT_MUTED).pack(side="left", padx=(10, 2))
        tk.Entry(bar, textvariable=self._busca, font=FONT,
                 width=28, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left", padx=(0, 14))

        # Filtros de status
        tk.Label(bar, text="Filtrar:", font=FONT_SMALL, bg=SURFACE,
                 fg=TEXT_MUTED).pack(side="left", padx=(0, 6))
        self._btn_filtros = {}
        for label, cor in [("Todos", TEXT), ("OK", SUCCESS),
                            ("Divergente", DANGER), ("Só no Balancete", WARNING),
                            ("Só no Financeiro", ACCENT)]:
            b = tk.Button(bar, text=label, font=FONT_SMALL,
                          bg=SURFACE2, fg=cor, activebackground=ACCENT2,
                          activeforeground="white", relief="flat", bd=0,
                          padx=10, pady=4, cursor="hand2",
                          command=lambda l=label: self._set_filtro(l))
            b.pack(side="left", padx=2)
            self._btn_filtros[label] = b

        # Tolerância
        tk.Label(bar, text="  Tolerância (R$):", font=FONT_SMALL,
                 bg=SURFACE, fg=TEXT_MUTED).pack(side="left")
        tk.Entry(bar, textvariable=self._tolerancia, font=FONT_SMALL,
                 width=8, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left", padx=(4, 6))
        tk.Button(bar, text="Recalcular", font=FONT_SMALL,
                  bg=SURFACE2, fg=ACCENT, activebackground=ACCENT2,
                  activeforeground="white", relief="flat", bd=0,
                  padx=8, pady=4, cursor="hand2",
                  command=self._recalcular).pack(side="left")

        # Export
        tk.Button(bar, text="💾  Exportar CSV", font=FONT_SMALL,
                  bg=ACCENT2, fg="white", activebackground=ACCENT,
                  activeforeground="white", relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._exportar_csv).pack(side="right", padx=10)

        # Conciliação Débito x Crédito
        tk.Button(bar, text="🔗  Conciliar Déb/Créd", font=FONT_SMALL,
                  bg=SURFACE2, fg=WARNING, activebackground=ACCENT2,
                  activeforeground="white", relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._conciliar_deb_cred).pack(side="right", padx=4)

    def _build_table(self):
        frm = tk.Frame(self, bg=BG)
        frm.pack(fill="both", expand=True, padx=16, pady=(0, 4))

        cols = ("emp", "cod", "nome", "balancete", "financeiro", "diferenca", "status")
        hdrs = ("Emp", "Cód", "Nome do Fornecedor", "Balancete (R$)",
                "Financeiro (R$)", "Diferença (R$)", "Status")
        widths = (50, 60, 320, 120, 120, 120, 130)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Treeview",
                        background=ROW_ODD, foreground=TEXT,
                        fieldbackground=ROW_ODD, rowheight=26,
                        font=FONT_MONO, borderwidth=0)
        style.configure("Dark.Treeview.Heading",
                        background=SURFACE2, foreground=TEXT,
                        font=FONT_BOLD, relief="flat", borderwidth=0)
        style.map("Dark.Treeview",
                  background=[("selected", ROW_SEL)],
                  foreground=[("selected", "white")])

        self._tree = ttk.Treeview(frm, columns=cols, show="headings",
                                   style="Dark.Treeview", selectmode="browse")
        for col, hdr, w in zip(cols, hdrs, widths):
            self._tree.heading(col, text=hdr,
                               command=lambda c=col: self._sort_col(c))
            anchor = "w" if col == "nome" else "center"
            self._tree.column(col, width=w, anchor=anchor, stretch=(col == "nome"))

        vsb = ttk.Scrollbar(frm, orient="vertical",
                            command=self._tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal",
                            command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set,
                             xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        # Tags de cor por status
        self._tree.tag_configure("OK",              foreground=SUCCESS)
        self._tree.tag_configure("Divergente",      foreground=DANGER)
        self._tree.tag_configure("SoBal",           foreground=WARNING)
        self._tree.tag_configure("SoFin",           foreground=ACCENT)
        self._tree.tag_configure("odd",             background=ROW_ODD)
        self._tree.tag_configure("even",            background=ROW_EVEN)

        self._tree.bind("<Double-1>", self._on_double_click)

        # Detalhe popup (tooltip)
        self._tooltip = tk.Toplevel(self)
        self._tooltip.withdraw()
        self._tooltip.overrideredirect(True)
        self._tooltip.configure(bg=SURFACE2)
        self._tip_lbl = tk.Label(self._tooltip, bg=SURFACE2, fg=TEXT,
                                  font=FONT_SMALL, justify="left", padx=12, pady=8)
        self._tip_lbl.pack()

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=SURFACE2, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self._status_var = tk.StringVar(value="Abra um arquivo XLSX para começar.")
        tk.Label(bar, textvariable=self._status_var, font=FONT_SMALL,
                 bg=SURFACE2, fg=TEXT_MUTED, anchor="w").pack(
            side="left", padx=10, pady=4)
        self._progress = ttk.Progressbar(bar, mode="indeterminate", length=150)
        self._progress.pack(side="right", padx=10, pady=4)

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _abrir_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar arquivo XLSX",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not path:
            return
        self._filepath.set(path)
        self._carregar(path)

    def _carregar(self, path):
        self._status("Carregando arquivo…")
        self._progress.start(12)
        self._btn_abrir.config(state="disabled")
        threading.Thread(target=self._worker_carregar,
                         args=(path,), daemon=True).start()

    def _worker_carregar(self, path):
        try:
            wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
            bal  = parse_balancete(wb["Balancete"])
            fin  = parse_financeiro(wb["Financeiro"])

            wb2  = openpyxl.load_workbook(path, read_only=True, data_only=True)
            nomes = build_nome_map(wb2["Financeiro"])

            res = comparar(bal, fin, nomes, self._tolerancia.get())
            self.after(0, self._carregar_ok, res)
        except Exception as e:
            self.after(0, self._carregar_erro, str(e))

    def _carregar_ok(self, res):
        self._resultados_completos = res
        self._progress.stop()
        self._btn_abrir.config(state="normal")
        self._aplicar_filtro()
        self._status(f"Arquivo carregado — {len(res)} fornecedores analisados.")

    def _carregar_erro(self, msg):
        self._progress.stop()
        self._btn_abrir.config(state="normal")
        self._status(f"Erro: {msg}")
        messagebox.showerror("Erro ao carregar", msg)

    def _recalcular(self):
        path = self._filepath.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Atenção", "Nenhum arquivo carregado.")
            return
        self._carregar(path)

    def _conciliar_deb_cred(self):
        path = self._filepath.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Atenção", "Nenhum arquivo carregado.")
            return
        self._status("Conciliando débitos e créditos…")
        self._progress.start(12)
        self._btn_abrir.config(state="disabled")
        threading.Thread(target=self._worker_conciliar,
                         args=(path,), daemon=True).start()

    def _worker_conciliar(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            lancamentos = parse_lancamentos_conciliacao(ws)
            conciliados, nao_conciliados = conciliar_debito_credito(
                lancamentos, self._tolerancia.get())
            self.after(0, self._conciliar_ok, conciliados, nao_conciliados,
                      len(lancamentos))
        except Exception as e:
            self.after(0, self._carregar_erro, str(e))

    def _conciliar_ok(self, conciliados, nao_conciliados, total_lancamentos):
        self._progress.stop()
        self._btn_abrir.config(state="normal")
        self._status(f"Conciliação concluída — {len(conciliados)} par(es) "
                     f"conciliado(s), {len(nao_conciliados)} de "
                     f"{total_lancamentos} lançamentos não conciliados.")
        self._abrir_janela_conciliacao(conciliados, nao_conciliados)

    def _abrir_janela_conciliacao(self, conciliados, nao_conciliados):
        win = tk.Toplevel(self)
        win.title("Conciliação Débito x Crédito")
        win.geometry("1150x650")
        win.configure(bg=BG)

        tk.Label(win, text=f"✅  {len(conciliados)} par(es) conciliado(s)   "
                 f"⚠  {len(nao_conciliados)} lançamento(s) não conciliado(s)",
                 font=FONT_BOLD, bg=BG, fg=TEXT).pack(anchor="w",
                 padx=14, pady=(12, 6))

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=14, pady=(0, 6))

        aba_ok  = tk.Frame(nb, bg=BG)
        aba_nok = tk.Frame(nb, bg=BG)
        nb.add(aba_ok,  text=f"Conciliados ({len(conciliados)})")
        nb.add(aba_nok, text=f"Não conciliados ({len(nao_conciliados)})")

        fmt = lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # ── Aba Conciliados ──
        frm_ok = tk.Frame(aba_ok, bg=BG)
        frm_ok.pack(fill="both", expand=True, padx=4, pady=4)

        cols_ok = ("centro_cod", "centro_nome", "valor",
                  "data_debito", "historico_debito",
                  "data_credito", "historico_credito")
        hdrs_ok = ("Cód. CC", "Fornecedor", "Valor (R$)",
                  "Data Déb.", "Histórico Débito",
                  "Data Créd.", "Histórico Crédito")
        widths_ok = (60, 220, 100, 80, 260, 80, 260)

        tree_ok = ttk.Treeview(frm_ok, columns=cols_ok, show="headings",
                               style="Dark.Treeview", selectmode="browse")
        for col, hdr, w in zip(cols_ok, hdrs_ok, widths_ok):
            tree_ok.heading(col, text=hdr)
            anchor = "w" if col in ("centro_nome", "historico_debito", "historico_credito") else "center"
            tree_ok.column(col, width=w, anchor=anchor,
                          stretch=(col in ("centro_nome", "historico_debito", "historico_credito")))

        vsb_ok = ttk.Scrollbar(frm_ok, orient="vertical", command=tree_ok.yview)
        hsb_ok = ttk.Scrollbar(frm_ok, orient="horizontal", command=tree_ok.xview)
        tree_ok.configure(yscrollcommand=vsb_ok.set, xscrollcommand=hsb_ok.set)
        tree_ok.grid(row=0, column=0, sticky="nsew")
        vsb_ok.grid(row=0, column=1, sticky="ns")
        hsb_ok.grid(row=1, column=0, sticky="ew")
        frm_ok.rowconfigure(0, weight=1)
        frm_ok.columnconfigure(0, weight=1)

        tree_ok.tag_configure("odd",  background=ROW_ODD)
        tree_ok.tag_configure("even", background=ROW_EVEN)

        for i, r in enumerate(conciliados):
            tag_linha = "even" if i % 2 == 0 else "odd"
            tree_ok.insert("", "end", values=(
                r["centro_cod"], r["centro_nome"], fmt(r["valor"]),
                r["data_debito"], r["historico_debito"],
                r["data_credito"], r["historico_credito"]), tags=(tag_linha,))

        # ── Aba Não conciliados ──
        frm = tk.Frame(aba_nok, bg=BG)
        frm.pack(fill="both", expand=True, padx=4, pady=4)

        cols = ("tipo", "data", "centro_cod", "centro_nome", "valor", "historico")
        hdrs = ("Tipo", "Data", "Cód. CC", "Fornecedor (Centro de Custo)",
                "Valor (R$)", "Histórico")
        widths = (70, 90, 70, 300, 110, 420)

        tree = ttk.Treeview(frm, columns=cols, show="headings",
                            style="Dark.Treeview", selectmode="browse")
        for col, hdr, w in zip(cols, hdrs, widths):
            tree.heading(col, text=hdr)
            anchor = "w" if col in ("centro_nome", "historico") else "center"
            tree.column(col, width=w, anchor=anchor,
                       stretch=(col in ("centro_nome", "historico")))

        vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        tree.tag_configure("DÉBITO",  foreground=DANGER)
        tree.tag_configure("CRÉDITO", foreground=SUCCESS)
        tree.tag_configure("odd",  background=ROW_ODD)
        tree.tag_configure("even", background=ROW_EVEN)

        for i, r in enumerate(nao_conciliados):
            valor = r["debito"] if r["tipo"] == "DÉBITO" else r["credito"]
            tag_linha = "even" if i % 2 == 0 else "odd"
            tree.insert("", "end",
                values=(r["tipo"], r["data"], r["centro_cod"], r["centro_nome"],
                        fmt(valor), r["historico"]),
                tags=(tag_linha, r["tipo"]))

        def _exportar():
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                title="Salvar conciliação")
            if not path:
                return
            exportar_conciliacao_xlsx(conciliados, nao_conciliados, path)
            messagebox.showinfo("Exportado", f"Arquivo salvo em:\n{path}")

        tk.Button(win, text="💾  Exportar XLSX (2 abas)", font=FONT_SMALL,
                  bg=ACCENT2, fg="white", activebackground=ACCENT,
                  activeforeground="white", relief="flat", bd=0,
                  padx=10, pady=6, cursor="hand2",
                  command=_exportar).pack(anchor="e", padx=14, pady=(0, 12))

    def _set_filtro(self, status):
        self._filtro_status = status
        # Highlight botão ativo
        for lbl, btn in self._btn_filtros.items():
            btn.config(relief="sunken" if lbl == status else "flat",
                       bg=ACCENT if lbl == status else SURFACE2)
        self._aplicar_filtro()

    def _aplicar_filtro(self):
        busca  = self._busca.get().lower()
        status = self._filtro_status

        filtrados = [
            r for r in self._resultados_completos
            if (status == "Todos" or r["status"] == status)
            and (not busca or busca in str(r["cod"]).lower()
                 or busca in r["nome"].lower()
                 or busca in str(r["emp"]))
        ]
        self._preencher_tabela(filtrados)
        self._update_cards(self._resultados_completos)
        self._status(f"Exibindo {len(filtrados)} de "
                     f"{len(self._resultados_completos)} fornecedores.")

    def _preencher_tabela(self, dados):
        self._tree.delete(*self._tree.get_children())
        fmt = lambda v: f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if v is not None else "—"

        for i, r in enumerate(dados):
            tag_linha = "even" if i % 2 == 0 else "odd"
            tag_status = {
                "OK":                 "OK",
                "Divergente":         "Divergente",
                "Só no Balancete":    "SoBal",
                "Só no Financeiro":   "SoFin",
            }.get(r["status"], "")

            self._tree.insert("", "end",
                values=(
                    r["emp"],
                    r["cod"],
                    r["nome"],
                    fmt(r["bal"]),
                    fmt(r["fin"]),
                    fmt(r["dif"]),
                    r["status"],
                ),
                tags=(tag_linha, tag_status))

    def _exportar_csv(self):
        if not self._resultados_completos:
            messagebox.showwarning("Atenção", "Nenhum dado para exportar.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Salvar CSV")
        if not path:
            return
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=[
                "emp","cod","nome","bal","fin","dif","status"])
            w.writeheader()
            w.writerows(self._resultados_completos)
        messagebox.showinfo("Exportado", f"Arquivo salvo em:\n{path}")

    def _sort_col(self, col):
        """Ordena tabela pela coluna clicada."""
        items = [(self._tree.set(iid, col), iid)
                 for iid in self._tree.get_children()]
        try:
            items.sort(key=lambda x: float(
                x[0].replace(".", "").replace(",", ".").replace("—", "0")))
        except ValueError:
            items.sort(key=lambda x: x[0].lower())
        for idx, (_, iid) in enumerate(items):
            self._tree.move(iid, "", idx)
            tag = "even" if idx % 2 == 0 else "odd"
            tags = [t for t in self._tree.item(iid, "tags")
                    if t not in ("even", "odd")] + [tag]
            self._tree.item(iid, tags=tags)

    def _on_double_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        vals = self._tree.item(iid, "values")
        if not vals:
            return
        emp, cod, nome, bal, fin, dif, status = vals
        msg = (f"Empresa:     {emp}\n"
               f"Código:      {cod}\n"
               f"Nome:        {nome}\n\n"
               f"Balancete:   R$ {bal}\n"
               f"Financeiro:  R$ {fin}\n"
               f"Diferença:   R$ {dif}\n\n"
               f"Status:      {status}")
        messagebox.showinfo(f"Fornecedor {emp}-{cod}", msg)

    def _status(self, msg):
        self._status_var.set(msg)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Comparador Balancete × Financeiro")
    root.geometry("1180x720")
    root.minsize(900, 560)
    App(root).pack(fill="both", expand=True)
    root.mainloop()
