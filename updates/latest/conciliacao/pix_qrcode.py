import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import os, re
from datetime import datetime
from core.theme import T, aplicar_estilos_ttk, aplicar_tags_tree, botao_tema, registrar_callback

def _cores():

    global COR_BG, COR_PAINEL, COR_BORDA, COR_ACENTO, COR_ACENTO2
    global COR_TEXTO, COR_TEXTO_SEC, COR_VERDE, COR_AMARELO
    global COR_VERMELHO, COR_CINZA, COR_AZUL, COR_LARANJA
    COR_BG        = T("BG")
    COR_PAINEL    = T("PAINEL")
    COR_BORDA     = T("BORDA")
    COR_ACENTO    = T("ACENTO")
    COR_ACENTO2   = T("ACENTO2")
    COR_TEXTO     = T("TEXTO")
    COR_TEXTO_SEC = T("TEXTO_SEC")
    COR_VERDE     = T("VERDE")
    COR_AMARELO   = T("AMARELO")
    COR_VERMELHO  = T("VERMELHO")
    COR_CINZA     = T("CINZA")
    COR_AZUL      = T("AZUL")
    COR_LARANJA   = T("LARANJA")

_cores()

def _num(val):

    if val is None:
        return 0.0
    s = str(val).strip()

    if re.search(r"\d\.\d{3},\d", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s)
    except:
        return 0.0


def _extrair_linhas_pdf(path):

    linhas = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            texto = page.extract_text(layout=True) or ""
            for linha in texto.splitlines():
                linhas.append(linha)
    return linhas


def _ultimo_num_linha(linha):

    tokens = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2}|\d+", linha)
    for tok in reversed(tokens):
        n = _num(tok)
        if n > 0:
            return n
    return 0.0

def ler_cupom_fiscal(path):

    linhas = _extrair_linhas_pdf(path)
    registros = []

    cupom_atual  = None
    cond_atual   = ""
    vender_atual = ""
    data_atual   = ""

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue
        up = linha_strip.upper()

        # ── Data do dia (ex: "DATA : 05/06/2026") ────────────────────────────
        m_data = re.search(r"DATA\s*:\s*(\d{2}/\d{2}/\d{4})", up)
        if m_data:
            data_atual = m_data.group(1)
            continue

        m = re.match(r"^(\d{5,})\s+\d{5,}\s+\d+\s+\w+\s+(\S+)\s+(\S+)", linha_strip)
        if m:
            cupom_atual  = m.group(1)
            cond_atual   = m.group(2)
            vender_atual = m.group(3)
            continue

        if any(t in up for t in [
            "TOTAL GERAL", "TOTAL :", "TOTAL:",
            "DESCRICAO", "TOTAIS", "CANCELADOS",
            "SERVICOS", "VENDAS"
        ]):
            cupom_atual = None
            continue

        # ✅ headers de página NÃO zeram o cupom
        if "RELATORIO CUPOM FISCAL" in up:
            continue
        if "FILIAL" in up and "PERIODO" in up:
            continue
        if "EMISSAO" in up:
            continue

        if "VENDA PIX" in up and "MAQUINETA" not in up:

            if not cupom_atual:
                continue
            valor = _ultimo_num_linha(linha_strip)
            if valor > 0:
                registros.append({
                    "origem":     "Cupom Fiscal",
                    "referencia": f"Cupom {cupom_atual}",
                    "data":       data_atual,
                    "valor":      round(valor, 2),
                    "descricao":  (f"VENDA PIX | Cupom {cupom_atual} "
                                   f"| Cond: {cond_atual} | Vend: {vender_atual}"),
                    "status":     "pendente",
                    "par_banco":  "",
                })

                cupom_atual = None

    return pd.DataFrame(registros)

def ler_nota_fiscal(path):

    linhas = _extrair_linhas_pdf(path)
    registros = []

    nota_atual    = None
    cliente_atual = ""
    cond_atual    = ""
    data_atual    = ""

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue
        up = linha_strip.upper()

        # ── Data do dia (ex: "09/06/2026" isolada como cabeçalho de seção) ───
        m_data = re.match(r"^(\d{2}/\d{2}/\d{4})$", linha_strip.strip())
        if m_data:
            data_atual = m_data.group(1)
            continue

        m = re.match(r"^(\d{5,})\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(.+?)\s{2,}(\S+)\s+([\d\.,]+)", linha_strip)
        if m:
            nota_atual    = m.group(1)
            cliente_atual = m.group(6).strip()
            cond_atual    = m.group(7).strip()
            continue

        if re.match(r"^\d{5,}\s", linha_strip):
            partes = linha_strip.split()
            nums_ini = sum(1 for p in partes[:5] if re.match(r"^\d+$", p))
            if nums_ini >= 3:
                nota_atual    = partes[0]
                cliente_atual = " ".join(p for p in partes[4:10]
                                         if not re.match(r"^[\d\.,]+$", p))
                cond_atual    = ""
            continue

        if "VENDA PIX" in up and "MAQUINETA" not in up:
            valor = _ultimo_num_linha(linha_strip)
            if valor > 0 and nota_atual:
                registros.append({
                    "origem":     "Nota Fiscal",
                    "referencia": f"NF {nota_atual}",
                    "data":       data_atual,
                    "valor":      round(valor, 2),
                    "descricao":  (f"VENDA PIX | NF {nota_atual} "
                                   f"| {cliente_atual} | Cond: {cond_atual}"),
                    "status":     "pendente",
                    "par_banco":  "",
                })

    return pd.DataFrame(registros)

def ler_recibos(path):

    linhas = _extrair_linhas_pdf(path)

    registros = []
    recibo_atual = None
    data_atual   = ""

    for linha in linhas:
        linha_limpa = linha.strip()
        up = linha_limpa.upper()

        # ── Data do dia (ex: "DATA EMISSAO : 05/06/2026") ────────────────────
        m_data = re.search(r"DATA\s+EMISSAO\s*:\s*(\d{2}/\d{2}/\d{4})", up)
        if m_data:
            data_atual = m_data.group(1)
            continue

        m_recibo = re.match(r"^(\d+)\s+(\d+)\s+", linha_limpa)
        tem_texto_apos = bool(re.search(r"[A-Za-z]", linha_limpa.split(None, 2)[-1])) \
                        if m_recibo else False
        if m_recibo and tem_texto_apos:
            if recibo_atual and recibo_atual["valor"] > 0:
                registros.append(recibo_atual)
            recibo_atual = None

            if "DEP. PIX QRCOD" in up:
                partes = linha_limpa.split()
                numero_recibo = partes[1]

                total_cabecalho = _ultimo_num_linha(linha_limpa)

                recibo_atual = {
                    "origem":     "Recibo",
                    "referencia": f"Recibo {numero_recibo}",
                    "data":       data_atual,
                    "valor":      0.0,
                    "descricao":  f"RECIBO PIX QRCOD | Recibo {numero_recibo}",
                    "status":     "pendente",
                    "par_banco":  "",
                }
            continue

        if recibo_atual and (up.startswith("DPP") or up.startswith("ANT")):
            nums = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", linha_limpa)
            INDICE_DEPOSITO = 9
            if len(nums) > INDICE_DEPOSITO:
                deposito = _num(nums[INDICE_DEPOSITO])
                recibo_atual["valor"] = round(recibo_atual["valor"] + deposito, 2)

    if recibo_atual and recibo_atual["valor"] > 0:
        registros.append(recibo_atual)

    df = pd.DataFrame(registros)
    if not df.empty:
        df["saldo_rest"] = df["valor"]
    return df

def ler_mov_pix(path):

    registros = []

    COLUNAS = ["FILIAL", "DT_RECEB", "HR_RECEB", "VENDEDOR", "PEDIDO", "TXID",
               "DT_ENVIO", "HR_ENVIO", "VALOR"]

    SEQUENCIAS = [
        (["DT", "RECEB."], "DT_RECEB"),
        (["HR", "RECEB."], "HR_RECEB"),
        (["DT", "ENVIO"],  "DT_ENVIO"),
        (["HR", "ENVIO"],  "HR_ENVIO"),
        (["FILIAL"],       "FILIAL"),
        (["VENDEDOR"],     "VENDEDOR"),
        (["PEDIDO"],       "PEDIDO"),
        (["TXID"],         "TXID"),
        (["VALOR"],        "VALOR"),
    ]

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            linhas_words = {}
            for w in words:
                y = round(float(w["top"]), 0)
                linhas_words.setdefault(y, []).append(w)

            col_x = {}   
            header_y = None

            for y in sorted(linhas_words):
                textos = " ".join(w["text"].upper() for w in linhas_words[y])
                if "TXID" in textos and "VALOR" in textos:
                    header_y = y
                    ws_linha = sorted(linhas_words[y], key=lambda w: float(w["x0"]))
                    tokens = [w["text"].upper() for w in ws_linha]
                    for seq, col_nome in SEQUENCIAS:
                        for i in range(len(tokens) - len(seq) + 1):
                            if tokens[i:i+len(seq)] == seq:
                                if col_nome not in col_x:
                                    col_x[col_nome] = float(ws_linha[i]["x0"])
                                break
                    break

            if not col_x or header_y is None:
                texto = page.extract_text(layout=True) or ""
                _parsear_mov_pix_texto(texto, registros)
                continue

            cols_ordenadas = sorted(col_x.items(), key=lambda kv: kv[1])
            faixas = {}
            for i, (col, x) in enumerate(cols_ordenadas):
                x_fim = cols_ordenadas[i+1][1] if i+1 < len(cols_ordenadas) else 9999
                faixas[col] = (x - 5, x_fim)

            for y in sorted(linhas_words):
                if y <= header_y:
                    continue
                ws = linhas_words[y]
                texto_linha = " ".join(w["text"] for w in ws).upper()

                if any(t in texto_linha for t in ["TOTAL", "PAGINA", "PÁGINA",
                                                   "EMPRESA", "FILIAL :", "PERIODO",
                                                   "RELATORIO", "EMISSAO"]):
                    continue

                reg = {c: "" for c in COLUNAS}
                for w in ws:
                    wx = float(w["x0"])
                    for col, (x_ini, x_fim) in faixas.items():
                        if x_ini <= wx < x_fim:
                            reg[col] = (reg[col] + " " + w["text"]).strip()
                            break

                valor  = _num(reg.get("VALOR", ""))
                txid   = reg.get("TXID",   "").strip()
                pedido = reg.get("PEDIDO", "").strip()
                if valor > 0 and (txid or pedido):
                    registros.append({
                        "FILIAL":    reg.get("FILIAL",   ""),
                        "DT_RECEB":  reg.get("DT_RECEB", ""),
                        "HR_RECEB":  reg.get("HR_RECEB", ""),
                        "VENDEDOR":  reg.get("VENDEDOR", ""),
                        "PEDIDO":    pedido,
                        "TXID":      txid,
                        "VALOR":     valor,
                        "status":    "pendente",
                        "par_venda": "",
                        "saldo_rest": valor,
                    })

    COLS = ["FILIAL","DT_RECEB","HR_RECEB","VENDEDOR","PEDIDO",
            "TXID","VALOR","status","par_venda","saldo_rest"]
    if not registros:
        df = pd.DataFrame(columns=COLS)
        df["VALOR"]      = pd.Series(dtype=float)
        df["saldo_rest"] = pd.Series(dtype=float)
    else:
        df = pd.DataFrame(registros, columns=COLS)
        df["saldo_rest"] = df["VALOR"]
    return df


def _parsear_mov_pix_texto(texto, registros):
    """
    Fallback: parseia o relatório de mov. PIX como texto puro linha a linha.
    Detecta linhas de dados pelo padrão: FILIAL  DATA  HORA  ...  TXID  ...  VALOR
    """
    cabecalho_encontrado = False
    for linha in texto.splitlines():
        up = linha.strip().upper()
        if not up:
            continue
        if "TXID" in up and "VALOR" in up:
            cabecalho_encontrado = True
            continue
        if not cabecalho_encontrado:
            continue
        if any(t in up for t in ["TOTAL", "PAGINA", "PÁGINA", "EMPRESA",
                                  "PERIODO", "RELATORIO", "EMISSAO"]):
            continue

        m = re.match(
            r"^\s*(\d{2})\s+"
            r"(\d{2}/\d{2}/\d{4})\s+"
            r"(\S+)\s+"
            r"(\S+)\s+"
            r"(\S+)\s+"
            r"(\S+)\s+"
            r"\S+\s+"
            r"\S+\s+"
            r"([\d\.,]+)\s*$",
            linha.strip())
        if m:
            valor = _num(m.group(7))
            if valor > 0:
                registros.append({
                    "FILIAL":    m.group(1),
                    "DT_RECEB":  m.group(2),
                    "HR_RECEB":  m.group(3),
                    "VENDEDOR":  m.group(4),
                    "PEDIDO":    m.group(5),
                    "TXID":      m.group(6),
                    "VALOR":     valor,
                    "status":    "pendente",
                    "par_venda": "",
                    "saldo_rest": valor,
                })
            continue

        partes = linha.strip().split()
        if (len(partes) >= 6
                and re.match(r"^\d{2}$", partes[0])
                and len(partes) > 1
                and re.match(r"^\d{2}/\d{2}/\d{4}$", partes[1])):
            valor = _num(partes[-1])
            if valor > 0:
                registros.append({
                    "FILIAL":    partes[0] if len(partes) > 0 else "",
                    "DT_RECEB":  partes[1] if len(partes) > 1 else "",
                    "HR_RECEB":  partes[2] if len(partes) > 2 else "",
                    "VENDEDOR":  partes[3] if len(partes) > 3 else "",
                    "PEDIDO":    partes[4] if len(partes) > 4 else "",
                    "TXID":      partes[5] if len(partes) > 5 else "",
                    "VALOR":     valor,
                    "status":    "pendente",
                    "par_venda": "",
                    "saldo_rest": valor,
                })



def conciliar_automatico(df_vendas, df_banco, tolerancia=0.01):
    """
    Casa cada linha de venda com linha(s) do banco pelo valor.
    Suporta conciliações parciais (um valor vendas → vários pix ou vice-versa).
    """
    dv = df_vendas.copy()
    db = df_banco.copy()

    dv["status"]   = "pendente"
    dv["par_banco"] = ""
    dv["saldo_rest"] = dv["valor"]

    db["status"]    = "pendente"
    db["par_venda"] = ""
    db["saldo_rest"] = db["VALOR"]

    par_counter = [0]

    def novo_par():
        par_counter[0] += 1
        return f"P{par_counter[0]:04d}"

    for iv, row_v in dv.iterrows():
        saldo_v = dv.at[iv, "saldo_rest"]
        if saldo_v <= tolerancia:
            continue

        candidatos = db[
            (db["saldo_rest"] > tolerancia) &
            (abs(db["VALOR"] - saldo_v) <= tolerancia)
        ]

        for ib, row_b in candidatos.iterrows():
            saldo_v = dv.at[iv, "saldo_rest"]
            saldo_b = db.at[ib, "saldo_rest"]
            if saldo_v <= tolerancia or saldo_b <= tolerancia:
                continue
            valor_match = min(saldo_v, saldo_b)
            par = novo_par()
            dv.at[iv, "par_banco"]  += ("," if dv.at[iv,"par_banco"] else "") + par
            dv.at[iv, "saldo_rest"]  = round(saldo_v - valor_match, 2)
            db.at[ib, "par_venda"]  += ("," if db.at[ib,"par_venda"] else "") + par
            db.at[ib, "saldo_rest"]  = round(saldo_b - valor_match, 2)

    def status_venda(row):
        if not row["par_banco"]:
            return "pendente"
        return "conciliado" if row["saldo_rest"] <= tolerancia else "parcial"

    def status_banco(row):
        if not row["par_venda"]:
            return "pendente"
        return "conciliado" if row["saldo_rest"] <= tolerancia else "parcial"

    dv["status"] = dv.apply(status_venda, axis=1)
    db["status"] = db.apply(status_banco, axis=1)

    return dv, db

class Conciliacao_Pix(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Conciliador PIX QRCODE — PMZ Peças e Pneus")
        self.geometry("1500x860")
        self.configure(bg=COR_BG)
        self.resizable(True, True)

        self.df_vendas  = None
        self.df_banco   = None
        self.paths      = {"cupom": None, "nf": None, "recibo": None, "banco": None}

        self.sel_vendas = []
        self.sel_bancos = []

        self._build_ui()
        self._aplicar_estilos()


    def _build_ui(self):
        self._build_topbar()
        corpo = tk.Frame(self, bg=COR_BG)
        corpo.pack(fill="both", expand=True, padx=10, pady=(0,10))
        self._build_painel_esq(corpo)
        self._build_tabelas(corpo)
        self._build_statusbar()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=COR_PAINEL, height=58)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        self._topbar = bar

        tk.Label(bar, text="🔵  Conciliador PIX QRCODE — PMZ",
                 bg=COR_PAINEL, fg=COR_TEXTO,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16, pady=12)

        btns = [
            ("🔄  Conciliar Auto",   COR_AZUL,    self.conciliar_auto),
            ("🤝  Conciliar Manual", COR_VERDE,   self.conciliar_manual),
            ("🔓  Desconciliar",     COR_AMARELO, self.desconciliar),
            ("🚫  Ignorar",          COR_CINZA,   self.ignorar),
        ]
        self._action_btns = []
        for txt, cor, cmd in btns:
            b = tk.Button(bar, text=txt, bg=cor, fg="white",
                          font=("Segoe UI", 9, "bold"), relief="flat",
                          padx=12, pady=6, cursor="hand2",
                          command=cmd)
            b.pack(side="left", padx=4, pady=12)
            self._action_btns.append((b, cor))

        # ── Filtros inline na topbar ──────────────────────────────────────────
        sep = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep.pack(side="left", fill="y", pady=10, padx=6)

        # Status
        tk.Label(bar, text="Status:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8)).pack(side="left", padx=(4, 2))
        self.filtro_status = ttk.Combobox(bar, state="readonly", width=11,
            values=["Todos", "pendente", "parcial", "conciliado", "ignorado"])
        self.filtro_status.set("Todos")
        self.filtro_status.pack(side="left", pady=12)
        self.filtro_status.bind("<<ComboboxSelected>>", lambda _: self.atualizar_tabelas())

        # Data De/Até
        sep2 = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep2.pack(side="left", fill="y", pady=10, padx=6)

        tk.Label(bar, text="De:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8)).pack(side="left", padx=(2, 2))
        self.filtro_data_ini = tk.Entry(bar, width=10, font=("Segoe UI", 8))
        self.filtro_data_ini.pack(side="left", pady=12)

        tk.Label(bar, text="Até:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8)).pack(side="left", padx=(6, 2))
        self.filtro_data_fim = tk.Entry(bar, width=10, font=("Segoe UI", 8))
        self.filtro_data_fim.pack(side="left", pady=12)

        tk.Button(bar, text="🔍", bg=COR_ACENTO2, fg="white",
                  font=("Segoe UI", 9), relief="flat",
                  padx=6, pady=4, cursor="hand2",
                  command=self.atualizar_tabelas).pack(side="left", padx=(4, 0), pady=12)

        tk.Button(bar, text="✖", bg=COR_BG, fg=COR_TEXTO_SEC,
                  font=("Segoe UI", 9), relief="flat",
                  padx=4, pady=4, cursor="hand2",
                  command=self._limpar_filtro_data).pack(side="left", padx=(2, 0), pady=12)

        self._btn_tema = botao_tema(bar, callback=self._aplicar_tema)
        self._btn_tema.pack(side="right", padx=12, pady=12)

    def _build_painel_esq(self, parent):
        frame = tk.Frame(parent, bg=COR_PAINEL, width=230)
        frame.pack(side="left", fill="y", padx=(0,10), pady=10)
        frame.pack_propagate(False)

        tk.Label(frame, text="ARQUIVOS", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(14,4))

        self.lbl_paths = {}
        arquivos = [
            ("cupom",  "📄 Cupom Fiscal",  self.carregar_cupom),
            ("nf",     "🧾 Nota Fiscal",   self.carregar_nf),
            ("recibo", "📋 Recibos",       self.carregar_recibo),
            ("banco",  "🏦 Mov. PIX Banco",self.carregar_banco),
        ]
        for chave, label, cmd in arquivos:
            tk.Button(frame, text=label, bg=COR_ACENTO2, fg="white",
                      font=("Segoe UI", 8, "bold"), relief="flat",
                      padx=8, pady=4, cursor="hand2", anchor="w",
                      command=cmd).pack(fill="x", padx=12, pady=(4,0))
            lbl = tk.Label(frame, text="(não carregado)", bg=COR_PAINEL,
                           fg=COR_TEXTO_SEC, font=("Segoe UI", 7),
                           wraplength=200, justify="left")
            lbl.pack(anchor="w", padx=14, pady=(0,4))
            self.lbl_paths[chave] = lbl

        sep = tk.Frame(frame, bg=COR_BORDA, height=1)
        sep.pack(fill="x", padx=12, pady=10)
        tk.Label(frame, text="RESUMO VENDAS PIX", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)

        self.lbl_res = {}
        itens_res = [
            ("v_total",      "Total registros:",   COR_TEXTO),
            ("v_conciliado", "✅ Conciliados:",    COR_VERDE),
            ("v_parcial",    "⚠ Parciais:",        COR_AMARELO),
            ("v_pendente",   "❌ Pendentes:",      COR_VERMELHO),
            ("v_ignorado",   "🚫 Ignorados:",      COR_CINZA),
            ("v_soma",       "Σ Valor Vendas:",    COR_TEXTO),
        ]
        for k, lbl, cor in itens_res:
            row = tk.Frame(frame, bg=COR_PAINEL)
            row.pack(fill="x", padx=12, pady=1)
            tk.Label(row, text=lbl, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 8)).pack(side="left")
            l = tk.Label(row, text="—", bg=COR_PAINEL, fg=cor,
                         font=("Segoe UI", 8, "bold"))
            l.pack(side="right")
            self.lbl_res[k] = l

        sep2 = tk.Frame(frame, bg=COR_BORDA, height=1)
        sep2.pack(fill="x", padx=12, pady=6)
        tk.Label(frame, text="RESUMO BANCO PIX", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)

        itens_banco = [
            ("b_total",      "Total registros:",   COR_TEXTO),
            ("b_conciliado", "✅ Conciliados:",    COR_VERDE),
            ("b_pendente",   "❌ Pendentes:",      COR_VERMELHO),
            ("b_soma",       "Σ Valor Banco:",     COR_TEXTO),
            ("diferenca",    "Δ Diferença:",       COR_AZUL),
        ]
        for k, lbl, cor in itens_banco:
            row = tk.Frame(frame, bg=COR_PAINEL)
            row.pack(fill="x", padx=12, pady=1)
            tk.Label(row, text=lbl, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 8)).pack(side="left")
            l = tk.Label(row, text="—", bg=COR_PAINEL, fg=cor,
                         font=("Segoe UI", 8, "bold"))
            l.pack(side="right")
            self.lbl_res[k] = l

        sep3 = tk.Frame(frame, bg=COR_BORDA, height=1)
        sep3.pack(fill="x", padx=12, pady=8)
        tk.Label(frame, text="MANUAL", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)
        tk.Label(frame,
                 text="1. Clique em uma ou mais VENDAS (tabela cima)\n"
                      "   (clique novamente para desmarcar)\n"
                      "2. Clique em um ou mais PIX (tabela baixo)\n"
                      "3. Pressione 'Conciliar Manual'",
                 bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=12, pady=2)

        self.lbl_sel_v = tk.Label(frame, text="Venda: (nenhuma)",
                                  bg=COR_PAINEL, fg=COR_VERDE,
                                  font=("Segoe UI", 8, "italic"), wraplength=210)
        self.lbl_sel_v.pack(anchor="w", padx=12)
        self.lbl_sel_b = tk.Label(frame, text="PIX banco: (nenhum)",
                                  bg=COR_PAINEL, fg=COR_AZUL,
                                  font=("Segoe UI", 8, "italic"), wraplength=210)
        self.lbl_sel_b.pack(anchor="w", padx=12, pady=(2,0))

        tk.Button(frame, text="Limpar seleção", bg=COR_BG, fg=COR_TEXTO_SEC,
                  font=("Segoe UI", 8), relief="flat", cursor="hand2",
                  command=self.limpar_selecao).pack(anchor="w", padx=12, pady=(6,0))



    def _build_tabelas(self, parent):
        frame = tk.Frame(parent, bg=COR_BG)
        frame.pack(side="left", fill="both", expand=True, pady=10)

        tk.Label(frame, text="VENDAS PIX  (Cupom Fiscal + Nota Fiscal + Recibos)",
                 bg=COR_BG, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0,2))

        frm_v = tk.Frame(frame, bg=COR_BG)
        frm_v.pack(fill="both", expand=True)

        self.cols_v = ["origem", "referencia", "valor", "saldo_rest", "descricao", "status", "par_banco"]
        self.tree_v = ttk.Treeview(frm_v, columns=self.cols_v, show="headings",
                                   selectmode="extended", height=12)
        largs_v = {"origem":80,"referencia":100,"valor":90,"saldo_rest":90,
                   "descricao":380,"status":90,"par_banco":90}
        for col in self.cols_v:
            self.tree_v.heading(col, text=col.upper())
            self.tree_v.column(col, width=largs_v.get(col,100),
                               anchor="center" if col in ("valor","saldo_rest","status","par_banco") else "w")

        sb_vy = ttk.Scrollbar(frm_v, orient="vertical", command=self.tree_v.yview)
        sb_vx = ttk.Scrollbar(frm_v, orient="horizontal", command=self.tree_v.xview)
        self.tree_v.configure(yscrollcommand=sb_vy.set, xscrollcommand=sb_vx.set)
        sb_vy.pack(side="right", fill="y")
        sb_vx.pack(side="bottom", fill="x")
        self.tree_v.pack(fill="both", expand=True)
        self.tree_v.bind("<ButtonRelease-1>", self._on_click_venda)
        self._cfg_tags(self.tree_v)

        tk.Label(frame, text="MOVIMENTAÇÃO PIX QRCODE — BANCO",
                 bg=COR_BG, fg=COR_LARANJA,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8,2))

        frm_b = tk.Frame(frame, bg=COR_BG)
        frm_b.pack(fill="both", expand=True)

        self.cols_b = ["DT_RECEB","HR_RECEB","VENDEDOR","PEDIDO","TXID","VALOR","saldo_rest","status","par_venda"]
        self.tree_b = ttk.Treeview(frm_b, columns=self.cols_b, show="headings",
                                   selectmode="extended", height=12)
        largs_b = {"DT_RECEB":90,"HR_RECEB":70,"VENDEDOR":70,"PEDIDO":70,
                   "TXID":260,"VALOR":90,"saldo_rest":90,"status":90,"par_venda":90}
        for col in self.cols_b:
            self.tree_b.heading(col, text=col.upper())
            self.tree_b.column(col, width=largs_b.get(col,90),
                               anchor="center" if col in ("VALOR","saldo_rest","status","par_venda","HR_RECEB","DT_RECEB") else "w")

        sb_by = ttk.Scrollbar(frm_b, orient="vertical", command=self.tree_b.yview)
        sb_bx = ttk.Scrollbar(frm_b, orient="horizontal", command=self.tree_b.xview)
        self.tree_b.configure(yscrollcommand=sb_by.set, xscrollcommand=sb_bx.set)
        sb_by.pack(side="right", fill="y")
        sb_bx.pack(side="bottom", fill="x")
        self.tree_b.pack(fill="both", expand=True)
        self.tree_b.bind("<ButtonRelease-1>", self._on_click_banco)
        self._cfg_tags(self.tree_b)

    def _cfg_tags(self, tree):
        aplicar_tags_tree(tree)

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=COR_PAINEL, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.status_var = tk.StringVar(value="Pronto. Carregue os relatórios para começar.")
        tk.Label(bar, textvariable=self.status_var, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8), anchor="w").pack(side="left", padx=10)

    def _aplicar_estilos(self):
        aplicar_estilos_ttk(ttk.Style(self))
        registrar_callback(self._aplicar_tema)

    def _aplicar_tema(self):
        """Reaplicar cores de todos os widgets após troca de tema."""
        _cores()
        self.configure(bg=T("BG"))
        aplicar_estilos_ttk(ttk.Style(self))
        for tree in (self.tree_v, self.tree_b):
            aplicar_tags_tree(tree)
        from core.theme import recolorir_widget
        recolorir_widget(self)
        self.atualizar_tabelas()


    def _carregar(self, chave, func, label, eh_pdf=False):
        ft_pdf   = [("PDF", "*.pdf"), ("Todos", "*.*")]
        ft_excel = [("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")]
        path = filedialog.askopenfilename(
            title=f"Selecionar {label}",
            filetypes=ft_pdf if eh_pdf else ft_excel)
        if not path:
            return
        try:
            self.status_var.set(f"⏳ Carregando {label}...")
            self.update()
            df = func(path)
            self.paths[chave] = path
            n = len(df)
            self.lbl_paths[chave].config(
                text=f"✅ {os.path.basename(path)} ({n} reg.)")
            self.status_var.set(
                f"✅ {label} carregado — {n} registros PIX encontrados.")
            return df
        except Exception as e:
            import traceback
            detalhe = traceback.format_exc()
            messagebox.showerror("Erro", f"Erro ao carregar {label}:\n{e}\n\n{detalhe}")
            return None

    def _unificar_vendas(self):
        partes = []
        for chave in ("cupom", "nf", "recibo"):
            if hasattr(self, f"_df_{chave}") and getattr(self, f"_df_{chave}") is not None:
                partes.append(getattr(self, f"_df_{chave}"))
        if not partes:
            return None
        df = pd.concat(partes, ignore_index=True)
        df["status"]    = "pendente"
        df["par_banco"] = ""
        df["saldo_rest"] = df["valor"]
        return df

    def carregar_cupom(self):
        df = self._carregar("cupom", ler_cupom_fiscal, "Cupom Fiscal", eh_pdf=True)
        if df is not None:
            self._df_cupom = df
            self._recarregar_vendas()

    def carregar_nf(self):
        df = self._carregar("nf", ler_nota_fiscal, "Nota Fiscal", eh_pdf=True)
        if df is not None:
            self._df_nf = df
            self._recarregar_vendas()

    def carregar_recibo(self):
        df = self._carregar("recibo", ler_recibos, "Recibos", eh_pdf=True)
        if df is not None:
            self._df_recibo = df
            self._recarregar_vendas()

    def carregar_banco(self):
        df = self._carregar("banco", ler_mov_pix, "Mov. PIX Banco", eh_pdf=True)
        if df is not None:
            self.df_banco = df
            self.atualizar_tabelas()
            self.atualizar_resumo()

    def _recarregar_vendas(self):
        self.df_vendas = self._unificar_vendas()
        self.limpar_selecao()
        self.atualizar_tabelas()
        self.atualizar_resumo()


    def conciliar_auto(self):
        if self.df_vendas is None or self.df_banco is None:
            messagebox.showwarning("Aviso", "Carregue os relatórios de vendas e o banco primeiro.")
            return
        self.status_var.set("⏳ Conciliando automaticamente...")
        self.update()
        self.df_vendas, self.df_banco = conciliar_automatico(self.df_vendas, self.df_banco)
        self.limpar_selecao()
        self.atualizar_tabelas()
        self.atualizar_resumo()
        n_cv = (self.df_vendas["status"] == "conciliado").sum()
        n_pv = (self.df_vendas["status"] == "pendente").sum()
        n_cb = (self.df_banco["status"]  == "conciliado").sum()
        n_pb = (self.df_banco["status"]  == "pendente").sum()
        self.status_var.set(
            f"🔄 Auto concluída — Vendas: ✅{n_cv} ❌{n_pv} | Banco: ✅{n_cb} ❌{n_pb}")

    def conciliar_manual(self):
        if not self.sel_vendas or not self.sel_bancos:
            messagebox.showwarning("Seleção incompleta",
                "Selecione pelo menos uma VENDA e pelo menos um registro PIX do banco.")
            return
        tol = 0.01
        for iv in self.sel_vendas:
            for ib in self.sel_bancos:
                sv = self.df_vendas.at[iv, "saldo_rest"]
                sb = self.df_banco.at[ib, "saldo_rest"]
                if sv <= tol or sb <= tol:
                    continue
                val = min(sv, sb)
                par = f"M{iv}-{ib}"
                def ap(df, idx, col, p):
                    a = df.at[idx, col]
                    df.at[idx, col] = f"{a},{p}" if a else p
                ap(self.df_vendas, iv, "par_banco", par)
                ap(self.df_banco,  ib, "par_venda", par)
                self.df_vendas.at[iv, "saldo_rest"] = round(sv - val, 2)
                self.df_banco.at[ib, "saldo_rest"]  = round(sb - val, 2)

        def novo_st_v(idx):
            r = self.df_vendas.loc[idx]
            if not r["par_banco"]: return "pendente"
            return "conciliado" if r["saldo_rest"] <= tol else "parcial"
        def novo_st_b(idx):
            r = self.df_banco.loc[idx]
            if not r["par_venda"]: return "pendente"
            return "conciliado" if r["saldo_rest"] <= tol else "parcial"

        for iv in self.sel_vendas:
            self.df_vendas.at[iv, "status"] = novo_st_v(iv)
        for ib in self.sel_bancos:
            self.df_banco.at[ib, "status"] = novo_st_b(ib)

        self.limpar_selecao()
        self.atualizar_tabelas()
        self.atualizar_resumo()
        self.status_var.set("🤝 Conciliação manual aplicada.")

    def desconciliar(self):
        self._alterar_status_selecionados(None)

    def ignorar(self):
        self._alterar_status_selecionados("ignorado")

    def _alterar_status_selecionados(self, novo_status):
        for tree, df, col_par in [
            (self.tree_v, self.df_vendas, "par_banco"),
            (self.tree_b, self.df_banco,  "par_venda"),
        ]:
            if df is None:
                continue
            for item in tree.selection():
                idx = self._item_id(tree, item)
                if idx is None:
                    continue
                if novo_status is None:
                    orig = df.at[idx, "valor"] if "valor" in df.columns else df.at[idx, "VALOR"]
                    df.at[idx, "status"]    = "pendente"
                    df.at[idx, col_par]     = ""
                    df.at[idx, "saldo_rest"] = orig
                else:
                    df.at[idx, "status"] = novo_status
        self.atualizar_tabelas()
        self.atualizar_resumo()
        acao = "desconciliados" if novo_status is None else "ignorados"
        self.status_var.set(f"✔ Registros {acao}.")


    def _on_click_venda(self, event):
        item = self.tree_v.identify_row(event.y)
        if not item:
            return
        idx = self._item_id(self.tree_v, item)
        if idx is None:
            return
        if idx in self.sel_vendas:
            self.sel_vendas.remove(idx)
        else:
            self.sel_vendas.append(idx)
        n = len(self.sel_vendas)
        if n == 0:
            self.lbl_sel_v.config(text="Venda: (nenhuma)")
        elif n == 1:
            row = self.df_vendas.loc[self.sel_vendas[0]]
            self.lbl_sel_v.config(
                text=f"Venda: R$ {row['valor']:,.2f} | {str(row['referencia'])[:30]}")
        else:
            soma = sum(self.df_vendas.at[i, "valor"] for i in self.sel_vendas)
            self.lbl_sel_v.config(
                text=f"Vendas: {n} selecionada(s) | Σ R$ {soma:,.2f}")
        self._destacar()

    def _on_click_banco(self, event):
        item = self.tree_b.identify_row(event.y)
        if not item:
            return
        idx = self._item_id(self.tree_b, item)
        if idx is None:
            return
        if self.sel_vendas:
            if idx in self.sel_bancos:
                self.sel_bancos.remove(idx)
            else:
                self.sel_bancos.append(idx)
            self.lbl_sel_b.config(
                text=f"PIX banco: {len(self.sel_bancos)} selecionado(s)")
        self._destacar()

    def _destacar(self):
        if self.df_vendas is not None:
            for item in self.tree_v.get_children():
                idx = self._item_id(self.tree_v, item)
                if idx is None: continue
                tag = self.df_vendas.at[idx, "status"]
                self.tree_v.item(item, tags=(tag,))
            for iv in self.sel_vendas:
                it = self._buscar_item(self.tree_v, iv)
                if it:
                    self.tree_v.item(it, tags=("selecionado",))
        if self.df_banco is not None:
            for item in self.tree_b.get_children():
                idx = self._item_id(self.tree_b, item)
                if idx is None: continue
                tag = self.df_banco.at[idx, "status"]
                self.tree_b.item(item, tags=(tag,))
            for ib in self.sel_bancos:
                it = self._buscar_item(self.tree_b, ib)
                if it:
                    self.tree_b.item(it, tags=("selecionado",))

    def limpar_selecao(self):
        self.sel_vendas = []
        self.sel_bancos = []
        self.lbl_sel_v.config(text="Venda: (nenhuma)")
        self.lbl_sel_b.config(text="PIX banco: (nenhum)")


    def atualizar_tabelas(self):
        self._popular_tree_vendas()
        self._popular_tree_banco()

    def _popular_tree_vendas(self):
        self._map_v = {}
        self._rmap_v = {}
        self.tree_v.delete(*self.tree_v.get_children())
        if self.df_vendas is None:
            return
        df = self._filtrar_v(self.df_vendas)
        for z, (idx, row) in enumerate(df.iterrows()):
            vals = (
                row.get("origem",""),
                row.get("referencia",""),
                f"{row['valor']:,.2f}",
                f"{row['saldo_rest']:,.2f}",
                str(row.get("descricao",""))[:70],
                row.get("status",""),
                row.get("par_banco",""),
            )
            st = row.get("status","pendente")
            tag = st if st in ("conciliado","parcial","pendente","ignorado") else "pendente"
            item = self.tree_v.insert("", "end", values=vals, tags=(tag,))
            self._map_v[item]  = idx
            self._rmap_v[idx]  = item

    def _popular_tree_banco(self):
        self._map_b  = {}
        self._rmap_b = {}
        self.tree_b.delete(*self.tree_b.get_children())
        if self.df_banco is None:
            return
        df = self._filtrar_b(self.df_banco)
        for z, (idx, row) in enumerate(df.iterrows()):
            vals = (
                str(row.get("DT_RECEB","")),
                str(row.get("HR_RECEB","")),
                str(row.get("VENDEDOR","")),
                str(row.get("PEDIDO","")),
                str(row.get("TXID",""))[:55],
                f"{row['VALOR']:,.2f}",
                f"{row['saldo_rest']:,.2f}",
                row.get("status",""),
                row.get("par_venda",""),
            )
            st  = row.get("status","pendente")
            tag = st if st in ("conciliado","parcial","pendente","ignorado") else "pendente"
            item = self.tree_b.insert("", "end", values=vals, tags=(tag,))
            self._map_b[item]  = idx
            self._rmap_b[idx]  = item

    def _filtrar_v(self, df):
        df = self._filtrar(df, "status")
        df = self._filtrar_data(df, "data")
        return df

    def _filtrar_b(self, df):
        df = self._filtrar(df, "status")
        df = self._filtrar_data(df, "DT_RECEB")
        return df

    def _filtrar(self, df, col_status):
        f = self.filtro_status.get()
        if f == "Todos":
            return df
        return df[df[col_status] == f]

    def _filtrar_data(self, df, col):
        """Filtra por intervalo de data. Espera formato dd/mm/aaaa nas entries."""
        if col not in df.columns:
            return df
        ini_txt = self.filtro_data_ini.get().strip()
        fim_txt = self.filtro_data_fim.get().strip()
        if not ini_txt and not fim_txt:
            return df

        def para_data(txt):
            for fmt in ("%d/%m/%Y", "%d/%m/%y"):
                try:
                    return datetime.strptime(txt, fmt).date()
                except ValueError:
                    pass
            return None

        ini = para_data(ini_txt) if ini_txt else None
        fim = para_data(fim_txt) if fim_txt else None

        def data_linha(v):
            v = str(v).strip()
            for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(v, fmt).date()
                except ValueError:
                    pass
            return None

        mask = pd.Series([True] * len(df), index=df.index)
        if ini:
            mask &= df[col].apply(lambda v: (data_linha(v) or datetime.min.date()) >= ini)
        if fim:
            mask &= df[col].apply(lambda v: (data_linha(v) or datetime.max.date()) <= fim)
        return df[mask]

    def _limpar_filtro_data(self):
        self.filtro_data_ini.delete(0, "end")
        self.filtro_data_fim.delete(0, "end")
        self.atualizar_tabelas()

    def atualizar_resumo(self):
        if self.df_vendas is not None:
            cnt = self.df_vendas["status"].value_counts()
            self.lbl_res["v_total"].config(text=str(len(self.df_vendas)))
            for s in ("conciliado","parcial","pendente","ignorado"):
                self.lbl_res[f"v_{s}"].config(text=str(cnt.get(s,0)))
            self.lbl_res["v_soma"].config(text=f"R$ {self.df_vendas['valor'].sum():,.2f}")

        if self.df_banco is not None:
            cnt = self.df_banco["status"].value_counts()
            self.lbl_res["b_total"].config(text=str(len(self.df_banco)))
            self.lbl_res["b_conciliado"].config(text=str(cnt.get("conciliado",0)))
            self.lbl_res["b_pendente"].config(text=str(cnt.get("pendente",0)))
            soma_b = self.df_banco["VALOR"].sum()
            self.lbl_res["b_soma"].config(text=f"R$ {soma_b:,.2f}")

            if self.df_vendas is not None:
                soma_v = self.df_vendas["valor"].sum()
                dif    = soma_v - soma_b
                cor    = COR_VERDE if abs(dif) < 0.05 else COR_VERMELHO
                self.lbl_res["diferenca"].config(
                    text=f"R$ {dif:,.2f}", fg=cor)


    def _item_id(self, tree, item):
        if tree == self.tree_v:
            return self._map_v.get(item)
        return self._map_b.get(item)

    def _buscar_item(self, tree, idx):
        if tree == self.tree_v:
            return self._rmap_v.get(idx)
        return self._rmap_b.get(idx)
    