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
    global COR_VERMELHO, COR_CINZA, COR_AZUL, COR_ROXO
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
    COR_ROXO      = T("ROXO")

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

    # Verifica se o PDF tem texto selecionavel
    with pdfplumber.open(path) as pdf:
        tem_texto = any(len(p.chars) > 0 for p in pdf.pages[:3])

    if tem_texto:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""
                for linha in texto.splitlines():
                    linhas.append(linha)
    else:
        # PDF baseado em imagem -- usa OCR
        try:
            from pdf2image import convert_from_path
            import pytesseract
            pages_img = convert_from_path(path, dpi=200)
            for img in pages_img:
                texto = pytesseract.image_to_string(
                    img, lang="por", config="--psm 6")
                for linha in texto.splitlines():
                    linhas.append(linha)
        except ImportError:
            # fallback sem OCR
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text() or ""
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

_GATILHOS_CARTAO_CUPOM = [
    "VENDA A CARTAO DEBITO",
    "VENDA A CARTAO DE CREDITO",
]

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

        m = re.match(r"^(\d{4,})\s+\d{4,}\s+\S+\s+\w+\s+(\S+)\s+(\S+)", linha_strip)
        if m:
            cond_atual   = m.group(2)
            if cond_atual.upper().startswith("GOOD"):
                cupom_atual = None
                continue
            cupom_atual  = m.group(1)
            vender_atual = m.group(3)
            continue

        if any(up.startswith(t) for t in ["TOTAL GERAL", "TOTAL DIA",
                                           "DESCRICAO"]):
            cupom_atual = None
            continue

        for gatilho in _GATILHOS_CARTAO_CUPOM:
            if up.startswith(gatilho):
                if not cupom_atual:
                    break
                valor = _ultimo_num_linha(linha_strip)
                if valor > 0:
                    tipo = "Débito" if "DEBITO" in up else "Crédito"
                    registros.append({
                        "origem":     "Cupom Fiscal",
                        "referencia": f"Cupom {cupom_atual}",
                        "data":       data_atual,
                        "tipo":       tipo,
                        "valor":      round(valor, 2),
                        "descricao":  (f"{gatilho.title()} | Cupom {cupom_atual} "
                                       f"| Cond: {cond_atual} | Vend: {vender_atual}"),
                        "status":     "pendente",
                        "par_banco":  "",
                    })
                break

    return pd.DataFrame(registros)

_GATILHOS_CARTAO_NF = [
    "VENDA A CARTAO DEBITO",
    "VENDA A CARTAO DE CREDITO",
]

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

        m = re.match(r"^(\d{4,})\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(.+?)\s{2,}(\S+)\s+([\d\.,]+)", linha_strip)
        if m:
            cond_atual    = m.group(7).strip()
            if cond_atual.upper().startswith("GOOD"):
                nota_atual = None
                continue
            nota_atual    = m.group(1)
            cliente_atual = m.group(6).strip()
            continue

        if re.match(r"^\d{4,}\s", linha_strip):
            partes = linha_strip.split()
            nums_ini = sum(1 for p in partes[:6] if re.match(r"^\d+$", p))
            if nums_ini >= 4:
                # tenta extrair a COND: token não-numérico logo antes dos valores finais
                cond_fb = ""
                for p in reversed(partes):
                    if re.match(r"^[\d\.,]+$", p):
                        continue
                    cond_fb = p
                    break
                if cond_fb.upper().startswith("GOOD"):
                    nota_atual = None
                    continue
                nota_atual    = partes[0]
                cliente_atual = " ".join(p for p in partes[4:10]
                                         if not re.match(r"^[\d\.,]+$", p))
                cond_atual    = cond_fb
            continue

        for gatilho in _GATILHOS_CARTAO_NF:
            if up.startswith(gatilho):
                if not nota_atual:
                    break
                valor = _ultimo_num_linha(linha_strip)
                if valor > 0:
                    tipo = "Débito" if "DEBITO" in up else "Crédito"
                    registros.append({
                        "origem":     "Nota Fiscal",
                        "referencia": f"NF {nota_atual}",
                        "data":       data_atual,
                        "tipo":       tipo,
                        "valor":      round(valor, 2),
                        "descricao":  (f"{gatilho.title()} | NF {nota_atual} "
                                       f"| {cliente_atual} | Cond: {cond_atual}"),
                        "status":     "pendente",
                        "par_banco":  "",
                    })
                break

    return pd.DataFrame(registros)

def ler_recibos(path):

    linhas = _extrair_linhas_pdf(path)
    registros    = []
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

            tem_cartao = ("CART. CRD" in up or "CART.CRD" in up or
                          "CART. DEB" in up or "CART.DEB" in up)
            if not tem_cartao:
                continue

            partes = linha_limpa.split()
            numero_recibo = partes[1] if len(partes) > 1 else "?"

            if "CART. CRD" in up or "CART.CRD" in up:
                tipo_pgto = "Crédito"
            elif "CART. DEB" in up or "CART.DEB" in up:
                tipo_pgto = "Débito"
            else:
                tipo_pgto = "Misto"

            recibo_atual = {
                "origem":     "Recibo",
                "referencia": f"Recibo {numero_recibo}",
                "data":       data_atual,
                "tipo":       tipo_pgto,
                "valor":      0.0,
                "descricao":  f"CARTÃO {tipo_pgto.upper()} | Recibo {numero_recibo}",
                "status":     "pendente",
                "par_banco":  "",
            }
            continue

        if recibo_atual and (up.startswith("DPP") or up.startswith("ANT")):
            partes_dpp = linha_limpa.split()
            if len(partes_dpp) < 3 or not re.match(r"^\d{8,}$", partes_dpp[1]):
                continue
            nums = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", linha_limpa)
            IDX_JR_CART   = 3
            IDX_CART_DEB  = 7
            IDX_CART_CRED = 8
            jr_cart   = _num(nums[IDX_JR_CART])   if len(nums) > IDX_JR_CART   else 0.0
            cart_deb  = _num(nums[IDX_CART_DEB])  if len(nums) > IDX_CART_DEB  else 0.0
            cart_cred = _num(nums[IDX_CART_CRED]) if len(nums) > IDX_CART_CRED else 0.0
            recibo_atual["valor"] = round(recibo_atual["valor"] + jr_cart + cart_deb + cart_cred, 2)

    if recibo_atual and recibo_atual["valor"] > 0:
        registros.append(recibo_atual)

    df = pd.DataFrame(registros)
    if not df.empty:
        df["saldo_rest"] = df["valor"]
    return df

def ler_mov_cartao(path):

    try:
        df_raw = pd.read_excel(path, sheet_name="ANALITICO", header=7,
                               engine="openpyxl")
    except Exception:
        df_raw = pd.read_excel(path, header=7, engine="openpyxl")

    df_raw.columns = [str(c).replace("\n", " ").strip() for c in df_raw.columns]

    col_desc  = "Descrição do Lançamento"
    col_valor = "Valor Bruto"
    col_dt    = "Data/Hora  da Venda"
    col_autho = "Número da Autorização"
    col_cv    = "Número do Comprovante  de Vendas"
    col_term  = "Número do Terminal"
    col_parc  = "Total de Parcelas"
    col_cartao= "Cartões"

    df_raw = df_raw[df_raw[col_desc].notna()].copy()
    mask = df_raw[col_desc].str.upper().str.contains(
        r"DEBITO|CREDITO|PARCELADO", na=False)
    df_raw = df_raw[mask].copy()

    df_raw[col_valor] = pd.to_numeric(df_raw[col_valor], errors="coerce")

    cancelamentos = df_raw[df_raw[col_valor] < 0].copy()
    if not cancelamentos.empty:
        cancel_por_autho = (
            cancelamentos.groupby(col_autho)[col_valor]
            .sum()
            .abs()
            .to_dict()
        )
        idx_cancelamentos = set(cancelamentos.index.tolist())

        idx_anulados = set()
        for autho, valor_cancel in cancel_por_autho.items():
            autho_str = str(autho).strip()
            positivos = df_raw[
                (df_raw[col_valor] > 0) &
                (df_raw[col_autho].astype(str).str.strip() == autho_str)
            ].copy()
            if positivos.empty:
                continue
            restante = round(valor_cancel, 2)
            for idx_pos in positivos.index:
                if restante <= 0:
                    break
                val_pos = round(float(df_raw.at[idx_pos, col_valor]), 2)
                if restante >= val_pos:
                    idx_anulados.add(idx_pos)
                    restante = round(restante - val_pos, 2)
                else:
                    df_raw.at[idx_pos, col_valor] = round(val_pos - restante, 2)
                    restante = 0.0

        idx_remover = idx_cancelamentos | idx_anulados
        df_raw = df_raw[~df_raw.index.isin(idx_remover)].copy()

    df_raw = df_raw[df_raw[col_valor] > 0].copy()
    df_raw[col_valor] = df_raw[col_valor].round(2)

    registros = []
    for _, row in df_raw.iterrows():
        dt_hora = str(row.get(col_dt, "")).strip()
        partes_dt = dt_hora.split(" ")
        dt   = partes_dt[0] if len(partes_dt) > 0 else ""
        hora = partes_dt[1] if len(partes_dt) > 1 else ""

        desc  = str(row.get(col_desc, "")).strip()
        up_d  = desc.upper()
        if "DEBITO" in up_d:
            tipo = "Débito"
        elif "PARCELADO" in up_d:
            tipo = "Crédito Parcelado"
        else:
            tipo = "Crédito"

        cv       = str(row.get(col_cv,    "")).strip()
        autho    = str(row.get(col_autho, "")).strip()
        terminal = str(row.get(col_term,  "")).strip()
        parcelas = row.get(col_parc, 1)
        try:
            parcelas = int(float(parcelas))
        except:
            parcelas = 1
        cartao   = str(row.get(col_cartao, "")).strip()
        valor    = float(row[col_valor])

        registros.append({
            "DT_VENDA":  dt,
            "HR_VENDA":  hora,
            "CARTAO":    cartao,
            "TIPO":      tipo,
            "AUTHO":     autho,
            "CV":        cv,
            "TERMINAL":  terminal,
            "PARCELAS":  parcelas,
            "VALOR":     valor,
            "status":    "pendente",
            "par_venda": "",
            "saldo_rest": valor,
        })

    COLS = ["DT_VENDA","HR_VENDA","CARTAO","TIPO","AUTHO","CV",
            "TERMINAL","PARCELAS","VALOR","status","par_venda","saldo_rest"]
    if not registros:
        df = pd.DataFrame(columns=COLS)
        df["VALOR"]      = pd.Series(dtype=float)
        df["saldo_rest"] = pd.Series(dtype=float)
    else:
        df = pd.DataFrame(registros, columns=COLS)
        df["saldo_rest"] = df["VALOR"]

    return df

def conciliar_automatico(df_vendas, df_banco, tolerancia=0.01):
    dv = df_vendas.copy()
    db = df_banco.copy()

    dv["status"]     = "pendente"
    dv["par_banco"]  = ""
    dv["saldo_rest"] = dv["valor"]
    db["status"]     = "pendente"
    db["par_venda"]  = ""
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

        for ib, _ in candidatos.iterrows():
            saldo_v = dv.at[iv, "saldo_rest"]
            saldo_b = db.at[ib, "saldo_rest"]
            if saldo_v <= tolerancia or saldo_b <= tolerancia:
                continue
            val = min(saldo_v, saldo_b)
            par = novo_par()
            dv.at[iv, "par_banco"]  += ("," if dv.at[iv, "par_banco"] else "") + par
            dv.at[iv, "saldo_rest"]  = round(saldo_v - val, 2)
            db.at[ib, "par_venda"]  += ("," if db.at[ib, "par_venda"] else "") + par
            db.at[ib, "saldo_rest"]  = round(saldo_b - val, 2)

    def st_v(row):
        if not row["par_banco"]: return "pendente"
        return "conciliado" if row["saldo_rest"] <= tolerancia else "parcial"

    def st_b(row):
        if not row["par_venda"]: return "pendente"
        return "conciliado" if row["saldo_rest"] <= tolerancia else "parcial"

    dv["status"] = dv.apply(st_v, axis=1)
    db["status"] = db.apply(st_b, axis=1)

    return dv, db

class Conciliacao_Cartao(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Conciliador CARTÕES — PMZ Peças e Pneus")
        self.geometry("1600x880")
        self.configure(bg=COR_BG)
        self.resizable(True, True)

        self.df_vendas = None
        self.df_banco  = None
        self.paths     = {"cupom": None, "nf": None, "recibo": None, "banco": None}
        self.sel_vendas = []
        self.sel_bancos = []

        self._build_ui()
        self._aplicar_estilos()

    def _build_ui(self):
        self._build_topbar()
        corpo = tk.Frame(self, bg=COR_BG)
        corpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._build_painel_esq(corpo)
        self._build_tabelas(corpo)
        self._build_statusbar()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=COR_PAINEL, height=58)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        self._topbar = bar

        tk.Label(bar, text="💳  Conciliador CARTÕES — PMZ",
                 bg=COR_PAINEL, fg=COR_TEXTO,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16, pady=12)

        for txt, cor, cmd in [
            ("🔄  Conciliar Auto",   COR_AZUL,    self.conciliar_auto),
            ("🤝  Conciliar Manual", COR_VERDE,   self.conciliar_manual),
            ("🔓  Desconciliar",     COR_AMARELO, self.desconciliar),
            ("🚫  Ignorar",          COR_CINZA,   self.ignorar),
            ("📊  Exportar XLSX",    COR_ROXO,    self.exportar_xlsx),
        ]:
            tk.Button(bar, text=txt, bg=cor, fg="white",
                      font=("Segoe UI", 9, "bold"), relief="flat",
                      padx=12, pady=6, cursor="hand2",
                      command=cmd).pack(side="left", padx=4, pady=12)

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

        # Tipo
        tk.Label(bar, text="Tipo:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8)).pack(side="left", padx=(8, 2))
        self.filtro_tipo = ttk.Combobox(bar, state="readonly", width=14,
            values=["Todos", "Débito", "Crédito", "Crédito Parcelado"])
        self.filtro_tipo.set("Todos")
        self.filtro_tipo.pack(side="left", pady=12)
        self.filtro_tipo.bind("<<ComboboxSelected>>", lambda _: self.atualizar_tabelas())

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
        frame = tk.Frame(parent, bg=COR_PAINEL, width=235)
        frame.pack(side="left", fill="y", padx=(0, 10), pady=10)
        frame.pack_propagate(False)

        tk.Label(frame, text="ARQUIVOS", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(14, 4))

        self.lbl_paths = {}
        for chave, label, cmd in [
            ("cupom",  "📄 Cupom Fiscal",  self.carregar_cupom),
            ("nf",     "🧾 Nota Fiscal",   self.carregar_nf),
            ("recibo", "📋 Recibos",       self.carregar_recibo),
            ("banco",  "🏦 Extrato Cartões", self.carregar_banco),
        ]:
            tk.Button(frame, text=label, bg=COR_ACENTO2, fg="white",
                      font=("Segoe UI", 8, "bold"), relief="flat",
                      padx=8, pady=4, cursor="hand2", anchor="w",
                      command=cmd).pack(fill="x", padx=12, pady=(4, 0))
            lbl = tk.Label(frame, text="(não carregado)", bg=COR_PAINEL,
                           fg=COR_TEXTO_SEC, font=("Segoe UI", 7),
                           wraplength=210, justify="left")
            lbl.pack(anchor="w", padx=14, pady=(0, 4))
            self.lbl_paths[chave] = lbl

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=12, pady=8)
        tk.Label(frame, text="RESUMO VENDAS CARTÃO", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)

        self.lbl_res = {}
        for k, lbl, cor in [
            ("v_total",      "Total registros:",  COR_TEXTO),
            ("v_conciliado", "✅ Conciliados:",   COR_VERDE),
            ("v_parcial",    "⚠ Parciais:",       COR_AMARELO),
            ("v_pendente",   "❌ Pendentes:",     COR_VERMELHO),
            ("v_ignorado",   "🚫 Ignorados:",     COR_CINZA),
            ("v_soma",       "Σ Valor Vendas:",   COR_TEXTO),
        ]:
            row = tk.Frame(frame, bg=COR_PAINEL)
            row.pack(fill="x", padx=12, pady=1)
            tk.Label(row, text=lbl, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 8)).pack(side="left")
            l = tk.Label(row, text="—", bg=COR_PAINEL, fg=cor,
                         font=("Segoe UI", 8, "bold"))
            l.pack(side="right")
            self.lbl_res[k] = l

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=12, pady=6)
        tk.Label(frame, text="RESUMO EXTRATO CARTÕES", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)

        for k, lbl, cor in [
            ("b_total",      "Total registros:",  COR_TEXTO),
            ("b_conciliado", "✅ Conciliados:",   COR_VERDE),
            ("b_pendente",   "❌ Pendentes:",     COR_VERMELHO),
            ("b_soma",       "Σ Valor Banco:",    COR_TEXTO),
            ("diferenca",    "Δ Diferença:",      COR_AZUL),
        ]:
            row = tk.Frame(frame, bg=COR_PAINEL)
            row.pack(fill="x", padx=12, pady=1)
            tk.Label(row, text=lbl, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 8)).pack(side="left")
            l = tk.Label(row, text="—", bg=COR_PAINEL, fg=cor,
                         font=("Segoe UI", 8, "bold"))
            l.pack(side="right")
            self.lbl_res[k] = l

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=12, pady=8)
        tk.Label(frame, text="MANUAL", bg=COR_PAINEL, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12)
        tk.Label(frame,
                 text="1. Clique em uma ou mais VENDAS (tabela cima)\n"
                      "   (clique novamente para desmarcar)\n"
                      "2. Clique em um ou mais do extrato (tabela baixo)\n"
                      "3. Pressione 'Conciliar Manual'",
                 bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 8), justify="left").pack(anchor="w", padx=12, pady=2)

        self.lbl_sel_v = tk.Label(frame, text="Venda: (nenhuma)",
                                  bg=COR_PAINEL, fg=COR_VERDE,
                                  font=("Segoe UI", 8, "italic"), wraplength=210)
        self.lbl_sel_v.pack(anchor="w", padx=12)
        self.lbl_sel_b = tk.Label(frame, text="Banco: (nenhum)",
                                  bg=COR_PAINEL, fg=COR_AZUL,
                                  font=("Segoe UI", 8, "italic"), wraplength=210)
        self.lbl_sel_b.pack(anchor="w", padx=12, pady=(2, 0))

        tk.Button(frame, text="Limpar seleção", bg=COR_BG, fg=COR_TEXTO_SEC,
                  font=("Segoe UI", 8), relief="flat", cursor="hand2",
                  command=self.limpar_selecao).pack(anchor="w", padx=12, pady=(6, 0))

    def _build_tabelas(self, parent):
        frame = tk.Frame(parent, bg=COR_BG)
        frame.pack(side="left", fill="both", expand=True, pady=10)

        tk.Label(frame,
                 text="VENDAS CARTÃO  (Cupom Fiscal + Nota Fiscal + Recibos)",
                 bg=COR_BG, fg=COR_ACENTO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 2))

        frm_v = tk.Frame(frame, bg=COR_BG)
        frm_v.pack(fill="both", expand=True)

        self.cols_v = ["origem", "referencia", "data", "tipo", "valor",
                       "saldo_rest", "descricao", "status", "par_banco"]
        self.tree_v = ttk.Treeview(frm_v, columns=self.cols_v,
                                   show="headings", selectmode="extended", height=12)
        largs_v = {"origem": 80, "referencia": 100, "data": 80, "tipo": 80, "valor": 80,
                   "saldo_rest": 80, "descricao": 380, "status": 80, "par_banco": 80}
        for col in self.cols_v:
            self.tree_v.heading(col, text=col.upper())
            self.tree_v.column(col, width=largs_v.get(col, 90),
                               anchor="center" if col in ("valor", "saldo_rest",
                                                          "status", "par_banco",
                                                          "tipo", "data") else "w")
        sb_vy = ttk.Scrollbar(frm_v, orient="vertical",   command=self.tree_v.yview)
        sb_vx = ttk.Scrollbar(frm_v, orient="horizontal", command=self.tree_v.xview)
        self.tree_v.configure(yscrollcommand=sb_vy.set, xscrollcommand=sb_vx.set)
        sb_vy.pack(side="right",  fill="y")
        sb_vx.pack(side="bottom", fill="x")
        self.tree_v.pack(fill="both", expand=True)
        self.tree_v.bind("<ButtonRelease-1>", self._on_click_venda)
        self._cfg_tags(self.tree_v)

        tk.Label(frame, text="EXTRATO CARTÕES — BANCO (XLSX)",
                 bg=COR_BG, fg=COR_ROXO,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 2))

        frm_b = tk.Frame(frame, bg=COR_BG)
        frm_b.pack(fill="both", expand=True)

        self.cols_b = ["DT_VENDA","HR_VENDA","CARTAO","TIPO","AUTHO",
                       "CV","TERMINAL","PARCELAS","VALOR",
                       "saldo_rest","status","par_venda"]
        self.tree_b = ttk.Treeview(frm_b, columns=self.cols_b,
                                   show="headings", selectmode="extended", height=12)
        largs_b = {"DT_VENDA": 80, "HR_VENDA": 65, "CARTAO": 130, "TIPO": 90,
                   "AUTHO": 80, "CV": 80, "TERMINAL": 75, "PARCELAS": 55,
                   "VALOR": 80, "saldo_rest": 80, "status": 80, "par_venda": 80}
        for col in self.cols_b:
            self.tree_b.heading(col, text=col.upper())
            self.tree_b.column(col, width=largs_b.get(col, 80),
                               anchor="center" if col in ("VALOR", "saldo_rest",
                                                          "status", "par_venda",
                                                          "PARCELAS", "HR_VENDA",
                                                          "DT_VENDA", "TIPO") else "w")
        sb_by = ttk.Scrollbar(frm_b, orient="vertical",   command=self.tree_b.yview)
        sb_bx = ttk.Scrollbar(frm_b, orient="horizontal", command=self.tree_b.xview)
        self.tree_b.configure(yscrollcommand=sb_by.set, xscrollcommand=sb_bx.set)
        sb_by.pack(side="right",  fill="y")
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
        _cores()
        self.configure(bg=T("BG"))
        aplicar_estilos_ttk(ttk.Style(self))
        for tree in (self.tree_v, self.tree_b):
            aplicar_tags_tree(tree)
        from core.theme import recolorir_widget
        recolorir_widget(self)
        self.atualizar_tabelas()


    def _carregar(self, chave, func, label, eh_pdf=True):
        ft_pdf   = [("PDF",   "*.pdf"),               ("Todos", "*.*")]
        ft_excel = [("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")]
        path = filedialog.askopenfilename(
            title=f"Selecionar {label}",
            filetypes=ft_pdf if eh_pdf else ft_excel)
        if not path:
            return None
        try:
            self.status_var.set(f"⏳ Carregando {label}...")
            self.update()
            df = func(path)
            self.paths[chave] = path
            n = len(df)
            self.lbl_paths[chave].config(
                text=f"✅ {os.path.basename(path)} ({n} reg.)")
            self.status_var.set(f"✅ {label} carregado — {n} registros.")
            return df
        except Exception as e:
            import traceback
            messagebox.showerror("Erro",
                f"Erro ao carregar {label}:\n{e}\n\n{traceback.format_exc()}")
            return None

    def _unificar_vendas(self):
        partes = []
        for chave in ("cupom", "nf", "recibo"):
            df = getattr(self, f"_df_{chave}", None)
            if df is not None and not df.empty:
                partes.append(df)
        if not partes:
            return None
        df = pd.concat(partes, ignore_index=True)
        df["status"]     = "pendente"
        df["par_banco"]  = ""
        df["saldo_rest"] = df["valor"]
        return df

    def carregar_cupom(self):
        df = self._carregar("cupom", ler_cupom_fiscal, "Cupom Fiscal")
        if df is not None:
            self._df_cupom = df
            self._recarregar_vendas()

    def carregar_nf(self):
        df = self._carregar("nf", ler_nota_fiscal, "Nota Fiscal")
        if df is not None:
            self._df_nf = df
            self._recarregar_vendas()

    def carregar_recibo(self):
        df = self._carregar("recibo", ler_recibos, "Recibos")
        if df is not None:
            self._df_recibo = df
            self._recarregar_vendas()

    def carregar_banco(self):
        df = self._carregar("banco", ler_mov_cartao,
                            "Extrato Cartões (XLSX)", eh_pdf=False)
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
            messagebox.showwarning("Aviso",
                "Carregue os relatórios de vendas e o extrato do banco primeiro.")
            return
        self.status_var.set("⏳ Conciliando automaticamente...")
        self.update()
        self.df_vendas, self.df_banco = conciliar_automatico(
            self.df_vendas, self.df_banco)
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
                "Selecione pelo menos uma VENDA e pelo menos um registro do extrato.")
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
                    df.at[idx, "status"]     = "pendente"
                    df.at[idx, col_par]      = ""
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
                text=f"Banco: {len(self.sel_bancos)} selecionado(s)")
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
        self.lbl_sel_b.config(text="Banco: (nenhum)")

    def atualizar_tabelas(self):
        self._popular_tree_vendas()
        self._popular_tree_banco()

    def _popular_tree_vendas(self):
        self._map_v  = {}
        self._rmap_v = {}
        self.tree_v.delete(*self.tree_v.get_children())
        if self.df_vendas is None:
            return
        df = self._filtrar_v(self.df_vendas)
        for _, (idx, row) in enumerate(df.iterrows()):
            vals = (
                row.get("origem",     ""),
                row.get("referencia", ""),
                row.get("data",       ""),
                row.get("tipo",       ""),
                f"{row['valor']:,.2f}",
                f"{row['saldo_rest']:,.2f}",
                str(row.get("descricao", ""))[:60],
                row.get("status",    ""),
                row.get("par_banco", ""),
            )
            st  = row.get("status", "pendente")
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
        for _, (idx, row) in enumerate(df.iterrows()):
            vals = (
                str(row.get("DT_VENDA",  "")),
                str(row.get("HR_VENDA",  "")),
                str(row.get("CARTAO",    ""))[:22],
                str(row.get("TIPO",      "")),
                str(row.get("AUTHO",     "")),
                str(row.get("CV",        "")),
                str(row.get("TERMINAL",  "")),
                str(row.get("PARCELAS",  "")),
                f"{row['VALOR']:,.2f}",
                f"{row['saldo_rest']:,.2f}",
                row.get("status",    ""),
                row.get("par_venda", ""),
            )
            st  = row.get("status", "pendente")
            tag = st if st in ("conciliado","parcial","pendente","ignorado") else "pendente"
            item = self.tree_b.insert("", "end", values=vals, tags=(tag,))
            self._map_b[item]  = idx
            self._rmap_b[idx]  = item

    def _filtrar_v(self, df):
        df = self._filtrar(df, "status")
        ft = self.filtro_tipo.get()
        if ft != "Todos" and "tipo" in df.columns:
            df = df[df["tipo"] == ft]
        df = self._filtrar_data(df, "data")
        return df

    def _filtrar_b(self, df):
        df = self._filtrar(df, "status")
        ft = self.filtro_tipo.get()
        if ft != "Todos" and "TIPO" in df.columns:
            df = df[df["TIPO"] == ft]
        df = self._filtrar_data(df, "DT_VENDA")
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
            for s in ("conciliado", "parcial", "pendente", "ignorado"):
                self.lbl_res[f"v_{s}"].config(text=str(cnt.get(s, 0)))
            self.lbl_res["v_soma"].config(
                text=f"R$ {self.df_vendas['valor'].sum():,.2f}")

        if self.df_banco is not None:
            cnt = self.df_banco["status"].value_counts()
            self.lbl_res["b_total"].config(text=str(len(self.df_banco)))
            self.lbl_res["b_conciliado"].config(text=str(cnt.get("conciliado", 0)))
            self.lbl_res["b_pendente"].config(text=str(cnt.get("pendente", 0)))
            soma_b = self.df_banco["VALOR"].sum()
            self.lbl_res["b_soma"].config(text=f"R$ {soma_b:,.2f}")

            if self.df_vendas is not None:
                soma_v = self.df_vendas["valor"].sum()
                dif    = soma_v - soma_b
                cor    = COR_VERDE if abs(dif) < 0.05 else COR_VERMELHO
                self.lbl_res["diferenca"].config(
                    text=f"R$ {dif:,.2f}", fg=cor)

    def exportar_xlsx(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        if self.df_vendas is None and self.df_banco is None:
            messagebox.showwarning("Aviso", "Nenhum dado carregado para exportar.")
            return

        path_out = filedialog.asksaveasfilename(
            title="Salvar relatório de conciliação",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialfile="conciliacao_cartao.xlsx")
        if not path_out:
            return

        try:
            wb = Workbook()
            borda = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))

            def hdr_style(cell, cor_hex="1E3A5F"):
                cell.font      = Font(bold=True, color="FFFFFF", size=9)
                cell.fill      = PatternFill("solid", fgColor=cor_hex)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border    = borda

            def val_cell(cell, bold=False):
                cell.font      = Font(bold=bold, size=9)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border    = borda

            def sub_hdr(cell):
                cell.font      = Font(size=9)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border    = borda

            # ── Aba Resumo ────────────────────────────────────────────────────
            ws_res = wb.active
            ws_res.title = "Resumo"
            ws_res.merge_cells("A1:B1")
            ws_res["A1"] = "RELATÓRIO DE CONCILIAÇÃO — CARTÕES"
            ws_res["A1"].font      = Font(bold=True, size=12, color="FFFFFF")
            ws_res["A1"].fill      = PatternFill("solid", fgColor="1E3A5F")
            ws_res["A1"].alignment = Alignment(horizontal="center")
            ws_res.merge_cells("A2:B2")
            ws_res["A2"] = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            ws_res["A2"].font      = Font(italic=True, size=9)
            ws_res["A2"].alignment = Alignment(horizontal="center")

            row = 4

            # Bloco Vendas
            ws_res.cell(row, 1, "VENDAS CARTÃO (PDFs)")
            hdr_style(ws_res.cell(row, 1))
            ws_res.merge_cells(f"A{row}:B{row}")
            row += 1

            if self.df_vendas is not None:
                dv = self.df_vendas
                cnt_v = dv["status"].value_counts()
                soma_v = dv["valor"].sum()
                soma_conc_v = dv.loc[dv["status"] == "conciliado", "valor"].sum()
                soma_pend_v = dv.loc[dv["status"] == "pendente",   "valor"].sum()
                soma_parc_v = dv.loc[dv["status"] == "parcial",    "valor"].sum()
                soma_ign_v  = dv.loc[dv["status"] == "ignorado",   "valor"].sum()
                items_v = [
                    ("Total de registros",  len(dv),                        False),
                    ("✅ Conciliados",       cnt_v.get("conciliado", 0),     False),
                    ("⚠ Parciais",          cnt_v.get("parcial",    0),     False),
                    ("❌ Pendentes",         cnt_v.get("pendente",   0),     False),
                    ("🚫 Ignorados",         cnt_v.get("ignorado",   0),     False),
                    ("Σ Valor Total Vendas", f"R$ {soma_v:,.2f}",           True),
                    ("Σ Valor Conciliado",   f"R$ {soma_conc_v:,.2f}",      False),
                    ("Σ Valor Pendente",     f"R$ {soma_pend_v:,.2f}",      False),
                    ("Σ Valor Parcial",      f"R$ {soma_parc_v:,.2f}",      False),
                    ("Σ Valor Ignorado",     f"R$ {soma_ign_v:,.2f}",       False),
                    ("% Conciliado",         f"{100*soma_conc_v/soma_v:.1f}%" if soma_v else "—", False),
                ]
            else:
                items_v = [("(sem dados)", "—", False)]
                soma_v = 0

            for label, valor, bold in items_v:
                c1 = ws_res.cell(row, 1, label); c2 = ws_res.cell(row, 2, valor)
                sub_hdr(c1); val_cell(c2, bold)
                row += 1

            row += 1

            # Bloco Banco
            ws_res.cell(row, 1, "EXTRATO CARTÕES — BANCO")
            hdr_style(ws_res.cell(row, 1), cor_hex="4A235A")
            ws_res.merge_cells(f"A{row}:B{row}")
            row += 1

            if self.df_banco is not None:
                db = self.df_banco
                cnt_b = db["status"].value_counts()
                soma_b = db["VALOR"].sum()
                soma_conc_b = db.loc[db["status"] == "conciliado", "VALOR"].sum()
                soma_pend_b = db.loc[db["status"] == "pendente",   "VALOR"].sum()
                items_b = [
                    ("Total de registros",  len(db),                        False),
                    ("✅ Conciliados",       cnt_b.get("conciliado", 0),     False),
                    ("❌ Pendentes",         cnt_b.get("pendente",   0),     False),
                    ("Σ Valor Total Banco",  f"R$ {soma_b:,.2f}",           True),
                    ("Σ Valor Conciliado",   f"R$ {soma_conc_b:,.2f}",      False),
                    ("Σ Valor Pendente",     f"R$ {soma_pend_b:,.2f}",      False),
                    ("% Conciliado",         f"{100*soma_conc_b/soma_b:.1f}%" if soma_b else "—", False),
                ]
            else:
                items_b = [("(sem dados)", "—", False)]
                soma_b = 0

            for label, valor, bold in items_b:
                c1 = ws_res.cell(row, 1, label); c2 = ws_res.cell(row, 2, valor)
                sub_hdr(c1); val_cell(c2, bold)
                row += 1

            row += 1

            # Bloco Diferença
            if self.df_vendas is not None and self.df_banco is not None:
                ws_res.cell(row, 1, "DIFERENÇA")
                hdr_style(ws_res.cell(row, 1), cor_hex="1D6A3A")
                ws_res.merge_cells(f"A{row}:B{row}")
                row += 1
                dif = soma_v - soma_b
                cor_dif = "1D6A3A" if abs(dif) < 0.05 else "C0392B"
                c1 = ws_res.cell(row, 1, "Δ Vendas − Banco")
                c2 = ws_res.cell(row, 2, f"R$ {dif:,.2f}")
                sub_hdr(c1)
                c2.font      = Font(bold=True, color=cor_dif)
                c2.alignment = Alignment(horizontal="right")
                c2.border    = borda

            ws_res.column_dimensions["A"].width = 32
            ws_res.column_dimensions["B"].width = 22

            # ── Aba Vendas ────────────────────────────────────────────────────
            if self.df_vendas is not None:
                ws_v = wb.create_sheet("Vendas")
                cols_exp = ["origem", "referencia", "data", "tipo", "valor",
                            "saldo_rest", "descricao", "status", "par_banco"]
                for ci, col in enumerate(cols_exp, 1):
                    hdr_style(ws_v.cell(1, ci, col.upper()))
                for ri, (_, r) in enumerate(self.df_vendas.iterrows(), 2):
                    for ci, col in enumerate(cols_exp, 1):
                        ws_v.cell(ri, ci, r.get(col, ""))
                largs = {"descricao": 50, "referencia": 16, "origem": 14,
                         "data": 12, "tipo": 12, "valor": 12,
                         "saldo_rest": 12, "status": 12, "par_banco": 14}
                for ci, col in enumerate(cols_exp, 1):
                    ws_v.column_dimensions[get_column_letter(ci)].width = largs.get(col, 14)

            # ── Aba Banco ─────────────────────────────────────────────────────
            if self.df_banco is not None:
                ws_b = wb.create_sheet("Banco")
                cols_b = ["DT_VENDA", "HR_VENDA", "CARTAO", "TIPO", "AUTHO",
                          "CV", "TERMINAL", "PARCELAS", "VALOR",
                          "saldo_rest", "status", "par_venda"]
                for ci, col in enumerate(cols_b, 1):
                    hdr_style(ws_b.cell(1, ci, col.upper()), cor_hex="4A235A")
                for ri, (_, r) in enumerate(self.df_banco.iterrows(), 2):
                    for ci, col in enumerate(cols_b, 1):
                        ws_b.cell(ri, ci, r.get(col, ""))
                for ci, col in enumerate(cols_b, 1):
                    ws_b.column_dimensions[get_column_letter(ci)].width = (
                        30 if col == "CARTAO" else 14)

            # ── Aba Pendentes ─────────────────────────────────────────────────
            ws_p = wb.create_sheet("Pendentes")
            ws_p["A1"] = "VENDAS PENDENTES"
            hdr_style(ws_p["A1"])
            ws_p.merge_cells("A1:I1")
            row_p = 2
            cols_pv = ["origem", "referencia", "data", "tipo", "valor",
                       "saldo_rest", "descricao", "status", "par_banco"]
            for ci, col in enumerate(cols_pv, 1):
                hdr_style(ws_p.cell(row_p, ci, col.upper()))
            row_p += 1
            if self.df_vendas is not None:
                for _, r in self.df_vendas[self.df_vendas["status"] == "pendente"].iterrows():
                    for ci, col in enumerate(cols_pv, 1):
                        ws_p.cell(row_p, ci, r.get(col, ""))
                    row_p += 1

            row_p += 1
            ws_p.cell(row_p, 1, "BANCO PENDENTES")
            hdr_style(ws_p.cell(row_p, 1), cor_hex="4A235A")
            ws_p.merge_cells(f"A{row_p}:L{row_p}")
            row_p += 1
            cols_pb = ["DT_VENDA", "HR_VENDA", "CARTAO", "TIPO", "AUTHO",
                       "CV", "TERMINAL", "PARCELAS", "VALOR", "saldo_rest",
                       "status", "par_venda"]
            for ci, col in enumerate(cols_pb, 1):
                hdr_style(ws_p.cell(row_p, ci, col.upper()), cor_hex="4A235A")
            row_p += 1
            if self.df_banco is not None:
                for _, r in self.df_banco[self.df_banco["status"] == "pendente"].iterrows():
                    for ci, col in enumerate(cols_pb, 1):
                        ws_p.cell(row_p, ci, r.get(col, ""))
                    row_p += 1

            for ci in range(1, 13):
                ws_p.column_dimensions[get_column_letter(ci)].width = 16
            ws_p.column_dimensions["G"].width = 40

            wb.save(path_out)
            self.status_var.set(f"✅ Exportado: {os.path.basename(path_out)}")
            messagebox.showinfo("Exportação concluída",
                f"Relatório salvo com sucesso!\n\n{path_out}")

        except Exception as e:
            import traceback
            messagebox.showerror("Erro na exportação",
                f"{e}\n\n{traceback.format_exc()}")

    def _item_id(self, tree, item):
        if tree == self.tree_v:
            return self._map_v.get(item)
        return self._map_b.get(item)

    def _buscar_item(self, tree, idx):
        if tree == self.tree_v:
            return self._rmap_v.get(idx)
        return self._rmap_b.get(idx)