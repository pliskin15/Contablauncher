"""
Cruzamento Balancete x Adiantamento
-----------------------------------
Cruza a aba 'NÃO CONCILIADOS OK (2)' da planilha de adiantamento (somente
linhas com status 'em aberto') com o balancete (.xls do Protheus, conta
ADIANTAMENTOS A FORN. DIVERSOS) e gera um .xlsx com a análise.

Dependências:  pip install pandas openpyxl xlrd
"""

import os
import re
import sys
import unicodedata
import threading
import traceback
from datetime import datetime

import pandas as pd

# --------------------------------------------------------------------------
# CONFIGURAÇÃO
# --------------------------------------------------------------------------
ABA_ADIANTAMENTO = "NÃO CONCILIADOS OK (2)"
STATUS_ALVO = "em aberto"
TOLERANCIA = 0.01  # diferença máxima (R$) para considerar "bateu"

# Colunas esperadas no adiantamento
COL_AD_CC = "Cód. CC"
COL_AD_VALOR = "Valor (R$)"
COL_AD_STATUS = "status"
COL_AD_TIPO = "Tipo"
COL_AD_FORN = "Fornecedor (Centro de Custo)"

# Colunas esperadas no balancete
COL_BAL_CC = "C. CUSTO"
COL_BAL_NOME = "CENTRO DE CUSTO"
COL_BAL_SALDO = "SALDO ATUAL"

# Status possíveis no resultado
ST_OK = "CONFERE"
ST_DIF = "DIVERGENTE"
ST_SEM_BAL = "NÃO ENCONTRADO NO BALANCETE"
ST_AMBIGUO = "AMBÍGUO (VÁRIAS FILIAIS)"
ST_SO_BAL = "SÓ NO BALANCETE"


# --------------------------------------------------------------------------
# UTILITÁRIOS
# --------------------------------------------------------------------------
def normaliza_nome(txt) -> str:
    """Maiúsculo, sem acento, sem pontuação, sem espaços duplicados."""
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return ""
    s = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", " ", s.upper())
    return s.strip()


def nomes_compativeis(a: str, b: str) -> bool:
    """
    Nomes do balancete costumam vir truncados (ex.: 'RIVAS RAITZ COM' x
    'RIVAS RAITZ COMERCIO DE PECAS LTDA'). Considera compatível se um
    é prefixo do outro ou se são iguais depois de normalizar.
    """
    na, nb = normaliza_nome(a), normaliza_nome(b)
    if not na or not nb:
        return False
    return na == nb or na.startswith(nb) or nb.startswith(na)


def normaliza_cc(valor) -> str:
    """Cód. CC do adiantamento: texto, sem espaços, sem '.0' de float."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = str(valor).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


# --------------------------------------------------------------------------
# LEITURA
# --------------------------------------------------------------------------
def le_adiantamento(caminho: str) -> pd.DataFrame:
    """Lê a aba alvo e mantém só as linhas com status 'em aberto'."""
    try:
        df = pd.read_excel(caminho, sheet_name=ABA_ADIANTAMENTO, dtype=object)
    except ValueError as e:
        raise ValueError(
            f"Aba '{ABA_ADIANTAMENTO}' não encontrada no arquivo de adiantamento."
        ) from e

    df.columns = [str(c).strip() for c in df.columns]
    faltando = [c for c in (COL_AD_CC, COL_AD_VALOR, COL_AD_STATUS) if c not in df.columns]
    if faltando:
        raise ValueError(
            "Colunas não encontradas no adiantamento: " + ", ".join(faltando)
        )

    # o status pode misturar texto e número (ex.: 2025) -> compara como texto
    status = df[COL_AD_STATUS].astype(str).str.strip().str.lower()
    df = df[status == STATUS_ALVO].copy()

    df["_cc"] = df[COL_AD_CC].map(normaliza_cc)
    df["_valor"] = pd.to_numeric(df[COL_AD_VALOR], errors="coerce")
    df = df[(df["_cc"] != "") & df["_valor"].notna()].copy()

    # DÉBITO soma, CRÉDITO subtrai. O balancete trata crédito como negativo
    # (ex.: CC 908 = -20.736,67), então a comparação usa este valor líquido.
    if COL_AD_TIPO in df.columns:
        tipo = df[COL_AD_TIPO].astype(str).str.strip().str.upper()
        df["_valor_liq"] = df["_valor"].where(~tipo.str.startswith("CR"), -df["_valor"])
    else:
        df["_valor_liq"] = df["_valor"]
    return df


def le_balancete(caminho: str) -> pd.DataFrame:
    """
    O balancete é um relatório do Protheus (não é uma tabela limpa).
    Localiza a linha de cabeçalho 'CENTRO DE CUSTO / C. CUSTO' e lê só as
    linhas seguintes que tenham C. CUSTO no formato número-número.
    """
    ext = os.path.splitext(caminho)[1].lower()
    engine = "xlrd" if ext == ".xls" else None
    raw = pd.read_excel(caminho, header=None, dtype=object, engine=engine)

    # acha a linha de cabeçalho
    hdr_idx = None
    for i in range(len(raw)):
        linha = [str(x).strip().upper() for x in raw.iloc[i].tolist() if pd.notna(x)]
        if COL_BAL_CC in linha and COL_BAL_NOME in linha:
            hdr_idx = i
            break
    if hdr_idx is None:
        raise ValueError(
            f"Não encontrei a linha de cabeçalho '{COL_BAL_NOME} / {COL_BAL_CC}' no balancete."
        )

    cab = [str(x).strip().upper() if pd.notna(x) else "" for x in raw.iloc[hdr_idx].tolist()]
    try:
        i_nome = cab.index(COL_BAL_NOME)
        i_cc = cab.index(COL_BAL_CC)
        i_saldo = cab.index(COL_BAL_SALDO)
    except ValueError as e:
        raise ValueError(
            f"Coluna obrigatória ausente no cabeçalho do balancete "
            f"({COL_BAL_NOME}, {COL_BAL_CC}, {COL_BAL_SALDO})."
        ) from e

    corpo = raw.iloc[hdr_idx + 1:, [i_nome, i_cc, i_saldo]].copy()
    corpo.columns = ["nome", "cc", "saldo"]
    corpo["cc"] = corpo["cc"].astype(str).str.strip()

    # só linhas de centro de custo (ex.: 3-1, 1-23, 10-118431); ignora totalizadores
    corpo = corpo[corpo["cc"].str.fullmatch(r"\d+-\d+")].copy()
    corpo["saldo"] = pd.to_numeric(corpo["saldo"], errors="coerce").fillna(0.0)
    corpo["nome"] = corpo["nome"].fillna("").astype(str).str.strip()
    corpo["filial"] = corpo["cc"].str.split("-").str[0]
    corpo["codigo"] = corpo["cc"].str.split("-", n=1).str[1]
    return corpo.reset_index(drop=True)


# --------------------------------------------------------------------------
# CRUZAMENTO
# --------------------------------------------------------------------------
def cruza(adiant: pd.DataFrame, bal: pd.DataFrame):
    """
    Agrupa o adiantamento por Cód. CC (somando Valor) e compara com o
    SALDO ATUAL do balancete.

    Regra do cruzamento: Cód. CC do adiantamento = parte do C. CUSTO após o
    hífen. Se o mesmo código existir em mais de uma filial no balancete, usa
    o nome do fornecedor para escolher a linha correta; se ainda assim não
    for possível decidir, marca como AMBÍGUO e traz todas as filiais.
    """
    ag = (
        adiant.groupby("_cc")
        .agg(
            qtd_linhas=("_valor", "size"),
            valor_bruto=("_valor", "sum"),
            valor_adiantamento=("_valor_liq", "sum"),  # líquido: DÉBITO - CRÉDITO
            fornecedor_adiant=(
                COL_AD_FORN if COL_AD_FORN in adiant.columns else "_cc",
                lambda s: next((str(x).strip() for x in s if pd.notna(x) and str(x).strip()), ""),
            ),
        )
        .reset_index()
        .rename(columns={"_cc": "cod_cc"})
    )

    por_codigo = {k: g for k, g in bal.groupby("codigo")}
    usados = set()  # C. CUSTO do balancete já casados
    linhas = []

    for _, r in ag.iterrows():
        cand = por_codigo.get(r["cod_cc"])
        base = {
            "cod_cc": r["cod_cc"],
            "fornecedor_adiant": r["fornecedor_adiant"],
            "qtd_linhas": int(r["qtd_linhas"]),
            "valor_adiantamento": round(float(r["valor_adiantamento"]), 2),
        }

        if cand is None or cand.empty:
            linhas.append({**base, "c_custo": "", "fornecedor_bal": "", "saldo_balancete": None,
                           "diferenca": None, "status": ST_SEM_BAL, "obs": ""})
            continue

        escolhida, obs = None, ""
        if len(cand) == 1:
            escolhida = cand.iloc[0]
        else:
            # várias filiais: tenta desempatar pelo nome
            por_nome = cand[cand["nome"].map(lambda n: nomes_compativeis(n, r["fornecedor_adiant"]))]
            if len(por_nome) == 1:
                escolhida = por_nome.iloc[0]
                obs = "Desempatado pelo nome (código existe em várias filiais)"
            elif len(por_nome) > 1:
                # mesmo fornecedor em mais de uma filial: compara com a soma e avisa
                escolhida = None
                obs = "Mesmo fornecedor em várias filiais: comparado com a soma das filiais"
                soma = float(por_nome["saldo"].sum())
                dif = round(float(r["valor_adiantamento"]) - soma, 2)
                for c in por_nome["cc"]:
                    usados.add(c)
                linhas.append({
                    **base,
                    "c_custo": " + ".join(por_nome["cc"]),
                    "fornecedor_bal": por_nome.iloc[0]["nome"],
                    "saldo_balancete": round(soma, 2),
                    "diferenca": dif,
                    "status": ST_OK if abs(dif) <= TOLERANCIA else ST_DIF,
                    "obs": obs,
                })
                continue

        if escolhida is None:
            for c in cand["cc"]:
                usados.add(c)
            linhas.append({
                **base,
                "c_custo": " | ".join(cand["cc"]),
                "fornecedor_bal": " | ".join(n or "(sem nome)" for n in cand["nome"]),
                "saldo_balancete": None,
                "diferenca": None,
                "status": ST_AMBIGUO,
                "obs": "Código em várias filiais e nome não permitiu decidir. "
                       "Saldos: " + " | ".join(f"{c}={s:,.2f}" for c, s in zip(cand["cc"], cand["saldo"])),
            })
            continue

        usados.add(escolhida["cc"])
        dif = round(float(r["valor_adiantamento"]) - float(escolhida["saldo"]), 2)
        nome_bal = escolhida["nome"]
        if not nome_bal:
            obs = (obs + "; " if obs else "") + "Balancete sem nome para este CC"
        elif not nomes_compativeis(nome_bal, r["fornecedor_adiant"]):
            obs = (obs + "; " if obs else "") + "Nome do fornecedor difere entre as planilhas"
        linhas.append({
            **base,
            "c_custo": escolhida["cc"],
            "fornecedor_bal": nome_bal,
            "saldo_balancete": round(float(escolhida["saldo"]), 2),
            "diferenca": dif,
            "status": ST_OK if abs(dif) <= TOLERANCIA else ST_DIF,
            "obs": obs,
        })

    res = pd.DataFrame(linhas)

    # CCs que têm saldo no balancete mas não têm nada em aberto no adiantamento
    so_bal = bal[(~bal["cc"].isin(usados)) & (bal["saldo"].abs() > TOLERANCIA)].copy()
    so_bal = so_bal.rename(columns={"cc": "c_custo", "nome": "fornecedor_bal", "saldo": "saldo_balancete"})
    so_bal["dif"] = -so_bal["saldo_balancete"]
    return res, so_bal[["c_custo", "fornecedor_bal", "saldo_balancete", "dif"]].reset_index(drop=True)


# --------------------------------------------------------------------------
# EXPORTAÇÃO
# --------------------------------------------------------------------------
# Layout ÚNICO da aba "Análise". O Resumo e as fórmulas leem daqui.
COLS_ANALISE = [
    # (chave em res,        cabeçalho,                     largura, tipo)
    ("cod_cc",             "Cód. CC",                      12, "txt"),
    ("c_custo",            "C. CUSTO (Balancete)",         20, "txt"),
    ("fornecedor_adiant",  "Fornecedor (Adiantamento)",    42, "txt"),
    ("fornecedor_bal",     "Fornecedor (Balancete)",       42, "txt"),
    ("qtd_linhas",         "Qtd. Linhas",                  10, "int"),
    ("valor_adiantamento", "Valor Adiantamento (R$)",      22, "moeda"),
    ("saldo_balancete",    "Saldo Balancete (R$)",         22, "moeda"),
    ("diferenca",          "Diferença (R$)",               18, "moeda"),   # fórmula
    ("status",             "Situação",                     30, "txt"),
    ("obs",                "Observação",                   60, "txt"),
]
IDX = {k: i + 1 for i, (k, *_r) in enumerate(COLS_ANALISE)}  # posição (1-based)


def _col(letra_de: str) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(IDX[letra_de])


def exporta_xlsx(res: pd.DataFrame, so_bal: pd.DataFrame, adiant: pd.DataFrame,
                 caminho_saida: str, arq_ad: str, arq_bal: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    F = "Arial"
    f_norm = Font(name=F, size=10)
    f_bold = Font(name=F, size=10, bold=True)
    f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
    f_title = Font(name=F, size=14, bold=True)
    fill_head = PatternFill("solid", fgColor="1F3864")
    fill_tot = PatternFill("solid", fgColor="D9E1F2")
    fill_por_status = {
        ST_OK: PatternFill("solid", fgColor="E2EFDA"),
        ST_DIF: PatternFill("solid", fgColor="FCE4D6"),
        ST_AMBIGUO: PatternFill("solid", fgColor="FFF2CC"),
        ST_SEM_BAL: PatternFill("solid", fgColor="EDEDED"),
    }
    lado = Side(style="thin", color="BFBFBF")
    borda = Border(left=lado, right=lado, top=lado, bottom=lado)
    fmt_moeda = '#,##0.00;(#,##0.00);"-"'
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def cabecalho(ws, linha, titulos):
        for c, t in enumerate(titulos, start=1):
            cel = ws.cell(linha, c, t)
            cel.font, cel.fill, cel.border, cel.alignment = f_head, fill_head, borda, center

    wb = Workbook()

    # ------------------------------------------------------------------ Análise
    wa = wb.active
    wa.title = "Análise"
    cabecalho(wa, 1, [c[1] for c in COLS_ANALISE])
    wa.row_dimensions[1].height = 32

    col_ad = _col("valor_adiantamento")
    col_bal = _col("saldo_balancete")
    col_dif = _col("diferenca")
    col_st = _col("status")

    for i, (_, r) in enumerate(res.iterrows(), start=2):
        for key, _tit, _w, tipo in COLS_ANALISE:
            c = IDX[key]
            v = r.get(key)
            if v is not None and isinstance(v, float) and pd.isna(v):
                v = None
            if key == "diferenca":
                # fórmula: só onde há saldo do balancete para comparar
                v = f"={col_ad}{i}-{col_bal}{i}" if r.get("saldo_balancete") is not None and not pd.isna(r.get("saldo_balancete")) else None
            cel = wa.cell(i, c, v)
            cel.font, cel.border = f_norm, borda
            if tipo == "moeda":
                cel.number_format = fmt_moeda
            elif tipo == "int":
                cel.number_format = "#,##0"
            cel.fill = fill_por_status.get(r["status"], PatternFill())
    ult = len(res) + 1

    for key, _t, w, _tipo in COLS_ANALISE:
        from openpyxl.utils import get_column_letter
        wa.column_dimensions[get_column_letter(IDX[key])].width = w
    wa.freeze_panes = "A2"
    wa.auto_filter.ref = f"A1:{get_column_letter(len(COLS_ANALISE))}{max(ult, 2)}"

    # ---------------------------------------------------------- Só no Balancete
    wb_ = wb.create_sheet("Só no Balancete")
    cabecalho(wb_, 1, ["C. CUSTO (Balancete)", "Fornecedor (Balancete)", "Saldo Balancete (R$)"])
    for i, (_, r) in enumerate(so_bal.iterrows(), start=2):
        vals = [r["c_custo"], r["fornecedor_bal"] or "(sem nome)", float(r["saldo_balancete"])]
        for c, v in enumerate(vals, start=1):
            cel = wb_.cell(i, c, v)
            cel.font, cel.border = f_norm, borda
            if c == 3:
                cel.number_format = fmt_moeda
    ult_sb = len(so_bal) + 1
    wb_.column_dimensions["A"].width = 22
    wb_.column_dimensions["B"].width = 48
    wb_.column_dimensions["C"].width = 22
    wb_.freeze_panes = "A2"
    if len(so_bal):
        wb_.auto_filter.ref = f"A1:C{ult_sb}"

    # -------------------------------------------------- Adiantamento (detalhe)
    wd = wb.create_sheet("Detalhe Adiantamento")
    colunas_det = [c for c in (COL_AD_TIPO, "Data", COL_AD_CC, COL_AD_FORN, COL_AD_VALOR, "Histórico")
                   if c in adiant.columns]
    cabecalho(wd, 1, colunas_det)
    for i, (_, r) in enumerate(adiant.iterrows(), start=2):
        for c, nome in enumerate(colunas_det, start=1):
            v = r[nome]
            if isinstance(v, float) and pd.isna(v):
                v = None
            if nome == COL_AD_VALOR:
                v = float(r["_valor"])
            cel = wd.cell(i, c, v)
            cel.font, cel.border = f_norm, borda
            if nome == COL_AD_VALOR:
                cel.number_format = fmt_moeda
    for c, nome in enumerate(colunas_det, start=1):
        wd.column_dimensions[get_column_letter(c)].width = 44 if nome in (COL_AD_FORN, "Histórico") else 16
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = f"A1:{get_column_letter(len(colunas_det))}{len(adiant) + 1}"

    # ------------------------------------------------------------------ Resumo
    ws = wb.create_sheet("Resumo", 0)
    ws["A1"] = "Cruzamento Balancete x Adiantamento"
    ws["A1"].font = f_title
    info = [
        f"Gerado em {datetime.now():%d/%m/%Y %H:%M}",
        f"Adiantamento: {os.path.basename(arq_ad)} (aba '{ABA_ADIANTAMENTO}', status = '{STATUS_ALVO}')",
        f"Balancete: {os.path.basename(arq_bal)}",
        f"Tolerância de comparação: R$ {TOLERANCIA:.2f}  (premissa do programa; ajustável em TOLERANCIA)",
    ]
    for i, t in enumerate(info, start=2):
        ws.cell(i, 1, t).font = f_norm

    cabecalho(ws, 7, ["Situação", "Qtd. CCs", "Valor Adiantamento (R$)", "Saldo Balancete (R$)", "Diferença (R$)"])
    ws.row_dimensions[7].height = 30
    situacoes = [ST_OK, ST_DIF, ST_AMBIGUO, ST_SEM_BAL]
    rng_st = f"Análise!${col_st}$2:${col_st}${ult}"
    rng_ad = f"Análise!${col_ad}$2:${col_ad}${ult}"
    rng_bal = f"Análise!${col_bal}$2:${col_bal}${ult}"
    rng_dif = f"Análise!${col_dif}$2:${col_dif}${ult}"
    for i, st in enumerate(situacoes):
        r = 8 + i
        ws.cell(r, 1, st)
        ws.cell(r, 2, f"=COUNTIF({rng_st},A{r})")
        ws.cell(r, 3, f"=SUMIF({rng_st},A{r},{rng_ad})")
        ws.cell(r, 4, f"=SUMIF({rng_st},A{r},{rng_bal})")
        ws.cell(r, 5, f"=SUMIF({rng_st},A{r},{rng_dif})")
    r_tot = 8 + len(situacoes)
    ws.cell(r_tot, 1, "TOTAL")
    for c, L in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
        ws.cell(r_tot, c, f"=SUM({L}8:{L}{r_tot - 1})")
    for r in range(8, r_tot + 1):
        for c in range(1, 6):
            cel = ws.cell(r, c)
            cel.font = f_bold if r == r_tot else f_norm
            cel.border = borda
            if r == r_tot:
                cel.fill = fill_tot
            elif c == 1:
                cel.fill = fill_por_status[situacoes[r - 8]]
            cel.number_format = "#,##0" if c == 2 else (fmt_moeda if c > 2 else "General")

    r0 = r_tot + 2
    ws.cell(r0, 1, "Saldo no balancete sem adiantamento em aberto").font = f_bold
    ws.cell(r0 + 1, 1, "Qtd. CCs").font = f_norm
    ws.cell(r0 + 2, 1, "Saldo (R$)").font = f_norm
    if len(so_bal):
        ws.cell(r0 + 1, 2, f"=COUNTA('Só no Balancete'!A2:A{ult_sb})")
        ws.cell(r0 + 2, 2, f"=SUM('Só no Balancete'!C2:C{ult_sb})")
    else:
        ws.cell(r0 + 1, 2, 0)
        ws.cell(r0 + 2, 2, 0)
    ws.cell(r0 + 1, 2).number_format = "#,##0"
    ws.cell(r0 + 2, 2).number_format = fmt_moeda
    ws.cell(r0 + 1, 2).font = ws.cell(r0 + 2, 2).font = f_norm

    r1 = r0 + 4
    ws.cell(r1, 1, "Como ler").font = f_bold
    notas = [
        "Valor Adiantamento = soma de 'Valor (R$)' das linhas em aberto, por Cód. CC (DÉBITO soma, CRÉDITO subtrai).",
        "Saldo Balancete = 'SALDO ATUAL' do C. CUSTO correspondente (Cód. CC = parte do C. CUSTO após o hífen).",
        "Diferença = Valor Adiantamento - Saldo Balancete (positivo: adiantamento maior que o balancete).",
        "O prefixo do C. CUSTO é a filial. Se o código existe em várias filiais, o nome do fornecedor desempata.",
        "AMBÍGUO: código em várias filiais e o nome não decidiu; os saldos de cada filial estão em Observação.",
    ]
    for i, t in enumerate(notas, start=1):
        ws.cell(r1 + i, 1, t).font = f_norm

    ws.column_dimensions["A"].width = 46
    for L in "BCDE":
        ws.column_dimensions[L].width = 24

    wb.save(caminho_saida)


# --------------------------------------------------------------------------
# PIPELINE (usado pela interface e pelos testes)
# --------------------------------------------------------------------------
def processa(arq_ad: str, arq_bal: str, arq_saida: str, log=print):
    log("Lendo adiantamento...")
    adiant = le_adiantamento(arq_ad)
    log(f"  {len(adiant):,} linhas em aberto | {adiant['_cc'].nunique():,} Cód. CC distintos")
    if adiant.empty:
        raise ValueError(
            f"Nenhuma linha com status '{STATUS_ALVO}' foi encontrada na aba "
            f"'{ABA_ADIANTAMENTO}'. Verifique se o arquivo está correto."
        )

    log("Lendo balancete...")
    bal = le_balancete(arq_bal)
    log(f"  {len(bal):,} centros de custo | {bal['codigo'].nunique():,} códigos distintos")

    log("Cruzando...")
    res, so_bal = cruza(adiant, bal)
    resumo = res["status"].value_counts().to_dict()
    for k in (ST_OK, ST_DIF, ST_AMBIGUO, ST_SEM_BAL):
        log(f"  {k}: {resumo.get(k, 0):,}")
    log(f"  {ST_SO_BAL}: {len(so_bal):,}")

    log("Gerando Excel...")
    exporta_xlsx(res, so_bal, adiant, arq_saida, arq_ad, arq_bal)
    log(f"Concluído: {arq_saida}")
    return res, so_bal


# --------------------------------------------------------------------------
# INTERFACE TKINTER
# --------------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class CruzamentoFrame(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.v_ad = tk.StringVar()
        self.v_bal = tk.StringVar()
        self.v_out = tk.StringVar()
        self._monta()

    def _monta(self):
        pad = {"padx": 10, "pady": 6}
        fr = ttk.Frame(self, padding=12)
        fr.pack(fill="both", expand=True)
        fr.columnconfigure(1, weight=1)
        fr.rowconfigure(5, weight=1)

        ttk.Label(fr, text="Cruzamento Balancete x Adiantamento",
                  font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", **pad)

        ttk.Label(fr, text="Adiantamento (.xlsx):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(fr, textvariable=self.v_ad).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(fr, text="Procurar...", command=self._pick_ad).grid(row=1, column=2, **pad)

        ttk.Label(fr, text="Balancete (.xls/.xlsx):").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(fr, textvariable=self.v_bal).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(fr, text="Procurar...", command=self._pick_bal).grid(row=2, column=2, **pad)

        ttk.Label(fr, text="Salvar análise em:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(fr, textvariable=self.v_out).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(fr, text="Escolher...", command=self._pick_out).grid(row=3, column=2, **pad)

        self.btn = ttk.Button(fr, text="Gerar análise", command=self._executar)
        self.btn.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)

        self.txt = tk.Text(fr, height=12, state="disabled", font=("Consolas", 9), wrap="word")
        self.txt.grid(row=5, column=0, columnspan=3, sticky="nsew", **pad)
        sb = ttk.Scrollbar(fr, command=self.txt.yview)
        sb.grid(row=5, column=3, sticky="ns")
        self.txt["yscrollcommand"] = sb.set

        self.pb = ttk.Progressbar(fr, mode="indeterminate")
        self.pb.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)

    # --- seleção de arquivos
    def _pick_ad(self):
        p = filedialog.askopenfilename(title="Planilha de adiantamento",
                                       filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if p:
            self.v_ad.set(p)
            self._sugere_saida(p)

    def _pick_bal(self):
        p = filedialog.askopenfilename(title="Balancete",
                                       filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
        if p:
            self.v_bal.set(p)

    def _pick_out(self):
        p = filedialog.asksaveasfilename(title="Salvar análise", defaultextension=".xlsx",
                                         filetypes=[("Excel", "*.xlsx")],
                                         initialfile=self._nome_padrao())
        if p:
            self.v_out.set(p)

    def _nome_padrao(self):
        return f"analise_balancete_adiantamento_{datetime.now():%Y%m%d_%H%M}.xlsx"

    def _sugere_saida(self, base):
        if not self.v_out.get():
            self.v_out.set(os.path.join(os.path.dirname(base), self._nome_padrao()))

    # --- log
    def _log(self, msg):
        def _w():
            self.txt.configure(state="normal")
            self.txt.insert("end", msg + "\n")
            self.txt.see("end")
            self.txt.configure(state="disabled")
        self.after(0, _w)

    # --- execução em thread (não trava a janela)
    def _executar(self):
        ad, bal, out = self.v_ad.get().strip(), self.v_bal.get().strip(), self.v_out.get().strip()
        if not (ad and os.path.isfile(ad)):
            return messagebox.showwarning("Atenção", "Selecione a planilha de adiantamento.")
        if not (bal and os.path.isfile(bal)):
            return messagebox.showwarning("Atenção", "Selecione o balancete.")
        if not out:
            return messagebox.showwarning("Atenção", "Escolha onde salvar a análise.")

        self.btn.configure(state="disabled")
        self.pb.start(12)
        self.txt.configure(state="normal"); self.txt.delete("1.0", "end"); self.txt.configure(state="disabled")

        def trabalho():
            try:
                processa(ad, bal, out, log=self._log)
                self.after(0, lambda: self._fim(True, out))
            except PermissionError:
                self._log("ERRO: não foi possível gravar o arquivo. Ele está aberto no Excel?")
                self.after(0, lambda: self._fim(False, "Não foi possível gravar o arquivo de saída.\nFeche-o no Excel e tente de novo."))
            except Exception as e:  # noqa: BLE001
                self._log("ERRO: " + str(e))
                self._log(traceback.format_exc())
                self.after(0, lambda: self._fim(False, str(e)))

        threading.Thread(target=trabalho, daemon=True).start()

    def _fim(self, ok, msg):
        self.pb.stop()
        self.btn.configure(state="normal")
        if ok:
            if messagebox.askyesno("Concluído", "Análise gerada com sucesso.\n\nAbrir o arquivo agora?"):
                try:
                    os.startfile(msg)  # Windows
                except AttributeError:
                    import subprocess
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", msg])
        else:
            messagebox.showerror("Erro", msg)


def main():
    root = tk.Tk()
    root.title("Cruzamento Balancete x Adiantamento")
    root.geometry("760x520")
    root.minsize(680, 460)
    CruzamentoFrame(root).pack(fill="both", expand=True)
    root.mainloop()



if __name__ == "__main__":
    main()