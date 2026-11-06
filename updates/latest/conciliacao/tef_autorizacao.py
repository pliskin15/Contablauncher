"""
Conciliador de Autorização — Transações x Banco

Cruza a coluna "Autorização" do relatório de Transações com a coluna
"AUTHO" da aba "Banco" do arquivo de conciliação de cartão, liga cada
autorização ao respectivo Nº de Nota/Cupom/Recibo (via aba "Vendas")
e gera um relatório em Excel com o resultado.

Layout esperado dos arquivos de entrada:
  - Transações: cabeçalho na linha 6, coluna L = "Autorização"
  - Conciliação: aba "Banco" (cabeçalho linha 1, coluna E = "AUTHO",
    coluna L = "PAR_VENDA") e aba "Vendas" (cabeçalho linha 1,
    coluna B = referência do documento, coluna I = "PAR_BANCO")

Abas geradas: Resumo, Relacionadas, Nao_Relacionadas, Sem_Autorizacao,
Base_Transacoes, Base_Banco, Base_Vendas.
"""

import collections
import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import openpyxl
from core import theme
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# ----------------------------------------------------------------------
# Posições fixas de leitura (layout dos arquivos de entrada)
# ----------------------------------------------------------------------

LINHA_CAB_TRANSACOES = 6      # cabeçalho do relatório de Transações
COL_DATA_TR = 1                # A: Data
COL_CARTAO_TR = 3              # C: Cartão
COL_PDV_TR = 4                 # D: PDV
COL_VALOR_TR = 8               # H: Valor
COL_TIPO_TR = 11               # K: Tipo Produto
COL_AUTORIZACAO_TR = 12        # L: Autorização
COL_PRODUTO_TR = 16            # P: Produto
COL_DESCRICAO_TR = 17          # Q: Descrição Transação
COL_ESTADO_TR = 19             # S: Estado Transação
COL_CUPOM_TEF_TR = 20          # T: Cupom Fiscal (TEF)

LINHA_CAB_BANCO = 1
COL_DATA_BK = 1                # A: DT_VENDA
COL_HORA_BK = 2                # B: HR_VENDA
COL_CARTAO_BK = 3              # C: CARTAO
COL_TIPO_BK = 4                # D: TIPO
COL_AUTHO_BK = 5               # E: AUTHO
COL_TERMINAL_BK = 7            # G: TERMINAL
COL_VALOR_BK = 9               # I: VALOR
COL_STATUS_BK = 11             # K: STATUS
COL_PAR_VENDA_BK = 12          # L: PAR_VENDA

LINHA_CAB_VENDAS = 1
COL_ORIGEM_VD = 1              # A: ORIGEM
COL_REF_VD = 2                 # B: REFERENCIA (nº da nota/cupom/recibo)
COL_DATA_VD = 3                # C: DATA
COL_TIPO_VD = 4                # D: TIPO
COL_VALOR_VD = 5               # E: VALOR
COL_STATUS_VD = 8              # H: STATUS
COL_PAR_BANCO_VD = 9           # I: PAR_BANCO


def chave(v):
    """Normaliza um valor de autorização para comparação (texto, sem espaços)."""
    if v is None:
        return ""
    return str(v).strip()


def em_branco(v):
    return v in (None, "", " ")


def num(s):
    """'1.400' / '687,16' (texto pt-BR) -> float. Aceita já-numérico e vazio."""
    if s is None or str(s).strip() == "":
        return None
    if isinstance(s, (int, float)):
        return float(s)
    return float(str(s).strip().replace(".", "").replace(",", "."))


def data_para_ordenacao(s):
    """'dd/mm/aaaa' (ou com hora) -> (aaaa, mm, dd) para comparar/ordenar."""
    if not s:
        return (0, 0, 0)
    parte_data = str(s).split(" ")[0]
    d, m, a = parte_data.split("/")
    return (int(a), int(m), int(d))


# ----------------------------------------------------------------------
# Leitura dos arquivos de origem
# ----------------------------------------------------------------------

def ler_transacoes(caminho):
    """Lê o relatório de Transações. Retorna (linhas_com_autorizacao, linhas_sem_autorizacao)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]

    com_auth, sem_auth = [], []
    for r in range(LINHA_CAB_TRANSACOES + 1, ws.max_row + 1):
        auth = ws.cell(r, COL_AUTORIZACAO_TR).value
        linha = dict(
            linha=r,
            data=ws.cell(r, COL_DATA_TR).value,
            cartao=ws.cell(r, COL_CARTAO_TR).value,
            pdv=ws.cell(r, COL_PDV_TR).value,
            valor=ws.cell(r, COL_VALOR_TR).value,
            tipo=ws.cell(r, COL_TIPO_TR).value,
            auth=chave(auth),
            produto=ws.cell(r, COL_PRODUTO_TR).value,
            desc=ws.cell(r, COL_DESCRICAO_TR).value,
            estado=ws.cell(r, COL_ESTADO_TR).value,
            cupom_tef=ws.cell(r, COL_CUPOM_TEF_TR).value,
        )
        linha["valor_n"] = num(linha["valor"])
        if em_branco(auth):
            sem_auth.append(linha)
        else:
            com_auth.append(linha)
    return com_auth, sem_auth


def ler_banco(caminho):
    """Lê a aba 'Banco' do arquivo de conciliação."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    if "Banco" not in wb.sheetnames:
        raise ValueError('O arquivo de conciliação não tem uma aba chamada "Banco".')
    ws = wb["Banco"]

    linhas = []
    for r in range(LINHA_CAB_BANCO + 1, ws.max_row + 1):
        auth = ws.cell(r, COL_AUTHO_BK).value
        if em_branco(auth):
            continue
        linhas.append(dict(
            linha=r,
            data=ws.cell(r, COL_DATA_BK).value,
            hora=ws.cell(r, COL_HORA_BK).value,
            cartao=ws.cell(r, COL_CARTAO_BK).value,
            tipo=ws.cell(r, COL_TIPO_BK).value,
            auth=chave(auth),
            terminal=ws.cell(r, COL_TERMINAL_BK).value,
            valor=ws.cell(r, COL_VALOR_BK).value,
            status=ws.cell(r, COL_STATUS_BK).value,
            par=ws.cell(r, COL_PAR_VENDA_BK).value,
        ))
    return linhas


def ler_vendas(caminho):
    """Lê a aba 'Vendas' do arquivo de conciliação."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    if "Vendas" not in wb.sheetnames:
        raise ValueError('O arquivo de conciliação não tem uma aba chamada "Vendas".')
    ws = wb["Vendas"]

    linhas = []
    for r in range(LINHA_CAB_VENDAS + 1, ws.max_row + 1):
        origem = ws.cell(r, COL_ORIGEM_VD).value
        ref = ws.cell(r, COL_REF_VD).value
        if em_branco(origem) and em_branco(ref):
            continue
        linhas.append(dict(
            linha=r,
            origem=origem,
            ref=ref,
            data=ws.cell(r, COL_DATA_VD).value,
            tipo=ws.cell(r, COL_TIPO_VD).value,
            valor=ws.cell(r, COL_VALOR_VD).value,
            status=ws.cell(r, COL_STATUS_VD).value,
            par=ws.cell(r, COL_PAR_BANCO_VD).value,
        ))
    return linhas


# ----------------------------------------------------------------------
# Cruzamento
# ----------------------------------------------------------------------

def cruzar(tr_rows, bk_rows, vd_rows):
    """Cruza Transações (com autorização) x Banco pela chave de autorização,
    e liga cada lado ao Nº de documento (Vendas) via PAR_VENDA/PAR_BANCO."""

    vd_by_par = collections.defaultdict(list)
    for v in vd_rows:
        if not em_branco(v["par"]):
            vd_by_par[v["par"]].append(v)

    def doc_txt(par):
        if em_branco(par):
            return ""
        return " + ".join(str(v["ref"]) for v in vd_by_par.get(par, []) if v["ref"])

    tr_by = {t["auth"]: t for t in tr_rows}
    bk_by = {b["auth"]: b for b in bk_rows}

    autorizacoes_tr = list(tr_by.keys())
    autorizacoes_bk = list(bk_by.keys())

    relacionadas_auth = [a for a in autorizacoes_tr if a in bk_by]
    so_transacoes_auth = [a for a in autorizacoes_tr if a not in bk_by]
    so_banco_auth = [a for a in autorizacoes_bk if a not in tr_by]

    # terminal mais frequente entre as relacionadas = terminal "principal" (da filial do relatório)
    contagem_terminais = collections.Counter(bk_by[a]["terminal"] for a in relacionadas_auth)
    terminal_principal = contagem_terminais.most_common(1)[0][0] if contagem_terminais else None

    # data de corte do relatório de Transações = maior data encontrada nele
    todas_datas_tr = [t["data"] for t in tr_rows if t["data"]]
    data_corte = max(todas_datas_tr, key=data_para_ordenacao) if todas_datas_tr else None
    data_corte_tupla = data_para_ordenacao(data_corte) if data_corte else None

    return dict(
        tr_by=tr_by, bk_by=bk_by, doc_txt=doc_txt,
        relacionadas_auth=relacionadas_auth,
        so_transacoes_auth=so_transacoes_auth,
        so_banco_auth=so_banco_auth,
        terminal_principal=terminal_principal,
        data_corte=data_corte, data_corte_tupla=data_corte_tupla,
    )


def motivo_so_banco(b, terminal_principal, data_corte_tupla):
    """Categoriza uma autorização 'só no Banco': 0=posterior à data de corte,
    1=mesmo terminal mas sem correspondência real, 2=outro terminal."""
    if terminal_principal and b["terminal"] != terminal_principal:
        return 2, "Terminal diferente do terminal principal do relatório de Transações"
    if data_corte_tupla and data_para_ordenacao(b["data"]) > data_corte_tupla:
        data_fmt = f"{data_corte_tupla[2]:02d}/{data_corte_tupla[1]:02d}/{data_corte_tupla[0]}"
        return 0, f"Posterior ao período do relatório de Transações (que termina em {data_fmt})"
    return 1, "Mesmo terminal do relatório, dentro do período, mas sem autorização na base de Transações"


# ----------------------------------------------------------------------
# Estilo
# ----------------------------------------------------------------------

FN = "Arial"
F_NORM = Font(name=FN, size=10)
F_BOLD = Font(name=FN, size=10, bold=True)
F_HDR = Font(name=FN, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FN, size=14, bold=True)
F_NOTE = Font(name=FN, size=9, italic=True, color="595959")

FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_OK = PatternFill("solid", fgColor="E2EFDA")
FILL_NO = PatternFill("solid", fgColor="FCE4D6")
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_DIVERGE = PatternFill("solid", bgColor="F8CBAD", fgColor="F8CBAD")
F_DIVERGE = Font(name=FN, size=10, bold=True, color="C00000")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTRO = Alignment(horizontal="center", vertical="center")
QUEBRA = Alignment(wrap_text=True, vertical="top")


def hdr(ws, linha, titulos, larguras=None):
    for i, h in enumerate(titulos, 1):
        c = ws.cell(linha, i, h)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = CENTRO
        c.border = BOX
    if larguras:
        for i, w in enumerate(larguras, 1):
            ws.column_dimensions[L(i)].width = w


def estilizar_corpo(ws, r0, r1, ncols, col_texto=()):
    for r in range(r0, r1 + 1):
        for c in range(1, ncols + 1):
            cel = ws.cell(r, c)
            cel.font = F_NORM
            cel.border = BOX
            if c in col_texto:
                cel.number_format = "@"


# ----------------------------------------------------------------------
# Geração do Excel
# ----------------------------------------------------------------------

def gerar_relatorio(caminho_transacoes, caminho_conciliacao, caminho_saida):
    tr_rows, tr_sem = ler_transacoes(caminho_transacoes)
    bk_rows = ler_banco(caminho_conciliacao)
    vd_rows = ler_vendas(caminho_conciliacao)

    x = cruzar(tr_rows, bk_rows, vd_rows)
    tr_by, bk_by, doc_txt = x["tr_by"], x["bk_by"], x["doc_txt"]
    relacionadas_auth = x["relacionadas_auth"]
    so_transacoes_auth = x["so_transacoes_auth"]
    so_banco_auth = x["so_banco_auth"]
    terminal_principal = x["terminal_principal"]
    data_corte_tupla = x["data_corte_tupla"]

    wb = openpyxl.Workbook()

    # =====================================================================
    # ABA: Base_Transacoes
    # =====================================================================
    wt = wb.active
    wt.title = "Base_Transacoes"
    hdr(wt, 1, ["Autorização", "Data", "Cartão", "PDV", "Tipo Produto", "Produto", "Descrição Transação",
                "Valor (R$)", "Estado Transação", "Cupom Fiscal (TEF)", "Linha original",
                "Existe no Banco?", "Terminal no Banco", "Valor Banco (R$)", "Nº Nota/Cupom/Recibo"],
        [14, 20, 20, 11, 13, 14, 26, 13, 16, 16, 10, 16, 16, 15, 24])
    for i, t in enumerate(tr_rows, 2):
        wt.cell(i, 1, t["auth"]); wt.cell(i, 2, t["data"]); wt.cell(i, 3, t["cartao"]); wt.cell(i, 4, t["pdv"])
        wt.cell(i, 5, t["tipo"]); wt.cell(i, 6, t["produto"]); wt.cell(i, 7, t["desc"])
        wt.cell(i, 8, t["valor_n"]); wt.cell(i, 9, t["estado"]); wt.cell(i, 10, t["cupom_tef"]); wt.cell(i, 11, t["linha"])
        b = bk_by.get(t["auth"])
        wt.cell(i, 12, "SIM" if b else "NÃO")
        wt.cell(i, 13, b["terminal"] if b else "")
        wt.cell(i, 14, b["valor"] if b else "")
        wt.cell(i, 15, doc_txt(b["par"]) if b else "")
    last_tr = len(tr_rows) + 1
    estilizar_corpo(wt, 2, last_tr, 15, col_texto=(1,))
    for r in range(2, last_tr + 1):
        wt.cell(r, 8).number_format = "#,##0.00"
        wt.cell(r, 14).number_format = "#,##0.00"
        wt.cell(r, 12).alignment = CENTRO
    wt.freeze_panes = "B2"
    wt.auto_filter.ref = f"A1:O{last_tr}"

    # =====================================================================
    # ABA: Base_Banco
    # =====================================================================
    wbk = wb.create_sheet("Base_Banco")
    hdr(wbk, 1, ["AUTHO", "Data", "Hora", "Cartão", "Tipo", "Terminal", "Valor (R$)", "Status conciliação",
                 "Linha original", "Existe em Transações?", "Situação", "PAR_VENDA", "Nº Nota/Cupom/Recibo"],
        [14, 12, 10, 22, 18, 12, 13, 16, 10, 18, 44, 12, 30])
    for i, b in enumerate(bk_rows, 2):
        wbk.cell(i, 1, b["auth"]); wbk.cell(i, 2, b["data"]); wbk.cell(i, 3, b["hora"]); wbk.cell(i, 4, b["cartao"])
        wbk.cell(i, 5, b["tipo"]); wbk.cell(i, 6, b["terminal"]); wbk.cell(i, 7, b["valor"]); wbk.cell(i, 8, b["status"])
        wbk.cell(i, 9, b["linha"])
        existe = b["auth"] in tr_by
        wbk.cell(i, 10, "SIM" if existe else "NÃO")
        if existe:
            situacao = "Confere nas duas bases"
        else:
            _, situacao = motivo_so_banco(b, terminal_principal, data_corte_tupla)
        wbk.cell(i, 11, situacao)
        wbk.cell(i, 12, b["par"] if not em_branco(b["par"]) else None)
        wbk.cell(i, 13, doc_txt(b["par"]))
    last_bk = len(bk_rows) + 1
    estilizar_corpo(wbk, 2, last_bk, 13, col_texto=(1,))
    for r in range(2, last_bk + 1):
        wbk.cell(r, 7).number_format = "#,##0.00"
        wbk.cell(r, 10).alignment = CENTRO
    wbk.freeze_panes = "B2"
    wbk.auto_filter.ref = f"A1:M{last_bk}"

    # =====================================================================
    # ABA: Base_Vendas
    # =====================================================================
    wvs = wb.create_sheet("Base_Vendas")
    hdr(wvs, 1, ["PAR_BANCO", "Nº Nota/Cupom/Recibo", "Origem", "Data", "Tipo", "Valor (R$)", "Status",
                 "Linha original", "Nº documento(s) do par"],
        [12, 24, 14, 12, 12, 13, 13, 10, 30])
    for i, v in enumerate(vd_rows, 2):
        wvs.cell(i, 1, v["par"] if not em_branco(v["par"]) else None)
        wvs.cell(i, 2, v["ref"]); wvs.cell(i, 3, v["origem"]); wvs.cell(i, 4, v["data"]); wvs.cell(i, 5, v["tipo"])
        wvs.cell(i, 6, v["valor"]); wvs.cell(i, 7, v["status"]); wvs.cell(i, 8, v["linha"])
    last_vd = len(vd_rows) + 1
    # coluna I: documentos agregados por PAR_BANCO, só na 1ª linha de cada par
    # (gravado como valor, não fórmula, pois 2 linhas do mesmo par podem não ser vizinhas)
    vistos = set()
    for i, v in enumerate(vd_rows, 2):
        p = v["par"]
        if em_branco(p) or p in vistos:
            continue
        vistos.add(p)
        wvs.cell(i, 9, doc_txt(p))
    estilizar_corpo(wvs, 2, last_vd, 9, col_texto=(1,))
    for r in range(2, last_vd + 1):
        wvs.cell(r, 6).number_format = "#,##0.00"
    wvs.freeze_panes = "A2"
    wvs.auto_filter.ref = f"A1:I{last_vd}"

    # =====================================================================
    # ABA: Relacionadas
    # =====================================================================
    wr = wb.create_sheet("Relacionadas")
    hdr(wr, 1, ["Autorização", "Nº Nota/Cupom/Recibo", "Data (Transações)", "Data (Banco)", "Terminal",
                "Cartão (Banco)", "Tipo (Banco)", "Valor Transações (R$)", "Valor Banco (R$)",
                "Diferença (R$)", "Confere valor?"],
        [14, 24, 20, 13, 12, 22, 18, 22, 18, 15, 14])
    relacionadas_ordenadas = sorted(relacionadas_auth, key=lambda a: tr_by[a]["linha"])
    for i, a in enumerate(relacionadas_ordenadas, 2):
        t, b = tr_by[a], bk_by[a]
        valor_t = t["valor_n"] or 0
        valor_b = b["valor"] or 0
        diferenca = round(valor_t - valor_b, 2)
        wr.cell(i, 1, a)
        wr.cell(i, 2, doc_txt(b["par"]))
        wr.cell(i, 3, t["data"]); wr.cell(i, 4, b["data"]); wr.cell(i, 5, b["terminal"])
        wr.cell(i, 6, b["cartao"]); wr.cell(i, 7, b["tipo"])
        wr.cell(i, 8, valor_t); wr.cell(i, 9, valor_b); wr.cell(i, 10, diferenca)
        wr.cell(i, 11, "OK" if diferenca == 0 else "DIVERGE")
    last_r = len(relacionadas_ordenadas) + 1
    estilizar_corpo(wr, 2, last_r, 11, col_texto=(1,))
    for r in range(2, last_r + 1):
        for c in (8, 9, 10):
            wr.cell(r, c).number_format = "#,##0.00;[Red]-#,##0.00;0.00"
        wr.cell(r, 11).alignment = CENTRO
    wr.conditional_formatting.add(
        f"A2:K{last_r}",
        FormulaRule(formula=['$K2="DIVERGE"'], fill=FILL_DIVERGE, font=F_DIVERGE),
    )
    wr.freeze_panes = "B2"
    wr.auto_filter.ref = f"A1:K{last_r}"

    # =====================================================================
    # ABA: Nao_Relacionadas
    # =====================================================================
    wn = wb.create_sheet("Nao_Relacionadas")
    hdr(wn, 1, ["Origem", "Autorização", "Nº Nota/Cupom/Recibo", "Data", "Hora",
                "Terminal (Banco) / PDV (Transações)", "Cartão", "Tipo", "Valor (R$)",
                "Status no Banco", "Motivo provável"],
        [22, 14, 26, 14, 10, 24, 22, 18, 13, 16, 52])

    linhas_n = []
    for a in so_transacoes_auth:
        t = tr_by[a]
        linhas_n.append(dict(
            origem="Só em Transações", auth=a, doc="", data=str(t["data"])[:10], hora=str(t["data"])[11:],
            terminal=t["pdv"], cartao=t["cartao"], tipo=t["tipo"], valor=t["valor_n"], status="—",
            motivo="Venda efetuada no PDV, mas o Banco não trouxe essa autorização. "
                   "Verifique manualmente na aba Vendas se há um documento pendente com valor e data compatíveis.",
            ordem=-1,
        ))
    for a in so_banco_auth:
        b = bk_by[a]
        ordem, motivo = motivo_so_banco(b, terminal_principal, data_corte_tupla)
        linhas_n.append(dict(
            origem="Só no Banco", auth=a, doc=doc_txt(b["par"]) or "sem venda vinculada",
            data=b["data"], hora=b["hora"], terminal=b["terminal"], cartao=b["cartao"], tipo=b["tipo"],
            valor=b["valor"], status=b["status"], motivo=motivo, ordem=ordem,
        ))
    linhas_n.sort(key=lambda x: (x["ordem"], x["data"] or ""))

    for i, ln in enumerate(linhas_n, 2):
        wn.cell(i, 1, ln["origem"]); wn.cell(i, 2, ln["auth"]); wn.cell(i, 3, ln["doc"])
        wn.cell(i, 4, ln["data"]); wn.cell(i, 5, ln["hora"]); wn.cell(i, 6, ln["terminal"])
        wn.cell(i, 7, ln["cartao"]); wn.cell(i, 8, ln["tipo"]); wn.cell(i, 9, ln["valor"])
        wn.cell(i, 10, ln["status"]); wn.cell(i, 11, ln["motivo"])
    last_n = len(linhas_n) + 1
    estilizar_corpo(wn, 2, last_n, 11, col_texto=(2,))
    for r in range(2, last_n + 1):
        wn.cell(r, 9).number_format = "#,##0.00"
        wn.cell(r, 11).alignment = QUEBRA
        origem = wn.cell(r, 1).value
        motivo = wn.cell(r, 11).value or ""
        if origem == "Só em Transações":
            fill = FILL_WARN
        elif "Terminal diferente" not in motivo and "Posterior" not in motivo:
            fill = FILL_NO
        else:
            fill = None
        if fill:
            for c in range(1, 12):
                wn.cell(r, c).fill = fill
    wn.cell(1, 6).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wn.row_dimensions[1].height = 30
    wn.freeze_panes = "C2"
    wn.auto_filter.ref = f"A1:K{last_n}"

    # =====================================================================
    # ABA: Sem_Autorizacao
    # =====================================================================
    wsa = wb.create_sheet("Sem_Autorizacao")
    hdr(wsa, 1, ["Linha original", "Data", "Cartão", "PDV", "Tipo Produto", "Descrição Transação",
                 "Valor (R$)", "Estado Transação"],
        [13, 20, 20, 11, 14, 28, 13, 20])
    for i, t in enumerate(tr_sem, 2):
        wsa.cell(i, 1, t["linha"]); wsa.cell(i, 2, t["data"]); wsa.cell(i, 3, t["cartao"]); wsa.cell(i, 4, t["pdv"])
        wsa.cell(i, 5, t["tipo"]); wsa.cell(i, 6, t["desc"]); wsa.cell(i, 7, t["valor_n"]); wsa.cell(i, 8, t["estado"])
    last_sa = len(tr_sem) + 1
    estilizar_corpo(wsa, 2, last_sa, 8)
    for r in range(2, last_sa + 1):
        wsa.cell(r, 7).number_format = "#,##0.00"
    wsa.freeze_panes = "A2"
    wsa.auto_filter.ref = f"A1:H{last_sa}"

    # =====================================================================
    # ABA: Resumo
    # =====================================================================
    wsu = wb.create_sheet("Resumo", 0)
    wsu.column_dimensions["A"].width = 58
    wsu.column_dimensions["B"].width = 16
    wsu.column_dimensions["C"].width = 62
    wsu["A1"] = "CRUZAMENTO DE AUTORIZAÇÃO — TRANSAÇÕES × BANCO"
    wsu["A1"].font = F_TITLE
    wsu["A2"] = ('Transações: coluna "Autorização"  |  Banco: coluna "AUTHO"  '
                 "(comparação exata, texto, zeros à esquerda preservados)")
    wsu["A2"].font = F_NOTE

    def sec(row, titulo):
        for c in range(1, 4):
            wsu.cell(row, c).fill = FILL_SUB
            wsu.cell(row, c).border = BOX
        wsu.cell(row, 1, titulo).font = F_BOLD

    def linha_resumo(row, rotulo, valor, nota="", negrito=False, fill=None, fmt="#,##0"):
        a = wsu.cell(row, 1, rotulo)
        b = wsu.cell(row, 2, valor)
        c = wsu.cell(row, 3, nota)
        for x in (a, b, c):
            x.font = F_BOLD if negrito else F_NORM
            x.border = BOX
        b.number_format = fmt
        b.alignment = Alignment(horizontal="right")
        c.font = F_NOTE
        c.alignment = QUEBRA
        if fill:
            for x in (a, b, c):
                x.fill = fill

    n_com_auth = len(tr_rows)
    n_sem_auth = len(tr_sem)
    n_banco = len(bk_rows)
    n_relacionadas = len(relacionadas_auth)
    n_so_tr = len(so_transacoes_auth)
    n_so_bk = len(so_banco_auth)
    n_outro_terminal = sum(1 for a in so_banco_auth if motivo_so_banco(bk_by[a], terminal_principal, data_corte_tupla)[0] == 2)
    n_posterior = sum(1 for a in so_banco_auth if motivo_so_banco(bk_by[a], terminal_principal, data_corte_tupla)[0] == 0)
    n_sem_corresp = sum(1 for a in so_banco_auth if motivo_so_banco(bk_by[a], terminal_principal, data_corte_tupla)[0] == 1)
    n_relac_sem_doc = sum(1 for a in relacionadas_auth if not doc_txt(bk_by[a]["par"]))
    n_divergentes = sum(
        1 for a in relacionadas_auth
        if round((tr_by[a]["valor_n"] or 0) - (bk_by[a]["valor"] or 0), 2) != 0
    )

    r = 4
    sec(r, "TRANSAÇÕES"); r += 1
    linha_resumo(r, "Total de linhas no relatório", n_com_auth + n_sem_auth); r += 1
    linha_resumo(r, "   Com número de autorização", n_com_auth, "Universo usado no cruzamento"); r += 1
    linha_resumo(r, "   Sem número de autorização", n_sem_auth, "Transações negadas ou sem autorização — ver aba Sem_Autorizacao"); r += 2

    sec(r, 'BANCO (aba "Banco")'); r += 1
    linha_resumo(r, "Total de registros", n_banco); r += 2

    sec(r, "RESULTADO DO CRUZAMENTO"); r += 1
    linha_resumo(r, "✅ Autorizações presentes nas DUAS bases", n_relacionadas, "Aba Relacionadas", negrito=True, fill=FILL_OK); r += 1
    linha_resumo(r, "❌ Só em Transações (sem correspondência no Banco)", n_so_tr, "Aba Nao_Relacionadas — requer investigação", negrito=True, fill=FILL_WARN); r += 1
    linha_resumo(r, "❌ Só no Banco (sem correspondência em Transações)", n_so_bk, "Aba Nao_Relacionadas — ver detalhamento abaixo", negrito=True, fill=FILL_NO); r += 2

    sec(r, 'DETALHAMENTO DOS "SÓ NO BANCO"'); r += 1
    linha_resumo(r, "   Outro terminal (fora do escopo do relatório)", n_outro_terminal,
                 f'Terminal principal identificado: {terminal_principal or "—"}'); r += 1
    linha_resumo(r, "   Mesmo terminal, mas posterior ao período do relatório", n_posterior,
                 "O relatório de Transações e o Banco podem cobrir períodos ligeiramente diferentes"); r += 1
    linha_resumo(r, "   Mesmo terminal, dentro do período, sem correspondência", n_sem_corresp,
                 "Divergência real a investigar", negrito=True); r += 1
    linha_resumo(r, "   Relacionadas sem nº de documento (Banco pendente)", n_relac_sem_doc,
                 "Autorização confere, mas o pagamento ainda está pendente no Banco, sem venda vinculada"); r += 2

    sec(r, "CONFERÊNCIAS"); r += 1
    ok1 = "OK" if n_relacionadas + n_so_tr == n_com_auth else "ERRO"
    ok2 = "OK" if n_relacionadas + n_so_bk == n_banco else "ERRO"
    ok3 = "OK" if n_outro_terminal + n_posterior + n_sem_corresp == n_so_bk else "ERRO"
    linha_resumo(r, "Soma: relacionadas + só Transações = com autorização", ok1, fmt="@"); r += 1
    linha_resumo(r, "Soma: relacionadas + só Banco = total Banco", ok2, fmt="@"); r += 1
    linha_resumo(r, 'Soma do detalhamento = total "só no Banco"', ok3, fmt="@"); r += 1
    linha_resumo(r, "Relacionadas com valor divergente", n_divergentes,
                 "Comparação entre o valor de Transações e o valor do Banco para a mesma autorização. "
                 "Divergências ficam destacadas em vermelho na aba Relacionadas.",
                 negrito=(n_divergentes > 0), fill=FILL_WARN if n_divergentes > 0 else None); r += 2

    sec(r, "NOTAS METODOLÓGICAS"); r += 1
    notas = [
        'A chave é o texto exato da autorização. Zeros à esquerda (ex.: "026172") são preservados.',
        "Nº Nota/Cupom/Recibo vem da aba Vendas, ligada ao Banco pelo par PAR_VENDA ↔ PAR_BANCO — "
        "não é a coluna \"Cupom Fiscal\" de Transações, que é um número interno do TEF.",
        "Quando um pagamento do Banco cobre 2 recibos, os dois aparecem juntos (ex.: "
        '"Recibo 47221 + Recibo 47222").',
        "Autorizações sem nº de documento: pendentes no Banco, ainda sem venda vinculada.",
    ]
    for nota in notas:
        wsu.cell(r, 1, "• " + nota).font = F_NOTE
        wsu.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        wsu.cell(r, 1).alignment = QUEBRA
        wsu.row_dimensions[r].height = 26
        r += 1

    ordem_abas = ["Resumo", "Relacionadas", "Nao_Relacionadas", "Sem_Autorizacao",
                  "Base_Transacoes", "Base_Banco", "Base_Vendas"]
    wb._sheets = [wb[n] for n in ordem_abas]
    wb.active = 0
    wb.save(caminho_saida)

    return {
        "com_autorizacao": n_com_auth,
        "sem_autorizacao": n_sem_auth,
        "banco": n_banco,
        "relacionadas": n_relacionadas,
        "so_transacoes": n_so_tr,
        "so_banco": n_so_bk,
        "divergentes": n_divergentes,
    }


# ----------------------------------------------------------------------
# Interface gráfica (Tkinter)
# ----------------------------------------------------------------------

class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master, bg=theme.BG_APP)

        self.caminho_transacoes = tk.StringVar()
        self.caminho_conciliacao = tk.StringVar()

        self._montar_interface()

    def _montar_interface(self):
        tk.Label(self, text="Conciliador de Autorização", font=("Segoe UI", 14, "bold"),
                 bg=theme.BG_APP, fg=theme.FG_TEXT).pack(pady=(18, 4))
        tk.Label(
            self,
            text='Cruza a coluna "Autorização" do relatório de Transações\n'
                 'com a coluna "AUTHO" da aba "Banco" do arquivo de conciliação.',
            font=("Segoe UI", 9), bg=theme.BG_APP, fg=theme.FG_MUTED, justify="center",
        ).pack(pady=(0, 12))

        frame = tk.Frame(self, bg=theme.BG_APP)
        frame.pack(fill="x", padx=16, pady=8)
        frame.columnconfigure(1, weight=1)

        self._linha_arquivo(frame, 0, "Relatório de Transações:", self.caminho_transacoes)
        self._linha_arquivo(frame, 1, "Arquivo de Conciliação:", self.caminho_conciliacao)

        self.botao_gerar = tk.Button(
            self, text="Gerar relatório", font=("Segoe UI", 11, "bold"),
            bg=theme.PRIMARY, fg="white", activebackground=theme.PRIMARY_HOVER,
            activeforeground="white", relief="flat", bd=0, cursor="hand2",
            padx=24, pady=9, command=self._ao_clicar_gerar,
        )
        self.botao_gerar.pack(pady=18)

        self.barra_progresso = ttk.Progressbar(self, mode="indeterminate", length=420)
        self.label_status = tk.Label(self, text="", font=("Segoe UI", 9),
                                     bg=theme.BG_APP, fg=theme.FG_MUTED)
        self.label_status.pack()

    def _linha_arquivo(self, frame, linha, rotulo, variavel):
        tk.Label(frame, text=rotulo, font=("Segoe UI", 10), bg=theme.BG_APP,
                 fg=theme.FG_TEXT).grid(row=linha, column=0, sticky="w", pady=6)
        tk.Entry(frame, textvariable=variavel, font=("Segoe UI", 9), state="readonly",
                 readonlybackground=theme.SURFACE2, fg=theme.FG_TEXT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=theme.GRID_LINE).grid(
            row=linha, column=1, sticky="ew", padx=8, ipady=4)
        tk.Button(frame, text="Procurar...", font=("Segoe UI", 9), bg=theme.SURFACE2,
                  fg=theme.FG_TEXT, activebackground=theme.GRID_LINE, activeforeground="white",
                  relief="flat", bd=0, cursor="hand2", padx=12, pady=4,
                  command=lambda: self._escolher_arquivo(variavel)).grid(row=linha, column=2)

    def _escolher_arquivo(self, variavel):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo", filetypes=[("Arquivos Excel", "*.xlsx *.xlsm")])
        if caminho:
            variavel.set(caminho)

    def _ao_clicar_gerar(self):
        caminho_t = self.caminho_transacoes.get().strip()
        caminho_c = self.caminho_conciliacao.get().strip()

        if not caminho_t or not caminho_c:
            messagebox.showwarning("Arquivos faltando", "Selecione os dois arquivos antes de gerar o relatório.")
            return

        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar relatório como", defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")], initialfile="cruzamento_autorizacao.xlsx",
        )
        if not caminho_saida:
            return

        self.botao_gerar.config(state="disabled")
        self.barra_progresso.pack(pady=(0, 8))
        self.barra_progresso.start(10)
        self.label_status.config(text="Processando...")

        threading.Thread(target=self._gerar_em_segundo_plano, args=(caminho_t, caminho_c, caminho_saida)).start()

    def _gerar_em_segundo_plano(self, caminho_t, caminho_c, caminho_saida):
        try:
            resultado = gerar_relatorio(caminho_t, caminho_c, caminho_saida)
            self.after(0, self._ao_terminar_sucesso, resultado, caminho_saida)
        except Exception as e:
            self.after(0, self._ao_terminar_erro, str(e))

    def _ao_terminar_sucesso(self, resultado, caminho_saida):
        self.barra_progresso.stop()
        self.barra_progresso.pack_forget()
        self.botao_gerar.config(state="normal")
        self.label_status.config(text=f"Relatório salvo em: {os.path.basename(caminho_saida)}")
        messagebox.showinfo(
            "Concluído",
            "Relatório gerado com sucesso!\n\n"
            f"Relacionadas: {resultado['relacionadas']}\n"
            f"Só em Transações: {resultado['so_transacoes']}\n"
            f"Só no Banco: {resultado['so_banco']}\n"
            f"Valores divergentes: {resultado['divergentes']}\n\n"
            f"Salvo em:\n{caminho_saida}",
        )

    def _ao_terminar_erro(self, mensagem_erro):
        self.barra_progresso.stop()
        self.barra_progresso.pack_forget()
        self.botao_gerar.config(state="normal")
        self.label_status.config(text="Erro ao gerar o relatório.")
        messagebox.showerror("Erro", f"Não foi possível gerar o relatório:\n\n{mensagem_erro}")


def _build_content(win) -> None:
    win.title("Conciliador de Autorização")
    win.geometry("580x320")
    win.minsize(540, 300)
    win.configure(bg=theme.BG_APP)
    theme.apply_theme(win)
    App(win).pack(fill="both", expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()
