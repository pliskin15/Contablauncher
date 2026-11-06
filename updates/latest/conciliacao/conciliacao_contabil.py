"""
conciliacao_contabil.py
=======================
Sistema de conciliação contábil — PMZ Peças e Pneus

ETAPA 1 — Contabilidade (XLS) × Relatórios (NF + Cupom Fiscal)
  · Cupons : agrupa VENDA PIX MAQUINETA por IMPR. + DATA → cruza com DOC. NRO. + DATA do razão
             (IMPR. 011 → DOC 11, por dia)
  · Notas  : cada NF individual → cruza número da nota com DOC. NRO. do razão, por dia

ETAPA 2 — Recibos (PDF) × Pagamentos Getnet (XLSX)
  · Cruzamento por valor + data (lógica conciliar_automatico do sistema original)

Parsers reutilizados do pix_maquineta.py:
  ler_cupom_fiscal, ler_nota_fiscal, ler_recibos, ler_mov_pix, ler_razao
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import os, re
import datetime as _dt
from theme import T, aplicar_estilos_ttk, aplicar_tags_tree, botao_tema, registrar_callback

try:
    from tkcalendar import DateEntry
    _TEM_TKCALENDAR = True
except ImportError:
    _TEM_TKCALENDAR = False


# ─────────────────────────────────────────────────────────────────────────────
# CORES
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS (copiados do pix_maquineta.py)
# ─────────────────────────────────────────────────────────────────────────────

def _num(val):
    if val is None:
        return 0.0
    s = str(val).strip()
    if re.search(r"\d\.\d{3},\d", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except Exception:
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


def _normalizar_data(val):
    import datetime as _dt
    if isinstance(val, _dt.datetime):
        return val.date()
    if isinstance(val, _dt.date):
        return val
    try:
        if isinstance(val, pd.Timestamp):
            return val.date()
    except Exception:
        pass
    s = str(val).strip().split(" ")[0].split("T")[0]
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS — copiados/adaptados do pix_maquineta.py
# ─────────────────────────────────────────────────────────────────────────────

def ler_cupom_fiscal(path):
    """
    Lê o Relatório Cupom Fiscal (PDF).
    Captura cupons com VENDA PIX MAQUINETA individualmente,
    mantendo IMPR. e DATA para agrupamento posterior.
    """
    linhas = _extrair_linhas_pdf(path)
    registros = []

    cupom_atual  = None
    impr_atual   = None
    cond_atual   = ""
    vender_atual = ""
    data_atual   = ""

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue
        up = linha_strip.upper()

        m_data = re.search(r"DATA\s*:\s*(\d{2}/\d{2}/\d{4})", up)
        if m_data:
            data_atual = m_data.group(1)
            continue

        # Linha principal do cupom: captura também o IMPR.
        # Ex: "24829 24829 011 NORMAL CD 5608 5.405 15,00 ..."
        m = re.match(r"^(\d{5,})\s+\d{5,}\s+(\d+)\s+\w+\s+(\S+)\s+(\S+)", linha_strip)
        if m:
            cupom_atual  = m.group(1)
            impr_atual   = str(int(m.group(2)))   # remove zeros à esquerda → "011" vira "11"
            cond_atual   = m.group(3)
            vender_atual = m.group(4)
            continue

        if any(up.startswith(t) for t in ["TOTAL GERAL", "TOTAL :", "TOTAL:",
                                           "DESCRICAO", "TOTAIS", "CANCELADOS",
                                           "SERVICOS", "VENDAS"]):
            cupom_atual = None
            impr_atual  = None
            continue

        if "VENDA PIX" in up and "MAQUINETA" in up:
            if not cupom_atual or not impr_atual:
                continue
            valor = _ultimo_num_linha(linha_strip)
            if valor > 0:
                registros.append({
                    "cupom":      cupom_atual,
                    "impr":       impr_atual,
                    "data":       data_atual,
                    "valor":      round(valor, 2),
                    "cond":       cond_atual,
                    "vendedor":   vender_atual,
                })
                cupom_atual = None
                impr_atual  = None

    return pd.DataFrame(registros) if registros else pd.DataFrame(
        columns=["cupom","impr","data","valor","cond","vendedor"])


def ler_nota_fiscal(path):
    """
    Lê o Relatório de Venda Avista — Notas Fiscais (PDF).
    Captura APENAS os valores de VENDA PIX MAQUINETA por nota.
    Notas sem nenhuma linha de PIX MAQUINETA são ignoradas.
    """
    linhas = _extrair_linhas_pdf(path)
    registros = []

    nota_atual    = None
    cliente_atual = ""
    cond_atual    = ""
    pix_maq_atual = 0.0   # acumulador só de PIX MAQUINETA desta nota
    data_atual    = ""

    def _parse_linha_nota(ls):
        """
        Detecta linha de cabeçalho de nota pelo padrão posicional:
        NOTA SERIE FILIAL FILIAL_VEND COD_CLIENTE NOME... COND TOTAL ...
        Os primeiros 5 tokens devem ser numéricos; o primeiro valor decimal
        é o TOTAL da nota; o token imediatamente anterior é a COND.
        Retorna (nota, cliente, cond) ou None se não for linha de nota.
        """
        partes = ls.split()
        if len(partes) < 7:
            return None
        # Primeiros 5 tokens devem ser inteiros (nota, serie, filial, ...)
        if not all(re.match(r"^\d+$", partes[i]) for i in range(5)):
            return None
        # Primeira posição com valor decimal  (ex: "142,52")
        num_pos = [i for i, p in enumerate(partes)
                   if re.match(r"^\d{1,3}(?:\.\d{3})*,\d{2}$|^\d+,\d{2}$", p)]
        if not num_pos or num_pos[0] < 6:
            return None
        idx_total = num_pos[0]
        cond    = partes[idx_total - 1]
        cliente = " ".join(partes[5:idx_total - 1])
        return partes[0], cliente, cond   # nota, cliente, cond

    def _salvar_nota():
        nonlocal nota_atual, cliente_atual, cond_atual, pix_maq_atual
        if nota_atual and pix_maq_atual > 0:
            registros.append({
                "nota":    nota_atual,
                "data":    data_atual,
                "valor":   round(pix_maq_atual, 2),
                "cliente": cliente_atual,
                "cond":    cond_atual,
            })
        nota_atual    = None
        cliente_atual = ""
        cond_atual    = ""
        pix_maq_atual = 0.0

    for linha in linhas:
        linha_strip = linha.strip()
        if not linha_strip:
            continue
        up = linha_strip.upper()

        # Data isolada como cabeçalho de seção  (ex: "30/05/2026")
        if re.match(r"^\d{2}/\d{2}/\d{4}$", linha_strip):
            data_atual = linha_strip
            continue

        # Fim do bloco de um dia — salva nota pendente
        if up.startswith("TOTAL DO DIA") or up.startswith("TOTAL GERAL"):
            _salvar_nota()
            continue

        # Ignora cabeçalhos, rodapés e linhas de totais do relatório
        if any(up.startswith(t) for t in ("NOTA ", "RUA ", "EMAIL", "C.N.P.J",
                                           "PAGINA", "PÁGINA", "FILIAL:", "CODPOS",
                                           "TOTAL ICMS", "TOTAL DE VENDAS",
                                           "DESCRICAO", "DINHEIRO", "CARTAO",
                                           "CHEQUE", "DEPOSITO", "PIX ", "PIX\t",
                                           "OUTROS", "RELATORIO", "NOTAS FISCAIS",
                                           "EMISSAO", "USUARIO")):
            continue

        # Linha de início de nova nota — parser posicional
        parsed = _parse_linha_nota(linha_strip)
        if parsed:
            _salvar_nota()
            nota_atual, cliente_atual, cond_atual = parsed
            continue

        # Linha de pagamento PIX MAQUINETA — acumula valor
        if "VENDA PIX MAQUINETA" in up and nota_atual:
            valor = _ultimo_num_linha(linha_strip)
            if valor > 0:
                pix_maq_atual = round(pix_maq_atual + valor, 2)
            continue

    # Última nota do arquivo
    _salvar_nota()

    return pd.DataFrame(registros) if registros else pd.DataFrame(
        columns=["nota","data","valor","cliente","cond"])


def ler_recibos(path):
    """
    Lê o Relatório de Recibos (PDF) — parser original do pix_maquineta.py.
    Captura recibos com DEP. GETNET PIX / DEP. PIX MAQUINETA.
    """
    linhas = _extrair_linhas_pdf(path)
    registros    = []
    recibo_atual = None
    data_atual   = ""

    for linha in linhas:
        linha_limpa = linha.strip()
        up = linha_limpa.upper()

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

            if "DEP. PIX MAQUINETA" in up or "DEP. GETNET PIX" in up:
                partes = linha_limpa.split()
                numero_recibo = partes[1]
                recibo_atual = {
                    "numero":  numero_recibo,
                    "data":    data_atual,
                    "valor":   0.0,
                    "status":  "pendente",
                    "par":     "",
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

    df = pd.DataFrame(registros) if registros else pd.DataFrame(
        columns=["numero","data","valor","status","par"])
    if not df.empty:
        df["saldo_rest"] = df["valor"]
    return df


def ler_mov_pix(path):
    """
    Lê extrato PIX Maquineta XLSX (Getnet/Santander) — parser original.
    Apenas registros com STATUS == 'Paga'.
    """
    try:
        df_raw = pd.read_excel(path, sheet_name="PIX", header=7, engine="openpyxl")
    except Exception:
        df_raw = pd.read_excel(path, header=7, engine="openpyxl")

    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    col_status = "STATUS"
    if col_status in df_raw.columns:
        df_raw = df_raw[df_raw[col_status].astype(str).str.strip().str.lower() == "paga"]

    col_valor = "VALOR DA VENDA"
    df_raw = df_raw[pd.to_numeric(df_raw[col_valor], errors="coerce") > 0].copy()
    df_raw[col_valor] = pd.to_numeric(df_raw[col_valor], errors="coerce").round(2)

    registros = []
    for _, row in df_raw.iterrows():
        dt_hora  = str(row.get("DATA/HORA DA VENDA", "")).strip()
        txid     = str(row.get("ID/TRANSAÇÃO (ID)", "")).strip()
        cv       = str(row.get("NÚMERO DO COMPROVANTE DE VENDAS (CV)", "")).strip()
        terminal = str(row.get("NÚMERO DO TERMINAL", "")).strip()
        valor    = float(row[col_valor])
        partes_dt = dt_hora.split(" ")
        dt   = partes_dt[0] if partes_dt else ""
        hora = partes_dt[1] if len(partes_dt) > 1 else ""
        registros.append({
            "DT_RECEB":   dt,
            "HR_RECEB":   hora,
            "TERMINAL":   terminal,
            "CV":         cv,
            "TXID":       txid,
            "VALOR":      valor,
            "status":     "pendente",
            "par_recibo": "",
            "saldo_rest": valor,
        })

    COLS = ["DT_RECEB","HR_RECEB","TERMINAL","CV","TXID","VALOR","status","par_recibo","saldo_rest"]
    if not registros:
        df = pd.DataFrame(columns=COLS)
        df["VALOR"]      = pd.Series(dtype=float)
        df["saldo_rest"] = pd.Series(dtype=float)
    else:
        df = pd.DataFrame(registros, columns=COLS)
    return df


def ler_razao(path):
    """
    Lê o Razão Modelo I (XLS) — parser original do pix_maquineta.py.
    Filtra CONTA PARTIDA *-2074. Mantém doc_nro e data por lançamento.
    """
    try:
        df_raw = pd.read_excel(path, engine="xlrd", header=None)
    except Exception:
        df_raw = pd.read_excel(path, header=None)

    registros   = []
    ano_atual   = ""
    conta_atual = ""

    for i, row in df_raw.iterrows():
        r = [str(v).strip() if v is not None else "" for v in row]

        if str(r[0]).upper().startswith("PERIODO"):
            m_ano = re.search(r"(\d{4})", r[1])
            if m_ano:
                ano_atual = m_ano.group(1)
            conta_atual = r[7] if len(r) > 7 else ""
            continue

        col_cp = r[7]
        if "2074" not in col_cp:
            continue

        data_str = r[0]
        m_data = re.match(r"^(\d{1,2})/(\d{2})$", data_str)
        if not m_data:
            continue

        dia = m_data.group(1).zfill(2)
        mes = m_data.group(2)
        data_completa = f"{dia}/{mes}/{ano_atual}" if ano_atual else data_str

        try:
            credito = float(str(r[9]).replace(",", ".")) if r[9] not in ("", "nan") else 0.0
        except ValueError:
            credito = 0.0
        try:
            debito = float(str(r[8]).replace(",", ".")) if r[8] not in ("", "nan") else 0.0
        except ValueError:
            debito = 0.0

        valor = credito if credito > 0 else debito
        if valor <= 0:
            continue

        doc_nro = str(r[4]).strip()
        if not doc_nro or doc_nro in ("nan", ""):
            continue

        registros.append({
            "data":          data_completa,
            "sequencia":     r[1],
            "voucher":       r[3],
            "doc_nro":       doc_nro,
            "historico":     r[5],
            "conta_partida": col_cp,
            "valor":         round(valor, 2),
        })

    return pd.DataFrame(registros) if registros else pd.DataFrame(
        columns=["data","sequencia","voucher","doc_nro","historico","conta_partida","valor"])


# ─────────────────────────────────────────────────────────────────────────────
# CONCILIAÇÃO — ETAPA 1
# ─────────────────────────────────────────────────────────────────────────────

def conciliar_etapa1(df_razao, df_cupom, df_nf, tolerancia=0.01):
    """
    Cupons:
      Agrupa df_cupom por (impr + data), soma valores.
      Para cada grupo, busca no razão as linhas com doc_nro == impr e mesma data.
      Soma os créditos do razão para aquele doc_nro + data.

    Notas Fiscais:
      Para cada nota (nota + data), busca no razão doc_nro == nota e mesma data.
      Soma os créditos do razão para aquela nota + data.

    Retorna (df_resultado_cupom, df_resultado_nf) — um registro por lote/nota por dia.
    """
    TOL = tolerancia

    def _doc_num(d):
        try:
            return str(int(float(str(d).strip())))
        except Exception:
            return str(d).strip()

    df_r = df_razao.copy()
    df_r["_doc_num"] = df_r["doc_nro"].apply(_doc_num)
    df_r["_data_norm"] = df_r["data"].apply(_normalizar_data)

    # ── CUPONS ────────────────────────────────────────────────────────────────
    result_cupom = []

    if df_cupom is not None and not df_cupom.empty:
        df_c = df_cupom.copy()
        df_c["_data_norm"] = df_c["data"].apply(_normalizar_data)

        # Agrupa por impressora + data
        grp = (df_c.groupby(["impr", "data", "_data_norm"])
               .agg(qtd_cupons=("valor", "count"),
                    val_relatorio=("valor", "sum"))
               .reset_index())

        for _, row in grp.iterrows():
            impr      = str(row["impr"])           # já sem zeros
            data_str  = row["data"]
            data_norm = row["_data_norm"]
            val_rel   = round(row["val_relatorio"], 2)
            qtd       = int(row["qtd_cupons"])

            # Busca no razão: doc_nro == impr E mesma data
            match = df_r[
                (df_r["_doc_num"] == impr) &
                (df_r["_data_norm"] == data_norm)
            ]

            if match.empty:
                val_cont = 0.0
                status   = "sem_match"
                seq_list = ""
            else:
                val_cont = round(match["valor"].sum(), 2)
                seq_list = ", ".join(match["sequencia"].astype(str).tolist())
                dif = round(val_rel - val_cont, 2)
                status = "conciliado" if abs(dif) <= TOL else "divergente"

            dif = round(val_rel - val_cont, 2)
            result_cupom.append({
                "impr":          impr.zfill(3),
                "doc_nro":       impr,
                "data":          data_str,
                "qtd_cupons":    qtd,
                "val_relatorio": val_rel,
                "val_contabil":  val_cont,
                "diferenca":     dif if val_cont > 0 else None,
                "sequencias":    seq_list,
                "status":        status,
            })

        # Lotes do razão sem match no relatório de cupons (doc_nro curto)
        docs_rel = set(
            (str(r["impr"]), r["_data_norm"])
            for _, r in df_c.iterrows()
        )
        razao_curtos = df_r[df_r["_doc_num"].str.match(r"^\d{1,4}$")]
        grp_r = (razao_curtos.groupby(["_doc_num", "_data_norm", "data"])
                 .agg(val_contabil=("valor", "sum"),
                      seq_list=("sequencia", lambda x: ", ".join(x.astype(str))))
                 .reset_index())
        for _, row in grp_r.iterrows():
            chave = (row["_doc_num"], row["_data_norm"])
            if chave not in docs_rel:
                result_cupom.append({
                    "impr":          row["_doc_num"].zfill(3),
                    "doc_nro":       row["_doc_num"],
                    "data":          row["data"],
                    "qtd_cupons":    0,
                    "val_relatorio": 0.0,
                    "val_contabil":  round(row["val_contabil"], 2),
                    "diferenca":     None,
                    "sequencias":    row["seq_list"],
                    "status":        "sem_match",
                })

    df_res_cupom = pd.DataFrame(result_cupom) if result_cupom else pd.DataFrame(
        columns=["impr","doc_nro","data","qtd_cupons","val_relatorio",
                 "val_contabil","diferenca","sequencias","status"])

    # ── NOTAS FISCAIS ─────────────────────────────────────────────────────────
    result_nf = []

    if df_nf is not None and not df_nf.empty:
        df_n = df_nf.copy()
        df_n["_data_norm"] = df_n["data"].apply(_normalizar_data)

        for _, row in df_n.iterrows():
            nota      = str(row["nota"]).strip()
            data_str  = row["data"]
            data_norm = row["_data_norm"]
            val_rel   = round(row["valor"], 2)

            match = df_r[
                (df_r["_doc_num"] == nota) &
                (df_r["_data_norm"] == data_norm)
            ]

            if match.empty:
                val_cont = 0.0
                status   = "sem_match"
                seq_list = ""
            else:
                val_cont = round(match["valor"].sum(), 2)
                seq_list = ", ".join(match["sequencia"].astype(str).tolist())
                dif = round(val_rel - val_cont, 2)
                status = "conciliado" if abs(dif) <= TOL else "divergente"

            dif = round(val_rel - val_cont, 2)
            result_nf.append({
                "nota":          nota,
                "doc_nro":       nota,
                "cliente":       row.get("cliente", ""),
                "cond":          row.get("cond", ""),
                "data":          data_str,
                "val_relatorio": val_rel,
                "val_contabil":  val_cont,
                "diferenca":     dif if val_cont > 0 else None,
                "sequencias":    seq_list,
                "status":        status,
            })

        # Notas do razão sem match no relatório
        notas_rel = set(
            (str(r["nota"]), r["_data_norm"])
            for _, r in df_n.iterrows()
        )
        razao_longos = df_r[df_r["_doc_num"].str.match(r"^\d{5,}$")]
        grp_rn = (razao_longos.groupby(["_doc_num", "_data_norm", "data"])
                  .agg(val_contabil=("valor", "sum"),
                       seq_list=("sequencia", lambda x: ", ".join(x.astype(str))))
                  .reset_index())
        for _, row in grp_rn.iterrows():
            chave = (row["_doc_num"], row["_data_norm"])
            if chave not in notas_rel:
                result_nf.append({
                    "nota":          row["_doc_num"],
                    "doc_nro":       row["_doc_num"],
                    "cliente":       "",
                    "cond":          "",
                    "data":          row["data"],
                    "val_relatorio": 0.0,
                    "val_contabil":  round(row["val_contabil"], 2),
                    "diferenca":     None,
                    "sequencias":    row["seq_list"],
                    "status":        "sem_match",
                })

    df_res_nf = pd.DataFrame(result_nf) if result_nf else pd.DataFrame(
        columns=["nota","doc_nro","cliente","cond","data","val_relatorio",
                 "val_contabil","diferenca","sequencias","status"])

    return df_res_cupom, df_res_nf


# ─────────────────────────────────────────────────────────────────────────────
# CONCILIAÇÃO — ETAPA 2 (valor + data, lógica original)
# ─────────────────────────────────────────────────────────────────────────────

def conciliar_etapa2(df_recibos, df_pagamentos, df_res_cupom=None, df_res_nf=None,
                     tolerancia=0.01):
    """
    Etapa 2 — dois passos:

    PASSO 1 — Recibos × Getnet (por valor + data, lógica original).
      Cada recibo é cruzado com um pagamento Getnet de mesmo valor e data.

    PASSO 2 — Check diário: (val_contábil cupons+NF) + recibos_conciliados
                            vs pagamentos Getnet restantes do dia.
      Os pagamentos que não foram consumidos pelo passo 1 são agrupados por dia
      e comparados com a soma contábil (cupons + NF) daquele mesmo dia.
      O resultado é uma tabela de diferenças por data.

    Retorna (dr, dp, df_check_diario).
    """
    TOL = tolerancia

    # ── PASSO 1: recibos × getnet ─────────────────────────────────────────────
    dr = df_recibos.copy()
    dp = df_pagamentos.copy()

    dr["status"]     = "pendente"
    dr["par"]        = ""
    dr["saldo_rest"] = dr["valor"]
    dp["status"]     = "pendente"
    dp["par_recibo"] = ""
    dp["saldo_rest"] = dp["VALOR"]

    dp["_data_norm"] = dp["DT_RECEB"].apply(_normalizar_data)
    dr["_data_norm"] = dr["data"].apply(_normalizar_data)

    par_counter = [0]

    def novo_par():
        par_counter[0] += 1
        return f"E2P{par_counter[0]:04d}"

    def _ap(df, idx, col, p):
        a = str(df.at[idx, col])
        df.at[idx, col] = f"{a},{p}" if a and a != "nan" else p

    for ir, row_r in dr.iterrows():
        saldo_r = dr.at[ir, "saldo_rest"]
        if saldo_r <= TOL:
            continue
        data_r = dr.at[ir, "_data_norm"]
        candidatos = dp[
            (dp["saldo_rest"] > TOL) &
            (abs(dp["VALOR"] - saldo_r) <= TOL) &
            (dp["_data_norm"] == data_r)
        ]
        for ip, _ in candidatos.iterrows():
            saldo_r = dr.at[ir, "saldo_rest"]
            saldo_p = dp.at[ip, "saldo_rest"]
            if saldo_r <= TOL or saldo_p <= TOL:
                continue
            valor_match = min(saldo_r, saldo_p)
            par = novo_par()
            _ap(dr, ir, "par", par)
            dr.at[ir, "saldo_rest"] = round(saldo_r - valor_match, 2)
            _ap(dp, ip, "par_recibo", par)
            dp.at[ip, "saldo_rest"] = round(saldo_p - valor_match, 2)
            break

    def _status_r(row):
        if not row["par"]:
            return "pendente"
        return "conciliado" if row["saldo_rest"] <= TOL else "parcial"

    def _status_p(row):
        if not row["par_recibo"]:
            return "pendente"
        return "conciliado" if row["saldo_rest"] <= TOL else "parcial"

    # ── PASSO 1B: combinações em grupo, por dia ─────────────────────────────────
    # O passo 1 só casa recibo×pagamento de valor IDÊNTICO (1:1). Aqui tentamos:
    #   (a) N recibos cuja soma bate com 1 pagamento pendente
    #   (b) 1 recibo cuja soma bate com N pagamentos pendentes
    #   (c) N recibos × N pagamentos: quando o total pendente do dia bate dos
    #       dois lados, concilia tudo junto (cobre o restante dos casos N:N)
    import itertools

    MAX_ITENS_COMBO = 6  # evita explosão combinatória em dias com muitos itens pendentes

    def _grupo_par(idxs_r, idxs_p, valor_match):
        par = novo_par()
        resto_r = valor_match
        for ir2 in idxs_r:
            saldo = dr.at[ir2, "saldo_rest"]
            if saldo <= TOL or resto_r <= TOL:
                continue
            consumido = min(saldo, resto_r)
            _ap(dr, ir2, "par", par)
            dr.at[ir2, "saldo_rest"] = round(saldo - consumido, 2)
            resto_r = round(resto_r - consumido, 2)
        resto_p = valor_match
        for ip2 in idxs_p:
            saldo = dp.at[ip2, "saldo_rest"]
            if saldo <= TOL or resto_p <= TOL:
                continue
            consumido = min(saldo, resto_p)
            _ap(dp, ip2, "par_recibo", par)
            dp.at[ip2, "saldo_rest"] = round(saldo - consumido, 2)
            resto_p = round(resto_p - consumido, 2)

    def _buscar_subset(itens, alvo, tol, max_itens=6):
        """itens: lista de (idx, valor) pendentes. Procura subconjunto cuja soma ≈ alvo."""
        if not itens or alvo <= tol:
            return None
        limite = min(len(itens), max_itens)
        for tam in range(1, limite + 1):
            for combo in itertools.combinations(itens, tam):
                soma = round(sum(v for _, v in combo), 2)
                if abs(soma - alvo) <= tol:
                    return [i for i, _ in combo]
        return None

    for data_norm in sorted({d for d in dr["_data_norm"] if d is not None} |
                             {d for d in dp["_data_norm"] if d is not None}):

        # (a) 1 pagamento pendente × N recibos pendentes
        pags_dia = dp[(dp["_data_norm"] == data_norm) & (dp["saldo_rest"] > TOL)]
        for ip in list(pags_dia.index):
            saldo_p = dp.at[ip, "saldo_rest"]
            if saldo_p <= TOL:
                continue
            recs_dia = dr[(dr["_data_norm"] == data_norm) & (dr["saldo_rest"] > TOL)]
            if len(recs_dia) < 2:
                continue
            itens = [(i, dr.at[i, "saldo_rest"]) for i in recs_dia.index]
            combo = _buscar_subset(itens, saldo_p, TOL, MAX_ITENS_COMBO)
            if combo:
                _grupo_par(combo, [ip], saldo_p)

        # (b) 1 recibo pendente × N pagamentos pendentes
        recs_dia = dr[(dr["_data_norm"] == data_norm) & (dr["saldo_rest"] > TOL)]
        for ir in list(recs_dia.index):
            saldo_r = dr.at[ir, "saldo_rest"]
            if saldo_r <= TOL:
                continue
            pags_dia2 = dp[(dp["_data_norm"] == data_norm) & (dp["saldo_rest"] > TOL)]
            if len(pags_dia2) < 2:
                continue
            itens = [(i, dp.at[i, "saldo_rest"]) for i in pags_dia2.index]
            combo = _buscar_subset(itens, saldo_r, TOL, MAX_ITENS_COMBO)
            if combo:
                _grupo_par([ir], combo, saldo_r)

        # (c) N recibos × N pagamentos — totais pendentes do dia coincidem
        recs_dia = dr[(dr["_data_norm"] == data_norm) & (dr["saldo_rest"] > TOL)]
        pags_dia = dp[(dp["_data_norm"] == data_norm) & (dp["saldo_rest"] > TOL)]
        if len(recs_dia) and len(pags_dia):
            soma_r = round(recs_dia["saldo_rest"].sum(), 2)
            soma_p = round(pags_dia["saldo_rest"].sum(), 2)
            if abs(soma_r - soma_p) <= TOL:
                _grupo_par(list(recs_dia.index), list(pags_dia.index), min(soma_r, soma_p))

    dr["status"] = dr.apply(_status_r, axis=1)
    dp["status"] = dp.apply(_status_p, axis=1)

    # ── PASSO 2: check diário ─────────────────────────────────────────────────
    COLS_CHECK = ["data", "val_recibos", "val_contabil", "val_getnet_rest",
                  "val_contabil_recibos", "diferenca", "status"]

    # Recibos conciliados (total ou parcial) por dia.
    # Para "parcial" soma-se apenas a parte efetivamente conciliada
    # (valor - saldo_rest), não o valor total do recibo.
    rec_conc = dr[dr["status"].isin(["conciliado", "parcial"])].copy()
    rec_conc["_data_norm"] = rec_conc["_data_norm"] if "_data_norm" in rec_conc.columns \
        else rec_conc["data"].apply(_normalizar_data)
    rec_conc["_valor_conciliado"] = rec_conc["valor"] - rec_conc["saldo_rest"]
    rec_por_dia = (rec_conc.groupby("_data_norm")["_valor_conciliado"]
                   .sum().round(2).rename("val_recibos"))

    # Getnet restante (não conciliado com recibos) por dia
    get_rest_por_dia = (dp.groupby("_data_norm")["VALOR"]
                    .sum().round(2).rename("val_getnet_rest"))


    # val_contabil (cupons + NF) por dia — usa val_contabil da Etapa 1
    contabil_por_dia = {}
    for df_e1, col_data in [(df_res_cupom, "data"), (df_res_nf, "data")]:
        if df_e1 is None or df_e1.empty:
            continue
        tmp = df_e1.copy()
        tmp["_data_norm"] = tmp[col_data].apply(_normalizar_data)
        grp = tmp.groupby("_data_norm")["val_contabil"].sum()
        for data_norm, val in grp.items():
            contabil_por_dia[data_norm] = round(
                contabil_por_dia.get(data_norm, 0.0) + val, 2)

    # Une todas as datas
    todas_datas = set(rec_por_dia.index) | set(get_rest_por_dia.index) | set(contabil_por_dia.keys())

    check_rows = []
    for data_norm in sorted(todas_datas):
        val_rec   = round(rec_por_dia.get(data_norm, 0.0), 2)
        val_cont  = round(contabil_por_dia.get(data_norm, 0.0), 2)
        val_grest = round(get_rest_por_dia.get(data_norm, 0.0), 2)
        val_cont_rec = round(val_cont + val_rec, 2)
        dif = round(val_cont_rec - val_grest, 2)
        status = "conciliado" if abs(dif) <= TOL else ("divergente" if dif != 0 else "pendente")
        check_rows.append({
            "data":                 data_norm.strftime("%d/%m/%Y") if data_norm else "",
            "val_recibos":          val_rec,
            "val_contabil":         val_cont,
            "val_getnet_rest":      val_grest,
            "val_contabil_recibos": val_cont_rec,
            "diferenca":            dif,
            "status":               status,
        })

    df_check = (pd.DataFrame(check_rows, columns=COLS_CHECK)
                if check_rows else pd.DataFrame(columns=COLS_CHECK))

    dr.drop(columns=["_data_norm"], inplace=True, errors="ignore")
    dp.drop(columns=["_data_norm"], inplace=True, errors="ignore")
    return dr, dp, df_check


# ─────────────────────────────────────────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────────────────────────────────────────

class ConciliacaoContabil(tk.Toplevel):

    def __init__(self, master=None):
        super().__init__()
        self.title("Conciliação Contábil — PMZ Peças e Pneus")
        self.geometry("1560x900")
        self.configure(bg=COR_BG)
        self.resizable(True, True)
        self.grab_set()

        self._df_razao   = None
        self._df_cupom   = None
        self._df_nf      = None
        self._df_recibos = None
        self._df_getnet  = None

        self.res_cupom   = None
        self.res_nf      = None
        self.res_recibos = None
        self.res_getnet  = None
        self.res_check   = None

        self._build_ui()
        self._aplicar_estilos()

    # ─── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_topbar()
        corpo = tk.Frame(self, bg=COR_BG)
        corpo.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._build_sidebar(corpo)
        self._build_area_principal(corpo)
        self._build_statusbar()

    def _build_topbar(self):
        container = tk.Frame(self, bg=COR_PAINEL)
        container.pack(fill="x")

        # ── Linha 1: título, etapas, ações de processamento, margem, exportar ──
        bar = tk.Frame(container, bg=COR_PAINEL, height=56)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        tk.Label(bar, text="📊  Conciliação Contábil — PMZ",
                 bg=COR_PAINEL, fg=COR_TEXTO,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16, pady=10)

        sep = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep.pack(side="left", fill="y", pady=10, padx=8)

        tk.Label(bar, text="Etapa:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 11)).pack(side="left", padx=(4, 4))

        self._btn_etapa1 = tk.Button(
            bar, text="1 · Contabilidade × Relatórios",
            bg=COR_ACENTO, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=5, cursor="hand2",
            command=lambda: self._trocar_etapa(1))
        self._btn_etapa1.pack(side="left", padx=2, pady=10)

        self._btn_etapa2 = tk.Button(
            bar, text="2 · Recibos × Pagamentos",
            bg=COR_CINZA, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=5, cursor="hand2",
            command=lambda: self._trocar_etapa(2))
        self._btn_etapa2.pack(side="left", padx=2, pady=10)

        sep2 = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep2.pack(side="left", fill="y", pady=10, padx=8)

        self._btn_processar_e1 = tk.Button(
            bar, text="▶  Processar Etapa 1",
            bg=COR_VERDE, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=5, cursor="hand2",
            command=self._processar_etapa1)
        self._btn_processar_e1.pack(side="left", padx=2, pady=10)

        self._btn_processar_e2 = tk.Button(
            bar, text="▶  Conciliar Etapa 2",
            bg=COR_AZUL, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=5, cursor="hand2",
            command=self._processar_etapa2)
        self._btn_processar_e2.pack(side="left", padx=2, pady=10)

        sep_tol = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep_tol.pack(side="left", fill="y", pady=10, padx=8)
        tk.Label(bar, text="Margem autom. (R$):", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 2))
        self._tolerancia_var = tk.StringVar(value="0,01")
        self._entry_tolerancia = tk.Entry(bar, width=6, font=("Segoe UI", 10),
                                          textvariable=self._tolerancia_var,
                                          bg=COR_BG, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                          relief="flat", bd=1)
        self._entry_tolerancia.pack(side="left", pady=10)
        # tooltip simples: a margem só afeta o algoritmo de conciliação automática
        self._criar_tooltip(self._entry_tolerancia,
            "Tolerância de valor (até R$ 0,50) usada apenas na\n"
            "conciliação AUTOMÁTICA. Use 'Processar/Conciliar'\n"
            "de novo após alterar. Não afeta a conciliação manual\n"
            "nem os filtros de R$ mín/máx das tabelas.")

        self._btn_exportar = tk.Button(
            bar, text="📥  Exportar Excel",
            bg=COR_LARANJA, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=5, cursor="hand2",
            command=self._exportar)
        self._btn_exportar.pack(side="left", padx=4, pady=10)

        self._btn_tema = botao_tema(bar, callback=self._aplicar_tema)
        self._btn_tema.pack(side="right", padx=12, pady=10)

        # ── Linha 2: filtros (status, datas, valor) ─────────────────────────
        bar2 = tk.Frame(container, bg=COR_PAINEL, height=50)
        bar2.pack(fill="x")
        bar2.pack_propagate(False)

        tk.Label(bar2, text="Filtros  ›  Status:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(16, 2))
        self._filtro_status = ttk.Combobox(bar2, state="readonly", width=13,
            values=["Todos", "conciliado", "divergente", "sem_match", "pendente", "parcial"])
        self._filtro_status.set("Todos")
        self._filtro_status.pack(side="left", pady=10)
        self._filtro_status.bind("<<ComboboxSelected>>", lambda _: self._atualizar_tabelas())

        sep4 = tk.Frame(bar2, bg=COR_BORDA, width=1)
        sep4.pack(side="left", fill="y", pady=10, padx=8)
        tk.Label(bar2, text="De:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 2))
        if _TEM_TKCALENDAR:
            self._filtro_data_ini = DateEntry(bar2, width=10, font=("Segoe UI", 10),
                                              background=COR_ACENTO, foreground="white",
                                              borderwidth=1, date_pattern="dd/mm/yyyy",
                                              locale="pt_BR",
                                              mindate=_dt.date(2015, 1, 1),
                                              maxdate=_dt.date(2035, 12, 31))
            self._filtro_data_ini.pack(side="left", pady=10)
            self._filtro_data_ini.set_date(_dt.date(2015, 1, 1))
        else:
            self._filtro_data_ini = tk.Entry(bar2, width=10, font=("Segoe UI", 10),
                                             bg=COR_BG, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                             relief="flat", bd=1)
            self._filtro_data_ini.pack(side="left", pady=10)
            self._filtro_data_ini.insert(0, "01/01/2015")
        tk.Label(bar2, text="Até:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(6, 2))
        if _TEM_TKCALENDAR:
            self._filtro_data_fim = DateEntry(bar2, width=10, font=("Segoe UI", 10),
                                              background=COR_ACENTO, foreground="white",
                                              borderwidth=1, date_pattern="dd/mm/yyyy",
                                              locale="pt_BR",
                                              mindate=_dt.date(2015, 1, 1),
                                              maxdate=_dt.date(2035, 12, 31))
            self._filtro_data_fim.pack(side="left", pady=10)
            self._filtro_data_fim.set_date(_dt.date(2035, 12, 31))
        else:
            self._filtro_data_fim = tk.Entry(bar2, width=10, font=("Segoe UI", 10),
                                             bg=COR_BG, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                             relief="flat", bd=1)
            self._filtro_data_fim.pack(side="left", pady=10)
            self._filtro_data_fim.insert(0, "31/12/2035")
        btn_filtrar = tk.Button(bar2, text="🔍 Filtrar", bg=COR_ACENTO2, fg="white",
                                font=("Segoe UI", 10, "bold"), relief="flat",
                                padx=8, pady=5, cursor="hand2",
                                command=self._atualizar_tabelas)
        btn_filtrar.pack(side="left", padx=4, pady=10)
        btn_limpar = tk.Button(bar2, text="✕ Datas", bg=COR_CINZA, fg="white",
                               font=("Segoe UI", 10, "bold"), relief="flat",
                               padx=6, pady=5, cursor="hand2",
                               command=self._limpar_filtro_data)
        btn_limpar.pack(side="left", pady=10)

        sep5 = tk.Frame(bar2, bg=COR_BORDA, width=1)
        sep5.pack(side="left", fill="y", pady=10, padx=8)
        tk.Label(bar2, text="R$ mín:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 2))
        self._filtro_val_min = tk.Entry(bar2, width=8, font=("Segoe UI", 10),
                                        bg=COR_BG, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                        relief="flat", bd=1)
        self._filtro_val_min.pack(side="left", pady=10)
        tk.Label(bar2, text="máx:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10)).pack(side="left", padx=(6, 2))
        self._filtro_val_max = tk.Entry(bar2, width=8, font=("Segoe UI", 10),
                                        bg=COR_BG, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                        relief="flat", bd=1)
        self._filtro_val_max.pack(side="left", pady=10)
        btn_limpar_val = tk.Button(bar2, text="✕ Valor", bg=COR_CINZA, fg="white",
                                   font=("Segoe UI", 10, "bold"), relief="flat",
                                   padx=6, pady=5, cursor="hand2",
                                   command=self._limpar_filtro_valor)
        btn_limpar_val.pack(side="left", padx=(2, 0), pady=10)

        btn_limpar_tudo = tk.Button(bar2, text="✕ Todos os filtros", bg=COR_VERMELHO, fg="white",
                                    font=("Segoe UI", 10, "bold"), relief="flat",
                                    padx=8, pady=5, cursor="hand2",
                                    command=self._limpar_todos_filtros)
        btn_limpar_tudo.pack(side="left", padx=(10, 0), pady=10)

    def _limpar_todos_filtros(self):
        self._filtro_status.set("Todos")
        self._limpar_filtro_data()
        self._limpar_filtro_valor()

    def _criar_tooltip(self, widget, texto):
        """Tooltip simples (sem dependências externas)."""
        tip = {"win": None}

        def _show(_e=None):
            if tip["win"] is not None:
                return
            x = widget.winfo_rootx()
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.wm_geometry(f"+{x}+{y}")
            tk.Label(win, text=texto, justify="left", bg="#222222", fg="white",
                     font=("Segoe UI", 9), padx=8, pady=6, relief="solid",
                     borderwidth=1).pack()
            tip["win"] = win

        def _hide(_e=None):
            if tip["win"] is not None:
                tip["win"].destroy()
                tip["win"] = None

        widget.bind("<Enter>", _show)
        widget.bind("<Leave>", _hide)

    def _build_sidebar(self, parent):
        outer = tk.Frame(parent, bg=COR_PAINEL, width=220)
        outer.pack(side="left", fill="y", padx=(0, 10), pady=10)
        outer.pack_propagate(False)

        canvas = tk.Canvas(outer, bg=COR_PAINEL, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)
        frame = tk.Frame(canvas, bg=COR_PAINEL)
        _win = canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_win, width=e.width))
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._lbl_arq = {}

        def _sec(txt):
            tk.Label(frame, text=txt, bg=COR_PAINEL, fg=COR_ACENTO,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 2))

        def _arq_btn(chave, label, cmd, cor):
            tk.Button(frame, text=label, bg=cor, fg="white",
                      font=("Segoe UI", 10, "bold"), relief="flat",
                      padx=8, pady=4, cursor="hand2", anchor="w",
                      command=cmd).pack(fill="x", padx=10, pady=(4, 0))
            lbl = tk.Label(frame, text="(não carregado)", bg=COR_PAINEL,
                           fg=COR_TEXTO_SEC, font=("Segoe UI", 9),
                           wraplength=196, justify="left")
            lbl.pack(anchor="w", padx=12, pady=(0, 2))
            self._lbl_arq[chave] = lbl

        _sec("ETAPA 1 — ARQUIVOS")
        _arq_btn("razao",  "📒 Razão Contábil",    self._carregar_razao,  COR_ACENTO2)
        _arq_btn("cupom",  "📄 Cupom Fiscal (PDF)", self._carregar_cupom,  COR_ACENTO2)
        _arq_btn("nf",     "🧾 Nota Fiscal (PDF)",  self._carregar_nf,     COR_ACENTO2)

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=10, pady=8)
        _sec("RESUMO ETAPA 1")
        self._res_e1 = {}
        for k, txt in [("e1_cupom_total", "Impressoras/dia:"),
                        ("e1_cupom_conc",  "  ✅ Conciliadas:"),
                        ("e1_cupom_div",   "  ⚠ Divergentes:"),
                        ("e1_nf_total",    "Notas Fiscais:"),
                        ("e1_nf_conc",     "  ✅ Conciliadas:"),
                        ("e1_nf_div",      "  ⚠ Divergentes:")]:
            r = tk.Frame(frame, bg=COR_PAINEL)
            r.pack(fill="x", padx=10, pady=1)
            tk.Label(r, text=txt, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 9)).pack(side="left")
            l = tk.Label(r, text="—", bg=COR_PAINEL, fg=COR_TEXTO,
                         font=("Segoe UI", 9, "bold"))
            l.pack(side="right")
            self._res_e1[k] = l

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=10, pady=8)
        _sec("ETAPA 2 — ARQUIVOS")
        _arq_btn("recibos", "📋 Recibos (PDF)",     self._carregar_recibos, COR_AZUL)
        _arq_btn("getnet",  "🏦 Pagamentos Getnet", self._carregar_getnet,  COR_AZUL)

        tk.Frame(frame, bg=COR_BORDA, height=1).pack(fill="x", padx=10, pady=8)
        _sec("RESUMO ETAPA 2")
        self._res_e2 = {}
        for k, txt in [("e2_rec_total", "Recibos:"),
                        ("e2_rec_conc",  "  ✅ Conciliados:"),
                        ("e2_rec_pend",  "  ❌ Pendentes:"),
                        ("e2_pag_total", "Pagamentos:"),
                        ("e2_pag_conc",  "  ✅ Conciliados:"),
                        ("e2_pag_pend",  "  ❌ Pendentes:"),
                        ("e2_check_dias",  "Check dias:"),
                        ("e2_check_conc",  "  ✅ Ok:"),
                        ("e2_check_div",   "  ⚠ Divergentes:")]:
            r = tk.Frame(frame, bg=COR_PAINEL)
            r.pack(fill="x", padx=10, pady=1)
            tk.Label(r, text=txt, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 9)).pack(side="left")
            l = tk.Label(r, text="—", bg=COR_PAINEL, fg=COR_TEXTO,
                         font=("Segoe UI", 9, "bold"))
            l.pack(side="right")
            self._res_e2[k] = l

    def _build_area_principal(self, parent):
        frame = tk.Frame(parent, bg=COR_BG)
        frame.pack(side="left", fill="both", expand=True, pady=10)

        self._nb = ttk.Notebook(frame)
        self._nb.pack(fill="both", expand=True)

        self._rodape_labels = {}  # chave → label

        # Aba Cupons
        self._frm_cupom = tk.Frame(self._nb, bg=COR_BG)
        self._nb.add(self._frm_cupom, text="  Cupons Fiscais  ")
        self._cols_cupom = ["impr","doc_nro","data","qtd_cupons",
                             "val_relatorio","val_contabil","diferenca","sequencias","status"]
        self.tree_cupom = self._build_tree(
            self._frm_cupom, self._cols_cupom,
            {"impr":60,"doc_nro":70,"data":100,"qtd_cupons":80,
             "val_relatorio":120,"val_contabil":120,"diferenca":110,
             "sequencias":160,"status":100})
        self._build_rodape(self._frm_cupom, "cupom",
            [("Relatório:", "val_relatorio"), ("Contábil:", "val_contabil"), ("Diferença:", "diferenca")])

        # Aba Notas
        self._frm_nf = tk.Frame(self._nb, bg=COR_BG)
        self._nb.add(self._frm_nf, text="  Notas Fiscais  ")
        self._cols_nf = ["nota","doc_nro","cliente","cond","data",
                          "val_relatorio","val_contabil","diferenca","sequencias","status"]
        self.tree_nf = self._build_tree(
            self._frm_nf, self._cols_nf,
            {"nota":80,"doc_nro":80,"cliente":220,"cond":80,"data":100,
             "val_relatorio":120,"val_contabil":120,"diferenca":110,
             "sequencias":160,"status":100})
        self._build_rodape(self._frm_nf, "nf",
            [("Relatório:", "val_relatorio"), ("Contábil:", "val_contabil"), ("Diferença:", "diferenca")])

        # Aba Recibos
        self._frm_rec = tk.Frame(self._nb, bg=COR_BG)
        self._nb.add(self._frm_rec, text="  Recibos  ")
        self._cols_rec = ["numero","data","valor","saldo_rest","par","status"]
        self.tree_rec = self._build_tree(
            self._frm_rec, self._cols_rec,
            {"numero":100,"data":100,"valor":110,"saldo_rest":110,"par":90,"status":100})
        self._add_menu_manual(self.tree_rec, "recibo")
        self._build_rodape(self._frm_rec, "rec",
            [("Total:", "valor"), ("Saldo rest.:", "saldo_rest")])
        self.tree_rec.bind("<<TreeviewSelect>>",
                           lambda _: self._atualizar_rodape_selecao("rec", "valor"))

        # Aba Pagamentos
        self._frm_pag = tk.Frame(self._nb, bg=COR_BG)
        self._nb.add(self._frm_pag, text="  Pagamentos Getnet  ")
        self._cols_pag = ["DT_RECEB","HR_RECEB","TERMINAL","CV","TXID",
                          "VALOR","saldo_rest","par_recibo","status"]
        self.tree_pag = self._build_tree(
            self._frm_pag, self._cols_pag,
            {"DT_RECEB":90,"HR_RECEB":60,"TERMINAL":90,"CV":90,
             "TXID":260,"VALOR":90,"saldo_rest":90,"par_recibo":90,"status":90})
        self._add_menu_manual(self.tree_pag, "pagamento")
        self._build_rodape(self._frm_pag, "pag",
            [("Total:", "VALOR"), ("Saldo rest.:", "saldo_rest")])
        self.tree_pag.bind("<<TreeviewSelect>>",
                           lambda _: self._atualizar_rodape_selecao("pag", "VALOR"))

        # Aba Check Diário
        self._frm_check = tk.Frame(self._nb, bg=COR_BG)
        self._nb.add(self._frm_check, text="  ✦ Check Diário  ")
        self._cols_check = ["data","val_recibos","val_contabil",
                            "val_contabil_recibos","val_getnet_rest","diferenca","status"]
        self.tree_check = self._build_tree(
            self._frm_check, self._cols_check,
            {"data":110,"val_recibos":140,"val_contabil":140,
             "val_contabil_recibos":180,"val_getnet_rest":160,
             "diferenca":130,"status":110})
        self._build_rodape(self._frm_check, "check",
            [("Recibos:", "val_recibos"), ("Contábil:", "val_contabil"),
             ("Getnet:", "val_getnet_rest"), ("Diferença:", "diferenca")])

    def _build_tree(self, parent, cols, larguras):
        frm = tk.Frame(parent, bg=COR_BG)
        frm.pack(fill="both", expand=True, padx=4, pady=4)
        tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="extended")
        num_cols = {"valor","val_relatorio","val_contabil","diferenca",
                    "VALOR","saldo_rest","qtd_cupons"}
        for col in cols:
            tree.heading(col, text=col.upper().replace("_", " "))
            tree.column(col, width=larguras.get(col, 100),
                        anchor="center" if col in num_cols else "w")
        sby = ttk.Scrollbar(frm, orient="vertical",   command=tree.yview)
        sbx = ttk.Scrollbar(frm, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sby.set, xscrollcommand=sbx.set)
        sby.pack(side="right",  fill="y")
        sbx.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        aplicar_tags_tree(tree)
        return tree

    def _build_rodape(self, parent, chave, campos):
        """
        Cria barra de rodapé com somas dos campos visíveis e (se houver) selecionados.
        campos = lista de (label_txt, col_name)
        """
        bar = tk.Frame(parent, bg=COR_PAINEL, height=28)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        tk.Label(bar, text="∑ visíveis:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 4))

        for lbl_txt, col in campos:
            tk.Label(bar, text=lbl_txt, bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 10)).pack(side="left", padx=(6, 2))
            lbl = tk.Label(bar, text="—", bg=COR_PAINEL, fg=COR_TEXTO,
                           font=("Segoe UI", 10, "bold"))
            lbl.pack(side="left", padx=(0, 4))
            self._rodape_labels[f"{chave}_vis_{col}"] = lbl

        sep = tk.Frame(bar, bg=COR_BORDA, width=1)
        sep.pack(side="left", fill="y", pady=4, padx=6)

        tk.Label(bar, text="∑ selecionados:", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=(4, 4))

        # Apenas primeira coluna numérica para seleção
        col_sel = campos[0][1]
        lbl_n = tk.Label(bar, text="0 itens", bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                         font=("Segoe UI", 10))
        lbl_n.pack(side="left", padx=(0, 4))
        lbl_s = tk.Label(bar, text="R$ —", bg=COR_PAINEL, fg=COR_ACENTO,
                         font=("Segoe UI", 10, "bold"))
        lbl_s.pack(side="left")
        self._rodape_labels[f"{chave}_sel_n"]   = lbl_n
        self._rodape_labels[f"{chave}_sel_sum"] = lbl_s
        self._rodape_labels[f"{chave}_sel_col"] = col_sel  # string, não label

    def _atualizar_rodapes(self, df_cupom, df_nf, df_rec, df_pag, df_check):
        """Atualiza os totais visíveis em todos os rodapés."""
        def _soma(df, col):
            if df is None or df.empty or col not in df.columns:
                return None
            return pd.to_numeric(df[col], errors="coerce").sum()

        def _fmt(v):
            if v is None or (isinstance(v, float) and v != v):
                return "—"
            return f"R$ {v:,.2f}"

        specs = [
            ("cupom", df_cupom,  [("val_relatorio","val_relatorio"),("val_contabil","val_contabil"),("diferenca","diferenca")]),
            ("nf",    df_nf,     [("val_relatorio","val_relatorio"),("val_contabil","val_contabil"),("diferenca","diferenca")]),
            ("rec",   df_rec,    [("valor","valor"),("saldo_rest","saldo_rest")]),
            ("pag",   df_pag,    [("VALOR","VALOR"),("saldo_rest","saldo_rest")]),
            ("check", df_check,  [("val_recibos","val_recibos"),("val_contabil","val_contabil"),
                                   ("val_getnet_rest","val_getnet_rest"),("diferenca","diferenca")]),
        ]
        for chave, df, cols in specs:
            for _, col in cols:
                key = f"{chave}_vis_{col}"
                if key in self._rodape_labels:
                    v = _soma(df, col)
                    self._rodape_labels[key].config(text=_fmt(v))
            # Resetar seleção
            if f"{chave}_sel_n" in self._rodape_labels:
                self._rodape_labels[f"{chave}_sel_n"].config(text="0 itens")
                self._rodape_labels[f"{chave}_sel_sum"].config(text="R$ —")

    def _atualizar_rodape_selecao(self, chave, col_val):
        """Atualiza soma dos itens selecionados na aba indicada."""
        tree_map = {"rec": self.tree_rec, "pag": self.tree_pag}
        df_map   = {"rec": self.res_recibos, "pag": self.res_getnet}
        tree = tree_map.get(chave)
        df   = df_map.get(chave)
        if tree is None or df is None:
            return
        sel = tree.selection()
        iid_map = getattr(tree, "_iid_map", {})
        total = 0.0
        for iid in sel:
            idx = iid_map.get(iid)
            if idx is not None and idx in df.index:
                v = df.at[idx, col_val]
                try:
                    total += float(v)
                except (TypeError, ValueError):
                    pass
        lbl_n   = self._rodape_labels.get(f"{chave}_sel_n")
        lbl_sum = self._rodape_labels.get(f"{chave}_sel_sum")
        if lbl_n:
            lbl_n.config(text=f"{len(sel)} {'item' if len(sel)==1 else 'itens'}")
        if lbl_sum:
            lbl_sum.config(text=f"R$ {total:,.2f}" if sel else "R$ —")

    def _add_menu_manual(self, tree, tipo):
        """Adiciona menu de contexto (botão direito) para conciliação manual."""
        menu = tk.Menu(self, tearoff=0, bg=COR_PAINEL, fg=COR_TEXTO,
                       activebackground=COR_ACENTO, activeforeground="white",
                       font=("Segoe UI", 11))
        if tipo == "recibo":
            menu.add_command(label="🔗  Conciliar com pagamento(s) selecionado(s)",
                             command=self._conciliar_manual)
            menu.add_command(label="🔗  Conciliar com pagamento por valor/data",
                             command=lambda: self._conciliar_manual(auto_pag=True))
        else:
            menu.add_command(label="🔗  Conciliar com recibo(s) selecionado(s)",
                             command=self._conciliar_manual)
            menu.add_command(label="🔗  Conciliar com recibo por valor/data",
                             command=lambda: self._conciliar_manual(auto_rec=True))
        menu.add_separator()
        menu.add_command(label="✂  Desfazer conciliação deste item",
                         command=lambda: self._desfazer_conciliacao(tipo))
        menu.add_command(label="✂  Desfazer par inteiro",
                         command=lambda: self._desfazer_par(tipo))

        def _popup(event):
            row = tree.identify_row(event.y)
            # Só substitui a seleção se o clique foi fora da seleção atual.
            # Se o item clicado já faz parte de uma seleção múltipla (Ctrl/Shift),
            # preserva todos os itens selecionados.
            if row and row not in tree.selection():
                tree.selection_set(row)
            tree.focus(row)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        tree.bind("<Button-3>", _popup)

    def _get_sel_idx(self, tree):
        """Retorna o índice do DataFrame para a linha selecionada na tree (primeira seleção)."""
        sel = tree.selection()
        if not sel:
            return None
        iid = sel[0]
        iid_map = getattr(tree, "_iid_map", {})
        return iid_map.get(iid)

    def _get_sel_idxs(self, tree):
        """Retorna lista de índices do DataFrame para todas as linhas selecionadas."""
        sel = tree.selection()
        iid_map = getattr(tree, "_iid_map", {})
        return [iid_map[iid] for iid in sel if iid in iid_map]

    def _conciliar_manual(self, auto_pag=False, auto_rec=False):
        """Concilia manualmente recibo(s) ↔ pagamento(s) selecionados (N×N ou auto-buscado)."""
        if self.res_recibos is None or self.res_getnet is None:
            messagebox.showwarning("Aviso", "Execute a Etapa 2 primeiro.")
            return

        dr = self.res_recibos
        dp = self.res_getnet

        idxs_r = self._get_sel_idxs(self.tree_rec)
        idxs_p = self._get_sel_idxs(self.tree_pag)

        if auto_pag and idxs_r:
            # Para cada recibo selecionado, busca pagamento pendente de mesmo valor e data
            idx_r = idxs_r[0]
            row_r = dr.loc[idx_r]
            cands = dp[
                (dp["status"] == "pendente") &
                (abs(dp["VALOR"] - row_r["valor"]) <= 0.01) &
                (dp["DT_RECEB"].apply(_normalizar_data) ==
                 _normalizar_data(row_r["data"]))
            ]
            if cands.empty:
                messagebox.showinfo("Sem match",
                    "Nenhum pagamento pendente com mesmo valor e data encontrado.")
                return
            idxs_p = [cands.index[0]]

        elif auto_rec and idxs_p:
            idx_p = idxs_p[0]
            row_p = dp.loc[idx_p]
            cands = dr[
                (dr["status"] == "pendente") &
                (abs(dr["valor"] - row_p["VALOR"]) <= 0.01) &
                (dr["data"].apply(_normalizar_data) ==
                 _normalizar_data(row_p["DT_RECEB"]))
            ]
            if cands.empty:
                messagebox.showinfo("Sem match",
                    "Nenhum recibo pendente com mesmo valor e data encontrado.")
                return
            idxs_r = [cands.index[0]]

        if not idxs_r or not idxs_p:
            messagebox.showwarning("Seleção necessária",
                "Selecione ao menos um recibo E um pagamento (ou use a opção automática).")
            return

        # Gera próximo número de par
        par_nums = [int(p[3:].rstrip("M")) for p in
                    list(dr["par"].str.extractall(r"E2P(\d{4})")[0])
                    + list(dp["par_recibo"].str.extractall(r"E2P(\d{4})")[0])
                    if p]
        next_n = (max(par_nums) + 1) if par_nums else 1
        par = f"E2P{next_n:04d}M"

        # Soma total dos recibos e pagamentos selecionados
        total_r = sum(dr.at[i, "saldo_rest"] for i in idxs_r)
        total_p = sum(dp.at[i, "saldo_rest"] for i in idxs_p)
        valor_match = min(total_r, total_p)

        if valor_match <= 0.01:
            messagebox.showinfo("Aviso", "Itens selecionados não possuem saldo disponível.")
            return

        # Distribui o match proporcionalmente entre os recibos
        resto_r = valor_match
        for idx_r in idxs_r:
            saldo = dr.at[idx_r, "saldo_rest"]
            if saldo <= 0.01 or resto_r <= 0.01:
                continue
            consumido = min(saldo, resto_r)
            par_ant = str(dr.at[idx_r, "par"])
            dr.at[idx_r, "par"] = (f"{par_ant},{par}" if par_ant and par_ant != "nan" else par)
            dr.at[idx_r, "saldo_rest"] = round(saldo - consumido, 2)
            dr.at[idx_r, "status"] = ("conciliado" if dr.at[idx_r, "saldo_rest"] <= 0.01
                                       else "parcial")
            resto_r = round(resto_r - consumido, 2)

        # Distribui o match proporcionalmente entre os pagamentos
        resto_p = valor_match
        for idx_p in idxs_p:
            saldo = dp.at[idx_p, "saldo_rest"]
            if saldo <= 0.01 or resto_p <= 0.01:
                continue
            consumido = min(saldo, resto_p)
            par_ant = str(dp.at[idx_p, "par_recibo"])
            dp.at[idx_p, "par_recibo"] = (f"{par_ant},{par}" if par_ant and par_ant != "nan" else par)
            dp.at[idx_p, "saldo_rest"] = round(saldo - consumido, 2)
            dp.at[idx_p, "status"] = ("conciliado" if dp.at[idx_p, "saldo_rest"] <= 0.01
                                       else "parcial")
            resto_p = round(resto_p - consumido, 2)

        self.res_recibos = dr
        self.res_getnet  = dp
        self._recalcular_check()
        self._atualizar_tabelas()
        self._atualizar_resumo_e2()
        n_r = len(idxs_r)
        n_p = len(idxs_p)
        self.status_var.set(
            f"✅ Par manual {par} — {n_r} recibo(s) ↔ {n_p} pagamento(s) "
            f"(R$ {valor_match:,.2f})")

    def _desfazer_conciliacao(self, tipo):
        """Desfaz a conciliação do item selecionado (remove seus pares)."""
        if tipo == "recibo":
            idx = self._get_sel_idx(self.tree_rec)
            if idx is None:
                return
            self._desfazer_por_pares(self.res_recibos.at[idx, "par"])
        else:
            idx = self._get_sel_idx(self.tree_pag)
            if idx is None:
                return
            self._desfazer_por_pares(self.res_getnet.at[idx, "par_recibo"])

    def _desfazer_par(self, tipo):
        """Desfaz um par inteiro (recibo + pagamento)."""
        if tipo == "recibo":
            idx = self._get_sel_idx(self.tree_rec)
            if idx is None:
                return
            pares_str = str(self.res_recibos.at[idx, "par"])
        else:
            idx = self._get_sel_idx(self.tree_pag)
            if idx is None:
                return
            pares_str = str(self.res_getnet.at[idx, "par_recibo"])
        self._desfazer_por_pares(pares_str)

    def _desfazer_por_pares(self, pares_str):
        """Remove todos os pares listados em pares_str de recibos e pagamentos."""
        if not pares_str or pares_str in ("", "nan"):
            messagebox.showinfo("Aviso", "Este item não possui conciliação para desfazer.")
            return
        pares = [p.strip() for p in str(pares_str).split(",") if p.strip()]
        if not pares:
            return
        dr = self.res_recibos
        dp = self.res_getnet

        for par in pares:
            # Reset recibos
            for idx_r in dr[dr["par"].str.contains(par, na=False)].index:
                dr.at[idx_r, "par"] = ""
                dr.at[idx_r, "saldo_rest"] = dr.at[idx_r, "valor"]
                dr.at[idx_r, "status"] = "pendente"
            # Reset pagamentos
            for idx_p in dp[dp["par_recibo"].str.contains(par, na=False)].index:
                dp.at[idx_p, "par_recibo"] = ""
                dp.at[idx_p, "saldo_rest"] = dp.at[idx_p, "VALOR"]
                dp.at[idx_p, "status"] = "pendente"

        self.res_recibos = dr
        self.res_getnet  = dp
        self._recalcular_check()
        self._atualizar_tabelas()
        self._atualizar_resumo_e2()
        self.status_var.set(f"↩ Par(es) {', '.join(pares)} desfeito(s).")

    def _recalcular_check(self):
        """Recalcula o Check Diário com base no estado atual de recibos e pagamentos."""
        if self.res_recibos is None or self.res_getnet is None:
            return
        dr = self.res_recibos.copy()
        dp = self.res_getnet.copy()
        dr["_data_norm"] = dr["data"].apply(_normalizar_data)
        dp["_data_norm"] = dp["DT_RECEB"].apply(_normalizar_data)

        rec_conc = dr[dr["status"].isin(["conciliado", "parcial"])].copy()
        rec_conc["_valor_conciliado"] = rec_conc["valor"] - rec_conc["saldo_rest"]
        rec_por_dia = (rec_conc.groupby("_data_norm")["_valor_conciliado"]
                       .sum().round(2).rename("val_recibos"))
        get_rest_por_dia = (dp.groupby("_data_norm")["VALOR"]
                            .sum().round(2).rename("val_getnet_rest"))

        contabil_por_dia = {}
        for df_e1, col_data in [(self.res_cupom, "data"), (self.res_nf, "data")]:
            if df_e1 is None or df_e1.empty:
                continue
            tmp = df_e1.copy()
            tmp["_data_norm"] = tmp[col_data].apply(_normalizar_data)
            grp = tmp.groupby("_data_norm")["val_contabil"].sum()
            for dn, val in grp.items():
                contabil_por_dia[dn] = round(contabil_por_dia.get(dn, 0.0) + val, 2)

        todas_datas = (set(rec_por_dia.index) | set(get_rest_por_dia.index)
                       | set(contabil_por_dia.keys()))
        COLS_CHECK = ["data", "val_recibos", "val_contabil", "val_getnet_rest",
                      "val_contabil_recibos", "diferenca", "status"]
        check_rows = []
        for dn in sorted(todas_datas):
            val_rec   = round(rec_por_dia.get(dn, 0.0), 2)
            val_cont  = round(contabil_por_dia.get(dn, 0.0), 2)
            val_grest = round(get_rest_por_dia.get(dn, 0.0), 2)
            val_cont_rec = round(val_cont + val_rec, 2)
            dif = round(val_cont_rec - val_grest, 2)
            status = "conciliado" if abs(dif) <= 0.01 else "divergente"
            check_rows.append({
                "data":                 dn.strftime("%d/%m/%Y") if dn else "",
                "val_recibos":          val_rec,
                "val_contabil":         val_cont,
                "val_getnet_rest":      val_grest,
                "val_contabil_recibos": val_cont_rec,
                "diferenca":            dif,
                "status":               status,
            })
        self.res_check = (pd.DataFrame(check_rows, columns=COLS_CHECK)
                          if check_rows else pd.DataFrame(columns=COLS_CHECK))

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=COR_PAINEL, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.status_var = tk.StringVar(value="Pronto. Carregue os arquivos para começar.")
        tk.Label(bar, textvariable=self.status_var,
                 bg=COR_PAINEL, fg=COR_TEXTO_SEC,
                 font=("Segoe UI", 10), anchor="w").pack(side="left", padx=10)

    def _aplicar_estilos(self):
        aplicar_estilos_ttk(ttk.Style(self))
        registrar_callback(self._aplicar_tema)

    def _aplicar_tema(self):
        _cores()
        self.configure(bg=T("BG"))
        aplicar_estilos_ttk(ttk.Style(self))
        for tree in (self.tree_cupom, self.tree_nf, self.tree_rec, self.tree_pag):
            aplicar_tags_tree(tree)
        from theme import recolorir_widget
        recolorir_widget(self)
        self._atualizar_tabelas()

    def _trocar_etapa(self, etapa):
        if etapa == 1:
            self._btn_etapa1.config(bg=COR_ACENTO)
            self._btn_etapa2.config(bg=COR_CINZA)
            self._nb.select(0)
        else:
            self._btn_etapa1.config(bg=COR_CINZA)
            self._btn_etapa2.config(bg=COR_AZUL)
            self._nb.select(2)

    # ─── Carregamento ─────────────────────────────────────────────────────────

    def _set_lbl(self, chave, path, n, extra=""):
        self._lbl_arq[chave].config(
            text=f"✅ {os.path.basename(path)} ({n} reg.){extra}")

    def _carregar_razao(self):
        path = filedialog.askopenfilename(
            title="Razão Contábil",
            filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.status_var.set("⏳ Lendo Razão Contábil...")
            self.update()
            df = ler_razao(path)
            self._df_razao = df
            self._set_lbl("razao", path, len(df),
                          f" | R$ {df['valor'].sum():,.2f}")
            self.status_var.set(
                f"✅ Razão — {len(df)} lançamentos (2074) | "
                f"Total: R$ {df['valor'].sum():,.2f}")
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Razão:\n{e}\n\n{traceback.format_exc()}")

    def _carregar_cupom(self):
        path = filedialog.askopenfilename(
            title="Relatório Cupom Fiscal",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.status_var.set("⏳ Lendo Cupom Fiscal...")
            self.update()
            df = ler_cupom_fiscal(path)
            self._df_cupom = df
            self._set_lbl("cupom", path, len(df),
                          f" | R$ {df['valor'].sum():,.2f}")
            self.status_var.set(
                f"✅ Cupom — {len(df)} registros PIX Maquineta | "
                f"Total: R$ {df['valor'].sum():,.2f}")
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Cupom:\n{e}\n\n{traceback.format_exc()}")

    def _carregar_nf(self):
        path = filedialog.askopenfilename(
            title="Relatório Nota Fiscal",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.status_var.set("⏳ Lendo Nota Fiscal...")
            self.update()
            df = ler_nota_fiscal(path)
            self._df_nf = df
            self._set_lbl("nf", path, len(df),
                          f" | R$ {df['valor'].sum():,.2f}")
            self.status_var.set(
                f"✅ NF — {len(df)} notas | "
                f"Total: R$ {df['valor'].sum():,.2f}")
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"NF:\n{e}\n\n{traceback.format_exc()}")

    def _carregar_recibos(self):
        path = filedialog.askopenfilename(
            title="Relatório Recibos",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.status_var.set("⏳ Lendo Recibos...")
            self.update()
            df = ler_recibos(path)
            self._df_recibos = df
            self._set_lbl("recibos", path, len(df),
                          f" | R$ {df['valor'].sum():,.2f}")
            self.status_var.set(
                f"✅ Recibos — {len(df)} registros | "
                f"Total: R$ {df['valor'].sum():,.2f}")
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Recibos:\n{e}\n\n{traceback.format_exc()}")

    def _carregar_getnet(self):
        path = filedialog.askopenfilename(
            title="Pagamentos Getnet (XLSX)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not path:
            return
        try:
            self.status_var.set("⏳ Lendo Pagamentos Getnet...")
            self.update()
            df = ler_mov_pix(path)
            self._df_getnet = df
            self._set_lbl("getnet", path, len(df),
                          f" | R$ {df['VALOR'].sum():,.2f}")
            self.status_var.set(
                f"✅ Getnet — {len(df)} transações 'Paga' | "
                f"Total: R$ {df['VALOR'].sum():,.2f}")
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Getnet:\n{e}\n\n{traceback.format_exc()}")

    def _get_tolerancia(self):
        """Lê a margem de tolerância (R$) digitada pelo usuário.
        Aceita vírgula ou ponto. Limita entre 0,00 e 0,50. Padrão: 0,01."""
        s = self._tolerancia_var.get().strip().replace(",", ".")
        try:
            v = float(s)
        except ValueError:
            v = 0.01
        v = max(0.0, min(v, 0.50))
        return round(v, 2)

    # ─── Processamento ────────────────────────────────────────────────────────

    def _processar_etapa1(self):
        if self._df_razao is None:
            messagebox.showwarning("Aviso", "Carregue a Razão Contábil primeiro.")
            return
        if self._df_cupom is None and self._df_nf is None:
            messagebox.showwarning("Aviso",
                "Carregue ao menos um relatório (Cupom Fiscal ou Nota Fiscal).")
            return
        try:
            self.status_var.set("⏳ Processando Etapa 1...")
            self.update()
            self.res_cupom, self.res_nf = conciliar_etapa1(
                self._df_razao, self._df_cupom, self._df_nf,
                tolerancia=self._get_tolerancia())
            self._atualizar_tabelas()
            self._atualizar_resumo_e1()
            n_c = (self.res_cupom["status"] == "conciliado").sum()
            n_n = (self.res_nf["status"]    == "conciliado").sum()
            self.status_var.set(
                f"✅ Etapa 1 — Cupons: {n_c} conciliados | NF: {n_n} conciliadas")
            self._trocar_etapa(1)
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Etapa 1:\n{e}\n\n{traceback.format_exc()}")

    def _processar_etapa2(self):
        if self._df_recibos is None or self._df_getnet is None:
            messagebox.showwarning("Aviso",
                "Carregue Recibos e Pagamentos Getnet primeiro.")
            return
        if self.res_cupom is None and self.res_nf is None:
            messagebox.showwarning("Aviso",
                "Execute a Etapa 1 antes de conciliar a Etapa 2.\n"
                "Os valores contábeis de cupons e NF são necessários para o check diário.")
            return
        try:
            self.status_var.set("⏳ Conciliando Etapa 2...")
            self.update()
            self.res_recibos, self.res_getnet, self.res_check = conciliar_etapa2(
                self._df_recibos, self._df_getnet,
                df_res_cupom=self.res_cupom,
                df_res_nf=self.res_nf)
            self._atualizar_tabelas()
            self._atualizar_resumo_e2()
            n_r = (self.res_recibos["status"] == "conciliado").sum()
            n_p = (self.res_getnet["status"]  == "conciliado").sum()
            n_ok  = (self.res_check["status"] == "conciliado").sum() if self.res_check is not None else 0
            n_div = (self.res_check["status"] == "divergente").sum() if self.res_check is not None else 0
            self.status_var.set(
                f"✅ Etapa 2 — Recibos: {n_r} | Pagamentos: {n_p} conciliados | "
                f"Check diário: {n_ok} ok, {n_div} divergentes")
            self._trocar_etapa(2)
        except Exception as e:
            import traceback
            messagebox.showerror("Erro", f"Etapa 2:\n{e}\n\n{traceback.format_exc()}")

    # ─── Tabelas ──────────────────────────────────────────────────────────────

    def _limpar_filtro_data(self):
        if _TEM_TKCALENDAR:
            self._filtro_data_ini.set_date(_dt.date(2015, 1, 1))
            self._filtro_data_fim.set_date(_dt.date(2035, 12, 31))
        else:
            self._filtro_data_ini.delete(0, "end")
            self._filtro_data_ini.insert(0, "01/01/2015")
            self._filtro_data_fim.delete(0, "end")
            self._filtro_data_fim.insert(0, "31/12/2035")
        self._atualizar_tabelas()

    def _limpar_filtro_valor(self):
        self._filtro_val_min.delete(0, "end")
        self._filtro_val_max.delete(0, "end")
        self._atualizar_tabelas()

    def _parse_filtro_valores(self):
        """Retorna (val_min, val_max) ou (None, None) se vazio/inválido."""
        def _parse(entry):
            s = entry.get().strip().replace(",", ".")
            try:
                return float(s) if s else None
            except ValueError:
                return None
        return _parse(self._filtro_val_min), _parse(self._filtro_val_max)

    def _parse_filtro_datas(self):
        import datetime as _dt
        ini_str = self._filtro_data_ini.get().strip()
        fim_str = self._filtro_data_fim.get().strip()
        ini = fim = None
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                ini = _dt.datetime.strptime(ini_str, fmt).date()
                break
            except ValueError:
                pass
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                fim = _dt.datetime.strptime(fim_str, fmt).date()
                break
            except ValueError:
                pass
        return ini, fim

    def _atualizar_tabelas(self):
        filtro = self._filtro_status.get()
        data_ini, data_fim = self._parse_filtro_datas()
        val_min, val_max = self._parse_filtro_valores()

        # Mapeamento aba → coluna de valor principal
        _col_valor_map = {
            "cupom":  "val_relatorio",
            "nf":     "val_relatorio",
            "rec":    "valor",
            "pag":    "VALOR",
            "check":  "diferenca",
        }

        def _filtrar(df, col_data=None, col_valor=None):
            if df is None or df.empty:
                return df
            r = df if filtro == "Todos" else df[df["status"] == filtro]
            if col_data and data_ini and data_fim and col_data in r.columns:
                def _in_range(val):
                    d = _normalizar_data(val)
                    if d is None:
                        return True
                    return data_ini <= d <= data_fim
                r = r[r[col_data].apply(_in_range)]
            if col_valor and col_valor in r.columns:
                if val_min is not None:
                    r = r[pd.to_numeric(r[col_valor], errors="coerce").abs() >= val_min]
                if val_max is not None:
                    r = r[pd.to_numeric(r[col_valor], errors="coerce").abs() <= val_max]
            return r

        num_cols = {"val_relatorio","val_contabil","diferenca","valor","VALOR","saldo_rest",
                    "val_recibos","val_contabil_recibos","val_getnet_rest"}

        def _preencher(tree, df, cols, col_data=None):
            tree.delete(*tree.get_children())
            if df is None or df.empty:
                return
            # store mapping iid → df index for manual match
            tree._iid_map = {}
            for orig_idx, (_, row) in enumerate(df.iterrows()):
                vals = []
                for c in cols:
                    v = row.get(c, "")
                    if isinstance(v, float):
                        if v != v:
                            v = "—"
                        elif c in num_cols:
                            v = f"{v:,.2f}"
                    vals.append("" if v is None else v)
                status = str(row.get("status", ""))
                tag = {"conciliado": "conciliado",
                       "divergente": "pendente",
                       "sem_match":  "ignorado",
                       "pendente":   "pendente",
                       "parcial":    "parcial"}.get(status, "")
                iid = tree.insert("", "end", values=vals, tags=(tag,))
                tree._iid_map[iid] = row.name  # original DataFrame index

        df_cupom_f  = _filtrar(self.res_cupom,   "data",     "val_relatorio")
        df_nf_f     = _filtrar(self.res_nf,      "data",     "val_relatorio")
        df_rec_f    = _filtrar(self.res_recibos, "data",     "valor")
        df_pag_f    = _filtrar(self.res_getnet,  "DT_RECEB", "VALOR")
        df_check_f  = _filtrar(self.res_check,   "data",     "diferenca")

        _preencher(self.tree_cupom,  df_cupom_f,  self._cols_cupom,  "data")
        _preencher(self.tree_nf,     df_nf_f,     self._cols_nf,     "data")
        _preencher(self.tree_rec,    df_rec_f,    self._cols_rec,    "data")
        _preencher(self.tree_pag,    df_pag_f,    self._cols_pag,    "DT_RECEB")
        _preencher(self.tree_check,  df_check_f,  self._cols_check,  "data")

        self._atualizar_rodapes(df_cupom_f, df_nf_f, df_rec_f, df_pag_f, df_check_f)

    def _atualizar_resumo_e1(self):
        if self.res_cupom is not None and not self.res_cupom.empty:
            vc = self.res_cupom["status"].value_counts()
            self._res_e1["e1_cupom_total"].config(text=str(len(self.res_cupom)))
            self._res_e1["e1_cupom_conc"].config(
                text=str(vc.get("conciliado", 0)), fg=COR_VERDE)
            self._res_e1["e1_cupom_div"].config(
                text=str(vc.get("divergente", 0) + vc.get("sem_match", 0)),
                fg=COR_VERMELHO)
        if self.res_nf is not None and not self.res_nf.empty:
            vn = self.res_nf["status"].value_counts()
            self._res_e1["e1_nf_total"].config(text=str(len(self.res_nf)))
            self._res_e1["e1_nf_conc"].config(
                text=str(vn.get("conciliado", 0)), fg=COR_VERDE)
            self._res_e1["e1_nf_div"].config(
                text=str(vn.get("divergente", 0) + vn.get("sem_match", 0)),
                fg=COR_VERMELHO)

    def _atualizar_resumo_e2(self):
        if self.res_recibos is not None and not self.res_recibos.empty:
            vr = self.res_recibos["status"].value_counts()
            self._res_e2["e2_rec_total"].config(text=str(len(self.res_recibos)))
            self._res_e2["e2_rec_conc"].config(
                text=str(vr.get("conciliado", 0)), fg=COR_VERDE)
            self._res_e2["e2_rec_pend"].config(
                text=str(vr.get("pendente", 0)), fg=COR_VERMELHO)
        if self.res_getnet is not None and not self.res_getnet.empty:
            vp = self.res_getnet["status"].value_counts()
            self._res_e2["e2_pag_total"].config(text=str(len(self.res_getnet)))
            self._res_e2["e2_pag_conc"].config(
                text=str(vp.get("conciliado", 0)), fg=COR_VERDE)
            self._res_e2["e2_pag_pend"].config(
                text=str(vp.get("pendente", 0)), fg=COR_VERMELHO)
        if self.res_check is not None and not self.res_check.empty:
            vc = self.res_check["status"].value_counts()
            self._res_e2["e2_check_dias"].config(text=str(len(self.res_check)))
            self._res_e2["e2_check_conc"].config(
                text=str(vc.get("conciliado", 0)), fg=COR_VERDE)
            self._res_e2["e2_check_div"].config(
                text=str(vc.get("divergente", 0)), fg=COR_VERMELHO)

    # ─── Exportação ───────────────────────────────────────────────────────────

    def _exportar(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        path_out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            title="Salvar relatório de conciliação")
        if not path_out:
            return

        try:
            wb = Workbook()

            def _hdr(ws, row, col, txt, cor="1F4E79"):
                c = ws.cell(row, col, txt)
                c.font      = Font(bold=True, color="FFFFFF", size=10)
                c.fill      = PatternFill("solid", fgColor=cor)
                c.alignment = Alignment(horizontal="center")

            def _escrever_df(ws, df, cor="1F4E79"):
                if df is None or df.empty:
                    ws.cell(1, 1, "(sem dados)")
                    return
                cols = list(df.columns)
                for ci, col in enumerate(cols, 1):
                    _hdr(ws, 1, ci, col.upper().replace("_", " "), cor)
                for ri, (_, row) in enumerate(df.iterrows(), 2):
                    for ci, col in enumerate(cols, 1):
                        v = row.get(col, "")
                        ws.cell(ri, ci, "" if v != v else v)
                for ci in range(1, len(cols) + 1):
                    ws.column_dimensions[get_column_letter(ci)].width = 18

            ws_res = wb.active
            ws_res.title = "Resumo"
            ws_res.cell(1, 1, "CONCILIAÇÃO CONTÁBIL — PMZ").font = Font(bold=True, size=13)
            r = 3

            for titulo, df, cor in [
                ("ETAPA 1 — CUPONS",        self.res_cupom,   "1F4E79"),
                ("ETAPA 1 — NOTAS",         self.res_nf,      "1A5276"),
                ("ETAPA 2 — RECIBOS",       self.res_recibos, "0B5394"),
                ("ETAPA 2 — GETNET",        self.res_getnet,  "7B3F00"),
                ("ETAPA 2 — CHECK DIÁRIO",  self.res_check,   "145A32"),
            ]:
                if df is not None and not df.empty:
                    vc = df["status"].value_counts()
                    ws_res.cell(r, 1, titulo).font = Font(bold=True, size=11)
                    r += 1
                    for k, v in [("Total", len(df)),
                                  ("Conciliados", vc.get("conciliado", 0)),
                                  ("Divergentes/Pendentes",
                                   vc.get("divergente", 0) +
                                   vc.get("sem_match", 0) +
                                   vc.get("pendente", 0))]:
                        ws_res.cell(r, 1, k)
                        ws_res.cell(r, 2, v)
                        r += 1
                    r += 1

            ws_res.column_dimensions["A"].width = 30
            ws_res.column_dimensions["B"].width = 18

            if self.res_cupom is not None and not self.res_cupom.empty:
                _escrever_df(wb.create_sheet("Cupons"), self.res_cupom, "1F4E79")
            if self.res_nf is not None and not self.res_nf.empty:
                _escrever_df(wb.create_sheet("Notas Fiscais"), self.res_nf, "1A5276")
            if self.res_recibos is not None and not self.res_recibos.empty:
                _escrever_df(wb.create_sheet("Recibos"), self.res_recibos, "0B5394")
            if self.res_getnet is not None and not self.res_getnet.empty:
                _escrever_df(wb.create_sheet("Pagamentos Getnet"), self.res_getnet, "7B3F00")
            if self.res_check is not None and not self.res_check.empty:
                _escrever_df(wb.create_sheet("Check Diario"), self.res_check, "145A32")

            wb.save(path_out)
            self.status_var.set(f"✅ Exportado: {os.path.basename(path_out)}")
            messagebox.showinfo("Exportação concluída",
                f"Relatório salvo!\n\n{path_out}")

        except Exception as e:
            import traceback
            messagebox.showerror("Erro na exportação",
                f"{e}\n\n{traceback.format_exc()}")