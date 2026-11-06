"""
Análise de Conta Contábil - Débito x Crédito
----------------------------------------------
Importa uma base exportada (colunas: DEBITO, CREDITO, VALOR EVENTO, DESCRICAO,
DATA, CODCOLIGADA, CNPJ, COD_EVENTO) e calcula:

  1. O saldo (Débito - Crédito) de UMA conta específica informada pelo usuário.
  2. Uma visão consolidada de TODAS as contas presentes na base (débito, crédito
     e saldo por conta), permitindo enxergar o comportamento de cada conta.
  3. Um de-para entre a base (por COD_EVENTO, considerando TODAS as contas) e
     uma Folha Analítica/Sintética em PDF (por Cod), apontando divergências
     de valor e eventos presentes em apenas um dos dois lados.
  4. Uma comparação 3 vias (Base x PDF x Razão) restrita à conta informada:
     casa cada evento (COD_EVENTO/Cod) com o lançamento correspondente no
     Razão (Modelo I, .xlsx ou .ods), casado por valor + lado (débito ou
     crédito), trazendo o histórico e a sequência do razão lado a lado.

Requisitos: pandas, openpyxl, pdfplumber, odfpy
(pip install pandas openpyxl pdfplumber odfpy)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import pdfplumber


# ----------------------------------------------------------------------------
# Tema (paleta escura "slate", seguindo o padrão já usado nas outras telas)
# ----------------------------------------------------------------------------
from core import theme as _t   # identidade visual do Prisma (core/theme.py)


class Theme:
    BG = _t.BG_APP
    BG_PANEL = _t.CARD_BG
    BG_CARD = _t.SURFACE2
    FG = _t.FG_TEXT
    FG_MUTED = _t.FG_MUTED
    ACCENT = _t.ACCENT_BLUE
    GREEN = _t.ACCENT_GREEN
    RED = _t.ACCENT_RED
    BORDER = _t.GRID_LINE
    SELECT = _t.SEL_BG
    WARN = _t.ACCENT_ORANGE
    CYAN = _t.ACCENT_CYAN
    FONT = ("Segoe UI", 10)
    FONT_BOLD = ("Segoe UI", 10, "bold")
    FONT_TITLE = ("Segoe UI", 14, "bold")
    FONT_BIG = ("Segoe UI", 20, "bold")


COL_DEBITO = "DEBITO"
COL_CREDITO = "CREDITO"
COL_VALOR = "VALOR EVENTO"
COL_DESC = "DESCRICAO"
COL_DATA = "DATA"
COL_CODCOLIGADA = "CODCOLIGADA"


# ----------------------------------------------------------------------------
# Exportação para o layout TXT (formato "FOLHA_jun26.txt")
#
# Layout de cada linha (separado por TAB, terminador de linha CRLF, sem
# cabeçalho):
#   1. CODCOLIGADA  - fixo "101" (ou coluna CODCOLIGADA da base, se existir)
#   2. Flag         - fixo "X"
#   3. DEBITO       - conta de débito
#   4. CREDITO      - conta de crédito
#   5. VALOR        - 10 dígitos, zero-padded, sem separador decimal
#                      (os 2 últimos dígitos são os centavos)
#   6. DESCRICAO    - preenchida/cortada para exatamente 100 caracteres
#   7. DATA         - formato DDMMAAAA
# ----------------------------------------------------------------------------
def gerar_linhas_txt_folha(df):
    """
    Recebe o DataFrame já importado (mesmas colunas usadas no restante do
    programa: DEBITO, CREDITO, VALOR EVENTO, DESCRICAO, DATA[, CODCOLIGADA])
    e devolve a lista de linhas (strings) no layout do TXT da folha.
    """
    linhas = []
    tem_codcoligada = COL_CODCOLIGADA in df.columns

    for _, r in df.iterrows():
        codcoligada = str(r[COL_CODCOLIGADA]).strip() if tem_codcoligada and pd.notna(r[COL_CODCOLIGADA]) else "101"

        debito = str(r[COL_DEBITO]).strip()
        credito = str(r[COL_CREDITO]).strip()

        valor = float(r[COL_VALOR]) if pd.notna(r[COL_VALOR]) else 0.0
        valor_centavos = int(round(abs(valor) * 100))
        valor_txt = f"{valor_centavos:010d}"

        descricao = str(r.get(COL_DESC, "")).strip() if pd.notna(r.get(COL_DESC, "")) else ""
        descricao_txt = f"{descricao:<100}"[:100]

        data_txt = ""
        if COL_DATA in df.columns and pd.notna(r[COL_DATA]):
            try:
                data_txt = pd.to_datetime(r[COL_DATA]).strftime("%d%m%Y")
            except Exception:
                data_txt = str(r[COL_DATA])

        linha = "\t".join([codcoligada, "X", debito, credito, valor_txt, descricao_txt, data_txt])
        linhas.append(linha)

    return linhas


# ----------------------------------------------------------------------------
# Parser da "Folha Analítica" (PDF) - relatório de folha com Cod/Descrição/
# N.F./Ref/Valor em DUAS colunas lado a lado por linha.
#
# Usa a posição (x0/x1) de cada palavra na página para separar corretamente
# as colunas, em vez de regex simples no texto corrido (que quebra porque
# a descrição pode conter números, %, barras, pontos, etc.).
# ----------------------------------------------------------------------------
def _clean_pdf_text(t):
    # a fonte do relatório usa um glyph "(cid:3)" para representar espaço
    return t.replace("(cid:3)", " ")


def parse_folha_analitica(path, apenas_total_geral=True):
    """
    Lê um PDF no formato 'FOLHA ANALITICA' (relatório de folha de
    pagamento, com colunas Cod/Descrição/N.F./Ref/Valor duplicadas lado a
    lado) e retorna uma lista de dicts:
        {'cod': '0002', 'descricao': 'SALARIO BASE', 'nf': '48',
         'ref': '1.304,00', 'valor': 123.45, 'is_base': False}

    'is_base' = True indica linhas marcadas com "B" no relatório, que
    representam bases de cálculo (não são lançamentos financeiros reais
    e por padrão são ignoradas na comparação com a base contábil).

    O relatório costuma trazer as páginas de cada Seção E as páginas de
    "TOTAL GERAL" (que repetem os mesmos valores, consolidados). Somar tudo
    contaria cada valor 2x. Por isso, quando o PDF tem páginas de TOTAL GERAL
    só elas são usadas (apenas_total_geral=True); sem elas, usa todas as páginas.
    """
    records = []
    paginas = []   # (é_total_geral, índice_inicial, índice_final) em `records`
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            ini = len(records)
            eh_total = "TOTAL GERAL" in _clean_pdf_text(page.extract_text() or "").upper()
            words = page.extract_words()
            header_tops = sorted(set(round(w["top"], 1) for w in words if w["text"] == "Cod"))
            for htop in header_tops:
                hdr = [w for w in words if abs(w["top"] - htop) < 1]

                def get(label, occ):
                    matches = sorted([w for w in hdr if w["text"] == label], key=lambda w: w["x0"])
                    return matches[occ] if len(matches) > occ else None

                cod1, cod2 = get("Cod", 0), get("Cod", 1)
                desc1, desc2 = get("Descrição", 0), get("Descrição", 1)
                nf1, nf2 = get("N.F.", 0), get("N.F.", 1)
                ref1, ref2 = get("Ref", 0), get("Ref", 1)
                val1, val2 = get("Valor", 0), get("Valor", 1)
                if not all([cod1, cod2, desc1, desc2, nf1, nf2, ref1, ref2, val1, val2]):
                    continue  # não é o cabeçalho esperado (formato diferente)

                def mid(a, b):
                    return (a + b) / 2

                cod1_desc1_mid = mid(cod1["x1"], desc1["x0"])
                desc1_nf1_thresh = nf1["x0"] - 3
                nf1_ref1_mid = mid(nf1["x1"], ref1["x1"])
                ref1_val1_mid = mid(ref1["x1"], val1["x1"])
                val1_flag_mid = mid(val1["x1"], cod2["x0"])
                cod2_thresh = cod2["x0"] - 3
                cod2_desc2_mid = mid(cod2["x1"], desc2["x0"])
                desc2_nf2_thresh = nf2["x0"] - 3
                nf2_ref2_mid = mid(nf2["x1"], ref2["x1"])
                ref2_val2_mid = mid(ref2["x1"], val2["x1"])

                def classify(w):
                    x0, x1 = w["x0"], w["x1"]
                    if x0 < desc1_nf1_thresh:
                        return "cod1" if x0 < cod1_desc1_mid else "desc1"
                    if x1 <= nf1_ref1_mid:
                        return "nf1"
                    if x1 <= ref1_val1_mid:
                        return "ref1"
                    if x1 <= val1_flag_mid:
                        return "valor1"
                    if x0 < cod2_thresh:
                        return "flag"
                    if x0 < desc2_nf2_thresh:
                        return "cod2" if x0 < cod2_desc2_mid else "desc2"
                    if x1 <= nf2_ref2_mid:
                        return "nf2"
                    if x1 <= ref2_val2_mid:
                        return "ref2"
                    return "valor2"

                data_words = [w for w in words if w["top"] > htop + 1]
                rows = {}
                for w in data_words:
                    rows.setdefault(round(w["top"], 1), []).append(w)

                for _, wlist in sorted(rows.items()):
                    wlist.sort(key=lambda w: w["x0"])
                    buckets = {k: [] for k in [
                        "cod1", "desc1", "nf1", "ref1", "valor1", "flag",
                        "cod2", "desc2", "nf2", "ref2", "valor2"
                    ]}
                    for w in wlist:
                        buckets[classify(w)].append(_clean_pdf_text(w["text"]))

                    codv1 = " ".join(buckets["cod1"]).strip()
                    if re.match(r"^\d{3,4}$", codv1):
                        valor_txt = " ".join(buckets["valor1"]).strip()
                        records.append({
                            "cod": codv1.zfill(4),
                            "descricao": " ".join(buckets["desc1"]).strip(),
                            "nf": " ".join(buckets["nf1"]).strip(),
                            "ref": " ".join(buckets["ref1"]).strip(),
                            "valor": _parse_valor_br(valor_txt),
                            "is_base": bool(buckets["flag"]),
                        })
                    codv2 = " ".join(buckets["cod2"]).strip()
                    if re.match(r"^\d{3,4}$", codv2):
                        valor_txt = " ".join(buckets["valor2"]).strip()
                        records.append({
                            "cod": codv2.zfill(4),
                            "descricao": " ".join(buckets["desc2"]).strip(),
                            "nf": " ".join(buckets["nf2"]).strip(),
                            "ref": " ".join(buckets["ref2"]).strip(),
                            "valor": _parse_valor_br(valor_txt),
                            "is_base": False,
                        })
            paginas.append((eh_total, ini, len(records)))

    if apenas_total_geral and any(t for t, _, _ in paginas):
        records = [r for t, a, b in paginas if t for r in records[a:b]]
    return records


def _parse_valor_br(txt):
    """Converte '1.234,56' -> 1234.56. Retorna 0.0 se não conseguir."""
    txt = txt.strip()
    if not txt:
        return 0.0
    try:
        return float(txt.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


# ----------------------------------------------------------------------------
# Parser do Razão (Modelo I) - aceita .xlsx e .ods.
#
# O relatório traz, para uma conta específica, uma sequência de lançamentos
# com as colunas DATA / SEQUENCIA / LOTE / VOUCHER / DOC. NRO. / HISTORICO /
# CENTRO DE CUSTO / CONTA PARTIDA / DEBITO / CREDITO. O cabeçalho real não
# está na primeira linha (tem um bloco de identificação da empresa antes),
# por isso localizamos a linha que contém "HISTORICO" para achar o começo
# da tabela. Linhas de subtotal ("SALDO FINAL" por dia/mês) são descartadas.
#
# A "SEQUENCIA" de cada lançamento é o código de evento do lado do razão
# (equivalente, em espírito, ao COD_EVENTO da base e ao Cod da folha em
# PDF) e é o que usamos para expor o de-para na aba "Comparação 3 Vias".
# ----------------------------------------------------------------------------
def listar_sheets_razao(path):
    """
    Retorna a lista de nomes das planilhas (sheets) de um arquivo de Razão.
    Alguns razões vêm com mais de uma sheet no mesmo arquivo, uma por
    filial (ex.: "101 - MATRIZ", "102 - FILIAL SP" etc.).
    """
    ext = os.path.splitext(path)[1].lower()
    engine = "odf" if ext == ".ods" else None
    xls = pd.ExcelFile(path, engine=engine)
    return xls.sheet_names


def _extrair_evento_cod_conta_partida(cp):
    """
    Extrai o código de evento (conta contábil de contrapartida) da coluna
    CONTA PARTIDA do Razão Modelo I.

    O formato da coluna é "filial - conta_cod", por exemplo:
        "31 - 282"   →  "282"
        "1 - 641"    →  "641"
        "28 - 1365"  →  "1365"
    """
    cp = str(cp).strip()
    partes = cp.split(" - ")
    return partes[-1].strip() if len(partes) >= 2 else cp


def carregar_mapeamento_pemazaweb(path):
    """
    Lê a aba 'PemazaWeb' de CONTAS_E_LANCTO_FOLHA_DE_PAGAMENTO.xlsx e
    devolve um dict  { 'codigo_str': 'NOME DA RUBRICA' }  para uso como
    de-para rápido entre o código numérico de evento (coluna DÉBITO ou
    CRÉDITO do PemazaWeb) e o nome da rubrica (coluna CONTAS).

    Exemplo de entrada (PemazaWeb):
        CONTAS              HISTÓRICO           DÉBITO  CRÉDITO
        PROVISÃO DE FÉRIAS  Provisao de Ferias  648     335
        SALARIO BASE        —                   641     333
        IRRF                IRF S/ SALARIOS     1439    284

    Resultado:
        { '648': 'PROVISÃO DE FÉRIAS',
          '335': 'PROVISÃO DE FÉRIAS',   ← crédito do mesmo evento
          '641': 'SALARIO BASE',
          '333': 'SALARIOS A PAGAR',
          '1439': 'IRRF',
          '284': 'IRRF', ... }

    Quando um mesmo código aparece em mais de um evento (ex.: 641 é
    DÉBITO de SALARIO BASE mas também CRÉDITO de FALTA), o primeiro
    mapeamento encontrado prevalece (ordem de leitura da planilha).
    """
    try:
        df = pd.read_excel(path, sheet_name="PemazaWeb")
    except Exception:
        try:
            # aceita qualquer nome de aba, desde que tenha as colunas certas
            xls = pd.ExcelFile(path)
            df = None
            for sh in xls.sheet_names:
                tentativa = pd.read_excel(path, sheet_name=sh)
                cols_upper = [str(c).strip().upper() for c in tentativa.columns]
                if "CONTAS" in cols_upper and ("DÉBITO" in cols_upper or "DEBITO" in cols_upper):
                    tentativa.columns = cols_upper
                    tentativa = tentativa.rename(columns={"DEBITO": "DÉBITO", "CREDITO": "CRÉDITO"})
                    df = tentativa
                    break
            if df is None:
                return {}
        except Exception:
            return {}

    mapa = {}
    for _, row in df.iterrows():
        rubrica = str(row["CONTAS"]).strip() if pd.notna(row.get("CONTAS")) else ""
        if not rubrica or rubrica.lower() == "nan":
            continue
        for col in ["DÉBITO", "CRÉDITO"]:
            if col in row and pd.notna(row[col]):
                try:
                    cod_str = str(int(row[col]))
                except (ValueError, TypeError):
                    continue
                if cod_str not in mapa:
                    mapa[cod_str] = rubrica
    return mapa


def parse_razao(path, sheet_name=0, pemazaweb_path=None):
    """
    Lê um Razão Modelo I (.xlsx ou .ods) e devolve um DataFrame com as
    colunas:
        SEQUENCIA (int)   — sequência interna do lançamento
        HISTORICO (str)   — histórico do lançamento
        DEBITO (float)    — valor no débito (0 se o lançamento é crédito)
        CREDITO (float)   — valor no crédito (0 se o lançamento é débito)
        EVENTO_COD (str)  — código do evento extraído de CONTA PARTIDA
                            (número após "filial - ", ex.: "282", "641")
        RUBRICA (str)     — nome da rubrica via PemazaWeb (vazio se não
                            houver mapeamento ou se pemazaweb_path=None)

    Apenas linhas de lançamento real são mantidas; linhas de saldo/subtotal
    e o cabeçalho são descartados.

    Parâmetros
    ----------
    path           : caminho para o arquivo do Razão
    sheet_name     : planilha a ler (padrão = 0, primeira)
    pemazaweb_path : caminho para o xlsx com a aba PemazaWeb; quando
                     fornecido, popula a coluna RUBRICA via de-para.
    """
    ext = os.path.splitext(path)[1].lower()
    engine = "odf" if ext == ".ods" else None
    raw = pd.read_excel(path, header=None, engine=engine, sheet_name=sheet_name)

    header_row = None
    for i in range(min(10, len(raw))):
        linha = raw.iloc[i].fillna("").astype(str)
        if linha.str.contains("HISTORICO", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            "Não encontrei o cabeçalho 'HISTORICO' nas primeiras linhas do "
            "arquivo. Verifique se é um Razão no formato Modelo I."
        )

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_row]]

    colunas_necessarias = ["SEQUENCIA", "HISTORICO", "DEBITO", "CREDITO", "DATA"]
    faltando = [c for c in colunas_necessarias if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas não encontradas no razão: {faltando}")

    df = df[df["HISTORICO"].notna()]
    df = df[df["HISTORICO"] != "SALDO FINAL"]
    # mantém só linhas de lançamento real (DATA no formato dd/mm/aaaa)
    df = df[df["DATA"].fillna("").astype(str).str.match(r"^\d{2}/\d{2}/\d{4}$")]

    df["DEBITO"] = pd.to_numeric(df["DEBITO"], errors="coerce").fillna(0.0)
    df["CREDITO"] = pd.to_numeric(df["CREDITO"], errors="coerce").fillna(0.0)
    df["SEQUENCIA"] = pd.to_numeric(df["SEQUENCIA"], errors="coerce")
    df = df.dropna(subset=["SEQUENCIA"]).copy()
    df["SEQUENCIA"] = df["SEQUENCIA"].astype(int)
    df["HISTORICO"] = df["HISTORICO"].astype(str).str.strip()

    # Extrai EVENTO_COD da coluna CONTA PARTIDA ("filial - cod" → "cod")
    if "CONTA PARTIDA" in df.columns:
        df["EVENTO_COD"] = df["CONTA PARTIDA"].apply(_extrair_evento_cod_conta_partida)
    else:
        df["EVENTO_COD"] = ""

    # De-para com PemazaWeb → coluna RUBRICA
    mapa_pema = carregar_mapeamento_pemazaweb(pemazaweb_path) if pemazaweb_path else {}
    df["RUBRICA"] = df["EVENTO_COD"].map(lambda c: mapa_pema.get(str(c), ""))

    cols = ["SEQUENCIA", "HISTORICO", "DEBITO", "CREDITO", "EVENTO_COD", "RUBRICA"]
    return df[cols].reset_index(drop=True)


import unicodedata
import difflib


def _normalizar_texto(s):
    """Maiúsculas, sem acento, sem pontuação, espaços colapsados."""
    if s is None:
        return ""
    s = str(s).upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def montar_crosswalk_evento_cod(df_base, mapa_pema):
    """
    Monta o de-para  EVENTO_COD (código de conta partida do Razão, ex. '641')
    -> COD_EVENTO (código de rubrica da Base/PDF, ex. '0002'), usando a
    RUBRICA (nome) como ponte:

        EVENTO_COD --(mapa_pema)--> RUBRICA --(nome bate com DESCRICAO da
        Base)--> COD_EVENTO

    Sem essa ponte, EVENTO_COD e COD_EVENTO são numerações diferentes e
    nunca batem diretamente — é a causa raiz de lançamentos do Razão
    caírem em "SÓ NO RAZÃO" mesmo quando existem na Base/PDF.

    Retorna:
        crosswalk       : { 'evento_cod_str': 'cod4' }
        rubricas_base    : { 'DESCRICAO_NORMALIZADA': 'cod4' }  (para o
                            fallback por similaridade de texto)
    """
    crosswalk = {}
    rubricas_base = {}

    if df_base is None or "COD_EVENTO" not in df_base.columns or "DESCRICAO" not in df_base.columns:
        return crosswalk, rubricas_base

    cod_num = pd.to_numeric(df_base["COD_EVENTO"], errors="coerce")
    tmp = df_base.assign(_cod_num=cod_num).dropna(subset=["_cod_num"])
    # nome de rubrica (DESCRICAO) -> cod4 mais frequente para essa descrição
    rub_to_cod = (
        tmp.groupby("DESCRICAO")["_cod_num"]
        .agg(lambda s: str(int(s.value_counts().idxmax())).zfill(4))
    )
    rub_to_cod_norm = {_normalizar_texto(k): v for k, v in rub_to_cod.items()}
    rubricas_base = rub_to_cod_norm

    for evento_cod, rubrica in mapa_pema.items():
        cod4 = rub_to_cod_norm.get(_normalizar_texto(rubrica))
        if cod4:
            crosswalk[str(evento_cod)] = cod4

    return crosswalk, rubricas_base


def _seq_para_exibicao(valor):
    """
    Converte SEQUENCIA para exibição na tabela. Pode vir como número único
    (ex.: 21875) ou como string agregada de várias sequências, uma por
    filial (ex.: '1.0, 2.0, 3.0'), produzida quando o razão de múltiplas
    filiais é agrupado por EVENTO_COD. Nesse segundo caso, int() direto
    quebra — aqui tratamos os dois formatos sem lançar exceção.
    """
    try:
        return int(valor)
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    partes = [p.strip() for p in texto.split(",") if p.strip()]
    if not partes:
        return texto
    nums = []
    for p in partes:
        try:
            nums.append(str(int(float(p))))
        except ValueError:
            nums.append(p)
    return ", ".join(nums)


def resolver_cod4_por_similaridade(historico, rubricas_base, limiar=0.72):
    """
    Fallback quando o crosswalk via pemazaweb não resolve: tenta achar a
    rubrica da Base cujo nome tem maior similaridade textual com o
    HISTORICO do lançamento do Razão (ex.: 'DESCONTADO ATRASOS N/FUNC...'
    contém 'ATRASO'). Só devolve resultado se a similaridade for boa o
    bastante (limiar), senão devolve None — para não arriscar casamento
    errado silenciosamente.
    """
    alvo = _normalizar_texto(historico)
    if not alvo or not rubricas_base:
        return None
    melhor_cod, melhor_score = None, 0.0
    for nome_norm, cod4 in rubricas_base.items():
        if not nome_norm:
            continue
        score = difflib.SequenceMatcher(None, alvo, nome_norm).ratio()
        # bônus se o nome da rubrica aparece inteiro dentro do histórico
        if nome_norm in alvo:
            score = max(score, 0.9)
        if score > melhor_score:
            melhor_score, melhor_cod = score, cod4
    return melhor_cod if melhor_score >= limiar else None


class _FopagFrame:
    """
    Constrói a tela de Análise de Conta Contábil x Folha (PDF) dentro de um
    'master' já existente (um tk.Toplevel aberto pelo menu principal, por
    exemplo), em vez de criar sua própria janela raiz (tk.Tk).
    """

    def __init__(self, master):
        self.master = master
        master.title("Análise de Conta Contábil - Débito x Crédito")
        master.geometry("1100x680")
        master.configure(bg=Theme.BG)
        master.minsize(900, 560)

        self.df = None          # base completa importada (xlsx)
        self.file_path = None
        self.pdf_records = None  # lançamentos extraídos do PDF (folha analítica)
        self.pdf_path = None
        self.df_razao = None     # lançamentos extraídos do Razão (xlsx/ods) — filial/total selecionado
        self.razao_path = None
        self.pemazaweb_path = None   # caminho do xlsx com aba PemazaWeb (de-para de rubricas)
        self._df_razao_por_filial = {}  # {nome_da_sheet: DataFrame} — uma entrada por filial
        self._df_divergencias = None
        self._df_tres_vias = None

        self._build_style()
        self._build_layout()

    # ------------------------------------------------------------------
    # Estilo ttk
    # ------------------------------------------------------------------
    def _build_style(self):
        style = ttk.Style(self.master)
        style.theme_use("clam")

        style.configure("TFrame", background=Theme.BG)
        style.configure("Panel.TFrame", background=Theme.BG_PANEL)
        style.configure("Card.TFrame", background=Theme.BG_CARD)

        style.configure("TLabel", background=Theme.BG, foreground=Theme.FG, font=Theme.FONT)
        style.configure("Panel.TLabel", background=Theme.BG_PANEL, foreground=Theme.FG, font=Theme.FONT)
        style.configure("Card.TLabel", background=Theme.BG_CARD, foreground=Theme.FG, font=Theme.FONT)
        style.configure("Muted.TLabel", background=Theme.BG, foreground=Theme.FG_MUTED, font=Theme.FONT)
        style.configure("Title.TLabel", background=Theme.BG, foreground=Theme.FG, font=Theme.FONT_TITLE)
        style.configure("CardTitle.TLabel", background=Theme.BG_CARD, foreground=Theme.FG_MUTED, font=Theme.FONT_BOLD)
        style.configure("CardValue.TLabel", background=Theme.BG_CARD, foreground=Theme.FG, font=Theme.FONT_BIG)

        style.configure("TButton", background=Theme.ACCENT, foreground="white",
                         font=Theme.FONT_BOLD, padding=8, borderwidth=0)
        style.map("TButton", background=[("active", "#3b66d9")])

        style.configure("TEntry", fieldbackground=Theme.BG_CARD, foreground=Theme.FG,
                         insertcolor=Theme.FG, borderwidth=1)

        style.configure("Treeview",
                         background=Theme.BG_CARD, fieldbackground=Theme.BG_CARD,
                         foreground=Theme.FG, rowheight=26, font=Theme.FONT, borderwidth=0)
        style.configure("Treeview.Heading", background=Theme.BG, foreground=Theme.CYAN,
                         font=("Segoe UI", 8, "bold"), relief="flat", padding=(6, 6))
        style.map("Treeview", background=[("selected", Theme.SELECT)],
                  foreground=[("selected", Theme.FG)])

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _build_layout(self):
        # ---- Topo: importação + conta alvo ----
        top = ttk.Frame(self.master, style="Panel.TFrame", padding=15)
        top.pack(fill="x", padx=0, pady=0)

        ttk.Label(top, text="Análise de Conta Contábil", style="Title.TLabel",
                  background=Theme.BG_PANEL).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 10))

        ttk.Button(top, text="📂 Importar Base (.xlsx)", command=self.importar_base).grid(
            row=1, column=0, padx=(0, 10), sticky="w")

        self.lbl_arquivo = ttk.Label(top, text="Nenhum arquivo importado", style="Panel.TLabel",
                                      foreground=Theme.FG_MUTED)
        self.lbl_arquivo.grid(row=1, column=1, columnspan=2, sticky="w")

        ttk.Label(top, text="Conta a analisar:", style="Panel.TLabel").grid(
            row=1, column=3, sticky="e", padx=(20, 6))
        self.entry_conta = ttk.Entry(top, width=16, font=Theme.FONT)
        self.entry_conta.insert(0, "214.04.001")
        self.entry_conta.grid(row=1, column=4, sticky="w")

        ttk.Button(top, text="Calcular", command=self.calcular).grid(row=1, column=5, padx=(10, 0))

        ttk.Separator(top, orient="vertical").grid(row=1, column=6, sticky="ns", padx=15)

        ttk.Button(top, text="📄 Importar Folha (PDF)", command=self.importar_pdf).grid(
            row=1, column=7, padx=(0, 10))
        self.lbl_pdf = ttk.Label(top, text="Nenhum PDF importado", style="Panel.TLabel",
                                  foreground=Theme.FG_MUTED)
        self.lbl_pdf.grid(row=1, column=8, sticky="w")

        ttk.Button(top, text="🔀 Comparar", command=self.comparar_com_pdf).grid(
            row=1, column=9, padx=(10, 0))

        top.columnconfigure(2, weight=1)

        # ---- Segunda linha: Razão + comparação 3 vias ----
        ttk.Button(top, text="📘 Importar Razão (.xlsx/.ods)", command=self.importar_razao).grid(
            row=2, column=0, padx=(0, 10), pady=(8, 0), sticky="w")
        self.lbl_razao = ttk.Label(top, text="Nenhum razão importado", style="Panel.TLabel",
                                    foreground=Theme.FG_MUTED)
        self.lbl_razao.grid(row=2, column=1, columnspan=2, pady=(8, 0), sticky="w")

        # Linha 3: PemazaWeb (de-para de rubricas)
        ttk.Button(top, text="📋 Importar PemazaWeb (.xlsx)", command=self.importar_pemazaweb).grid(
            row=3, column=0, padx=(0, 10), pady=(8, 0), sticky="w")
        self.lbl_pemazaweb = ttk.Label(top, text="PemazaWeb não importado (opcional — enriquece rubricas no Razão)",
                                        style="Panel.TLabel", foreground=Theme.FG_MUTED)
        self.lbl_pemazaweb.grid(row=3, column=1, columnspan=5, pady=(8, 0), sticky="w")

        self.lbl_filial = ttk.Label(top, text="Filial:", style="Panel.TLabel")
        self.lbl_filial.grid(row=2, column=3, sticky="e", padx=(20, 6), pady=(8, 0))
        self.combo_filial = ttk.Combobox(top, width=24, state="disabled", font=Theme.FONT)
        self.combo_filial.grid(row=2, column=4, sticky="w", pady=(8, 0))
        self.combo_filial.bind("<<ComboboxSelected>>", self._aplicar_filial_razao)
        # some até um Razão com mais de uma sheet ser importado
        self.lbl_filial.grid_remove()
        self.combo_filial.grid_remove()

        ttk.Button(top, text="🔗 Comparar 3 Vias (Base × PDF × Razão)",
                   command=self.comparar_tres_vias).grid(
            row=2, column=7, columnspan=3, pady=(8, 0), sticky="w")

        # ---- Cards de resumo da conta selecionada ----
        cards = ttk.Frame(self.master, style="TFrame", padding=(15, 15, 15, 5))
        cards.pack(fill="x")

        self.card_conta = self._make_card(cards, "CONTA", "-")
        self.card_conta.grid(row=0, column=0, padx=(0, 10), sticky="nsew")

        self.card_debito = self._make_card(cards, "TOTAL DÉBITO", "R$ 0,00")
        self.card_debito.grid(row=0, column=1, padx=10, sticky="nsew")

        self.card_credito = self._make_card(cards, "TOTAL CRÉDITO", "R$ 0,00")
        self.card_credito.grid(row=0, column=2, padx=10, sticky="nsew")

        self.card_saldo = self._make_card(cards, "SALDO (DÉBITO - CRÉDITO)", "R$ 0,00")
        self.card_saldo.grid(row=0, column=3, padx=(10, 0), sticky="nsew")

        for i in range(4):
            cards.columnconfigure(i, weight=1)

        # ---- Notebook: Visão por Conta / Lançamentos da conta ----
        notebook_frame = ttk.Frame(self.master, style="TFrame", padding=15)
        notebook_frame.pack(fill="both", expand=True)

        nb_style = ttk.Style(self.master)
        nb_style.configure("TNotebook", background=Theme.BG, borderwidth=0)
        nb_style.configure("TNotebook.Tab", background=Theme.BG_PANEL, foreground=Theme.FG,
                            padding=(14, 8), font=Theme.FONT_BOLD)
        nb_style.map("TNotebook.Tab", background=[("selected", Theme.ACCENT)])

        self.nb = ttk.Notebook(notebook_frame)
        self.nb.pack(fill="both", expand=True)

        # Aba 1: visão por conta
        aba_contas = ttk.Frame(self.nb, style="TFrame")
        self.nb.add(aba_contas, text="Visão por Conta")
        self._build_tree_contas(aba_contas)

        # Aba 2: lançamentos da conta filtrada
        aba_lanc = ttk.Frame(self.nb, style="TFrame")
        self.nb.add(aba_lanc, text="Lançamentos da Conta")
        self._build_tree_lancamentos(aba_lanc)

        # Aba 3: de-para / divergências com o PDF (folha analítica)
        aba_div = ttk.Frame(self.nb, style="TFrame")
        self.nb.add(aba_div, text="Divergências x PDF")
        self._build_tree_divergencias(aba_div)

        # Aba 4: comparação 3 vias (Base x PDF x Razão), restrita à conta alvo
        aba_3vias = ttk.Frame(self.nb, style="TFrame")
        self.nb.add(aba_3vias, text="Comparação 3 Vias")
        self._build_tree_tres_vias(aba_3vias)

        # ---- Rodapé: exportar ----
        footer = ttk.Frame(self.master, style="TFrame", padding=(15, 0, 15, 15))
        footer.pack(fill="x")
        ttk.Button(footer, text="💾 Exportar Visão por Conta (.xlsx)",
                   command=self.exportar_excel).pack(side="right")
        ttk.Button(footer, text="📤 Exportar TXT Folha",
                   command=self.exportar_txt_folha).pack(side="right", padx=(0, 10))

        self.lbl_status = ttk.Label(footer, text="", style="Muted.TLabel")
        self.lbl_status.pack(side="left")

    def _make_card(self, parent, titulo, valor):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=15)
        ttk.Label(frame, text=titulo, style="CardTitle.TLabel").pack(anchor="w")
        lbl_valor = ttk.Label(frame, text=valor, style="CardValue.TLabel")
        lbl_valor.pack(anchor="w", pady=(6, 0))
        frame.lbl_valor = lbl_valor
        return frame

    def _build_tree_contas(self, parent):
        cols = ("conta", "debito", "credito", "saldo", "qtd")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headers = {
            "conta": ("Conta", 140),
            "debito": ("Total Débito", 150),
            "credito": ("Total Crédito", 150),
            "saldo": ("Saldo (Déb - Créd)", 160),
            "qtd": ("Qtd. Lançamentos", 130),
        }
        for c in cols:
            text, width = headers[c]
            tree.heading(c, text=text, command=lambda cc=c: self._sort_tree(tree, cc, False))
            tree.column(c, width=width, anchor="center" if c != "conta" else "w")

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        tree.tag_configure("pos", foreground=Theme.GREEN)
        tree.tag_configure("neg", foreground=Theme.RED)
        tree.tag_configure("alvo", background=Theme.SELECT)

        self.tree_contas = tree

    def _build_tree_lancamentos(self, parent):
        cols = ("data", "debito", "credito", "valor", "descricao", "cod_evento")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headers = {
            "data": ("Data", 100),
            "debito": ("Débito", 110),
            "credito": ("Crédito", 110),
            "valor": ("Valor", 120),
            "descricao": ("Descrição", 320),
            "cod_evento": ("Cód. Evento", 100),
        }
        for c in cols:
            text, width = headers[c]
            tree.heading(c, text=text)
            tree.column(c, width=width, anchor="center" if c != "descricao" else "w")

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree_lanc = tree

    def _build_tree_divergencias(self, parent):
        info = ttk.Label(
            parent,
            text="De-Para por Código de Evento — TODAS as contas da base (COD_EVENTO  ↔  Cod da folha em PDF)",
            style="Muted.TLabel", padding=(0, 0, 0, 8)
        )
        info.pack(anchor="w")

        cols = ("cod", "descricao", "debito", "credito", "valor_base", "valor_pdf", "diferenca", "status")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headers = {
            "cod": ("Cód. Evento", 85),
            "descricao": ("Descrição", 230),
            "debito": ("Conta Débito", 110),
            "credito": ("Conta Crédito", 110),
            "valor_base": ("Valor na Base", 120),
            "valor_pdf": ("Valor no PDF", 120),
            "diferenca": ("Diferença", 120),
            "status": ("Status", 150),
        }
        for c in cols:
            text, width = headers[c]
            tree.heading(c, text=text, command=lambda cc=c: self._sort_tree(tree, cc, False))
            tree.column(c, width=width, anchor="center" if c != "descricao" else "w")

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        tree.tag_configure("ok", foreground=Theme.GREEN)
        tree.tag_configure("divergente", foreground=Theme.RED)
        tree.tag_configure("so_base", foreground=Theme.WARN)
        tree.tag_configure("so_pdf", foreground=Theme.WARN)

        self.tree_div = tree

    def _build_tree_tres_vias(self, parent):
        info = ttk.Label(
            parent,
            text=("Base (COD_EVENTO) × PDF (Cod) × Razão (casado por valor + lado "
                  "Débito/Crédito) — restrito à conta informada em 'Conta a analisar'"),
            style="Muted.TLabel", padding=(0, 0, 0, 8)
        )
        info.pack(anchor="w")

        cols = ("cod", "descricao", "lado", "valor_base", "valor_pdf", "valor_razao",
                 "historico_razao", "evento_cod_razao", "rubrica_razao", "seq_razao", "status")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headers = {
            "cod": ("Cód. Evento", 85),
            "descricao": ("Descrição", 200),
            "lado": ("Lado", 70),
            "valor_base": ("Valor Base", 110),
            "valor_pdf": ("Valor PDF", 110),
            "valor_razao": ("Valor Razão", 110),
            "historico_razao": ("Histórico (Razão)", 260),
            "evento_cod_razao": ("Cód. Conta Partida", 130),
            "rubrica_razao": ("Rubrica PemazaWeb", 200),
            "seq_razao": ("Seq. Razão", 90),
            "status": ("Status", 150),
        }
        for c in cols:
            text, width = headers[c]
            tree.heading(c, text=text, command=lambda cc=c: self._sort_tree(tree, cc, False))
            tree.column(c, width=width,
                        anchor="center" if c not in ("descricao", "historico_razao", "rubrica_razao") else "w")

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        tree.tag_configure("ok", foreground=Theme.GREEN)
        tree.tag_configure("divergente", foreground=Theme.RED)
        tree.tag_configure("parcial", foreground=Theme.WARN)

        self.tree_3vias = tree

    def _sort_tree(self, tree, col, reverse):
        def parse(v):
            try:
                return float(str(v).replace("R$", "").replace(".", "").replace(",", ".").strip())
            except ValueError:
                return v
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        data.sort(key=lambda t: parse(t[0]), reverse=reverse)
        for i, (_, k) in enumerate(data):
            tree.move(k, "", i)
        tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))

    # ------------------------------------------------------------------
    # Ações
    # ------------------------------------------------------------------
    def importar_base(self):
        path = filedialog.askopenfilename(
            title="Selecione a base exportada",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not path:
            return
        try:
            df = pd.read_excel(path, sheet_name=0, header=0)
        except Exception as e:
            messagebox.showerror("Erro ao importar", f"Não foi possível ler o arquivo:\n{e}")
            return

        # Normaliza nomes de colunas (remove espaços extras)
        df.columns = [str(c).strip() for c in df.columns]

        colunas_necessarias = [COL_DEBITO, COL_CREDITO, COL_VALOR]
        faltando = [c for c in colunas_necessarias if c not in df.columns]
        if faltando:
            messagebox.showerror(
                "Colunas não encontradas",
                f"As colunas obrigatórias não foram encontradas na planilha: {faltando}\n\n"
                f"Colunas encontradas: {list(df.columns)}"
            )
            return

        # Remove linhas de rodapé/vazias (ex.: linha "Filtros aplicados...")
        df = df.dropna(subset=[COL_DEBITO, COL_CREDITO, COL_VALOR]).copy()
        df[COL_DEBITO] = df[COL_DEBITO].astype(str).str.strip()
        df[COL_CREDITO] = df[COL_CREDITO].astype(str).str.strip()
        df[COL_VALOR] = pd.to_numeric(df[COL_VALOR], errors="coerce")
        df = df.dropna(subset=[COL_VALOR])

        self.df = df
        self.file_path = path
        self.lbl_arquivo.configure(text=f"{os.path.basename(path)}  ({len(df)} lançamentos)")
        self.lbl_status.configure(text="Base importada com sucesso. Informe a conta e clique em Calcular.")

        # calcula automaticamente com a conta default já preenchida
        self.calcular()

    def importar_pemazaweb(self):
        path = filedialog.askopenfilename(
            title="Selecione o arquivo com a aba PemazaWeb",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not path:
            return
        try:
            mapa = carregar_mapeamento_pemazaweb(path)
        except Exception as e:
            messagebox.showerror("Erro ao importar PemazaWeb", f"Não foi possível ler o arquivo:\n{e}")
            return
        if not mapa:
            messagebox.showwarning(
                "Aba não encontrada",
                "Não encontrei a aba 'PemazaWeb' no arquivo selecionado ou ela está vazia.\n"
                "Verifique o nome da aba e as colunas CONTAS, DÉBITO e CRÉDITO."
            )
            return
        self.pemazaweb_path = path
        self.lbl_pemazaweb.configure(
            text=f"{os.path.basename(path)}  ({len(mapa)} código(s) mapeados)",
            foreground=Theme.FG
        )
        self.lbl_status.configure(text="PemazaWeb importado. Re-importe o Razão para aplicar o de-para de rubricas.")

    def calcular(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Importe uma base antes de calcular.")
            return

        conta = self.entry_conta.get().strip()
        if not conta:
            messagebox.showwarning("Atenção", "Informe a conta que deseja analisar.")
            return

        df = self.df

        # ---- Resumo da conta alvo ----
        deb_mask = df[COL_DEBITO] == conta
        cred_mask = df[COL_CREDITO] == conta
        total_debito = df.loc[deb_mask, COL_VALOR].sum()
        total_credito = df.loc[cred_mask, COL_VALOR].sum()
        saldo = total_debito - total_credito

        if not deb_mask.any() and not cred_mask.any():
            self.lbl_status.configure(
                text=f"Atenção: a conta '{conta}' não foi encontrada na base importada.")

        self.card_conta.lbl_valor.configure(text=conta if (deb_mask.any() or cred_mask.any()) else f"{conta} (não encontrada)")
        self.card_debito.lbl_valor.configure(text=self._fmt(total_debito))
        self.card_credito.lbl_valor.configure(text=self._fmt(total_credito))
        self.card_saldo.lbl_valor.configure(
            text=self._fmt(saldo),
            foreground=Theme.GREEN if saldo >= 0 else Theme.RED
        )

        # ---- Lançamentos da conta alvo (débito ou crédito) ----
        for row in self.tree_lanc.get_children():
            self.tree_lanc.delete(row)

        sub = df[deb_mask | cred_mask].copy()
        if COL_DATA in sub.columns:
            sub = sub.sort_values(COL_DATA)
        for _, r in sub.iterrows():
            data_fmt = ""
            if COL_DATA in sub.columns and pd.notna(r[COL_DATA]):
                try:
                    data_fmt = pd.to_datetime(r[COL_DATA]).strftime("%d/%m/%Y")
                except Exception:
                    data_fmt = str(r[COL_DATA])
            self.tree_lanc.insert("", "end", values=(
                data_fmt,
                r[COL_DEBITO],
                r[COL_CREDITO],
                self._fmt(r[COL_VALOR]),
                r.get(COL_DESC, ""),
                r.get("COD_EVENTO", ""),
            ))

        # ---- Visão consolidada por conta (todas as contas da base) ----
        self._calcular_visao_por_conta(conta)

        self.lbl_status.configure(
            text=f"Cálculo realizado para a conta {conta} — {len(sub)} lançamento(s) encontrado(s)."
        )

    def importar_pdf(self):
        path = filedialog.askopenfilename(
            title="Selecione a Folha Analítica (PDF)",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not path:
            return
        try:
            records = parse_folha_analitica(path)
        except Exception as e:
            messagebox.showerror("Erro ao importar PDF", f"Não foi possível ler o PDF:\n{e}")
            return

        if not records:
            messagebox.showwarning(
                "Nenhum lançamento encontrado",
                "Não consegui reconhecer o layout esperado (Cod/Descrição/N.F./Ref/Valor) "
                "neste PDF. Verifique se é uma Folha Analítica no formato padrão."
            )
            return

        self.pdf_records = records
        self.pdf_path = path
        qtd_base = sum(1 for r in records if not r["is_base"])
        self.lbl_pdf.configure(text=f"{os.path.basename(path)}  ({qtd_base} eventos, {len(records)-qtd_base} bases)")
        self.lbl_status.configure(text="PDF importado. Clique em Comparar para gerar o de-para.")

    def importar_razao(self):
        path = filedialog.askopenfilename(
            title="Selecione o Razão (Modelo I)",
            filetypes=[("Planilhas", "*.xlsx *.xls *.ods")]
        )
        if not path:
            return

        try:
            sheets = listar_sheets_razao(path)
        except Exception as e:
            messagebox.showerror("Erro ao importar Razão", f"Não foi possível ler o arquivo:\n{e}")
            return

        # tenta interpretar cada sheet como um Razão Modelo I (uma sheet por
        # filial); sheets que não têm o formato esperado são ignoradas
        df_por_filial = {}
        erros = []
        for sh in sheets:
            try:
                df_sh = parse_razao(path, sheet_name=sh, pemazaweb_path=self.pemazaweb_path)
            except Exception as e:
                erros.append(f"{sh}: {e}")
                continue
            if not df_sh.empty:
                df_por_filial[str(sh)] = df_sh

        if not df_por_filial:
            detalhe = "\n".join(erros)
            messagebox.showwarning(
                "Nenhum lançamento encontrado",
                "Não encontrei linhas de lançamento no formato esperado (Razão Modelo I) "
                "em nenhuma planilha do arquivo." + (f"\n\nDetalhes:\n{detalhe}" if detalhe else "")
            )
            return

        self._df_razao_por_filial = df_por_filial
        self.razao_path = path

        filiais = list(df_por_filial.keys())
        if len(filiais) > 1:
            # múltiplas filiais: mostra o seletor, com "TOTAL" como padrão
            opcoes = filiais + ["TOTAL (todas as filiais)"]
            self.combo_filial.configure(values=opcoes, state="readonly")
            self.combo_filial.set("TOTAL (todas as filiais)")
            self.lbl_filial.grid()
            self.combo_filial.grid()
        else:
            # uma única filial/sheet: não há o que selecionar
            self.combo_filial.configure(values=filiais, state="disabled")
            self.combo_filial.set(filiais[0])
            self.lbl_filial.grid_remove()
            self.combo_filial.grid_remove()

        self._aplicar_filial_razao()

    def _aplicar_filial_razao(self, event=None):
        """
        Aplica a filial (ou o total consolidado) selecionada no combobox,
        atualizando self.df_razao — o DataFrame efetivamente usado na
        Comparação 3 Vias — e o texto de status ao lado do botão.
        """
        if not self._df_razao_por_filial:
            return

        sel = self.combo_filial.get()

        if sel.startswith("TOTAL"):
            partes = []
            for filial, df_f in self._df_razao_por_filial.items():
                df_c = df_f.copy()
                df_c["FILIAL"] = filial
                partes.append(df_c)
            self.df_razao = pd.concat(partes, ignore_index=True)
            desc_filial = f"TOTAL — {len(self._df_razao_por_filial)} filial(is)"
        else:
            df_f = self._df_razao_por_filial.get(sel)
            if df_f is None:
                return
            df_c = df_f.copy()
            df_c["FILIAL"] = sel
            self.df_razao = df_c
            desc_filial = f"Filial: {sel}"

        self.lbl_razao.configure(
            text=f"{os.path.basename(self.razao_path)} — {desc_filial}  "
                 f"({len(self.df_razao)} lançamento(s))"
        )
        self.lbl_status.configure(
            text="Razão importado. Clique em 'Comparar 3 Vias' para gerar o de-para."
        )

    def comparar_com_pdf(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Importe a base (.xlsx) antes de comparar.")
            return
        if not self.pdf_records:
            messagebox.showwarning("Atenção", "Importe o PDF da Folha Analítica antes de comparar.")
            return

        df = self.df
        if "COD_EVENTO" not in df.columns:
            messagebox.showerror(
                "Coluna não encontrada",
                "A base importada não possui a coluna 'COD_EVENTO', necessária para "
                "cruzar com o código de evento (Cod) do PDF."
            )
            return

        # agrega TODA a base (todas as contas) por código de evento
        base = df.dropna(subset=["COD_EVENTO"]).copy()
        base["_cod4"] = base["COD_EVENTO"].apply(lambda v: str(int(v)).zfill(4))
        base_group = base.groupby("_cod4").agg(
            valor=(COL_VALOR, "sum"),
            descricao=(COL_DESC, "first"),
            contas_debito=(COL_DEBITO, lambda s: ", ".join(sorted(set(s)))),
            contas_credito=(COL_CREDITO, lambda s: ", ".join(sorted(set(s)))),
        )

        # agrega o PDF por código, ignorando linhas de base de cálculo (flag "B")
        pdf_group = {}
        for r in self.pdf_records:
            if r["is_base"]:
                continue
            cod = r["cod"]
            if cod not in pdf_group:
                pdf_group[cod] = {"valor": 0.0, "descricao": r["descricao"]}
            pdf_group[cod]["valor"] += r["valor"]

        todos_codigos = sorted(set(base_group.index) | set(pdf_group.keys()))

        for row in self.tree_div.get_children():
            self.tree_div.delete(row)

        self._df_divergencias = []
        tolerancia = 0.01

        for cod in todos_codigos:
            tem_base = cod in base_group.index
            tem_pdf = cod in pdf_group

            valor_base = float(base_group.loc[cod, "valor"]) if tem_base else None
            valor_pdf = float(pdf_group[cod]["valor"]) if tem_pdf else None
            descricao = (base_group.loc[cod, "descricao"] if tem_base
                         else pdf_group[cod]["descricao"])
            conta_deb = base_group.loc[cod, "contas_debito"] if tem_base else ""
            conta_cred = base_group.loc[cod, "contas_credito"] if tem_base else ""

            if tem_base and tem_pdf:
                diferenca = valor_base - valor_pdf
                if abs(diferenca) <= tolerancia:
                    status, tag = "OK", "ok"
                else:
                    status, tag = "DIVERGENTE", "divergente"
            elif tem_base and not tem_pdf:
                diferenca = valor_base
                status, tag = "SÓ NA BASE (XLSX)", "so_base"
            else:
                diferenca = -valor_pdf
                status, tag = "SÓ NO PDF", "so_pdf"

            self.tree_div.insert("", "end", values=(
                cod,
                descricao,
                conta_deb or "—",
                conta_cred or "—",
                self._fmt(valor_base) if tem_base else "—",
                self._fmt(valor_pdf) if tem_pdf else "—",
                self._fmt(diferenca),
                status,
            ), tags=(tag,))

            self._df_divergencias.append({
                "Cód. Evento": cod,
                "Descrição": descricao,
                "Conta Débito": conta_deb,
                "Conta Crédito": conta_cred,
                "Valor na Base": valor_base if tem_base else None,
                "Valor no PDF": valor_pdf if tem_pdf else None,
                "Diferença": diferenca,
                "Status": status,
            })

        n_ok = sum(1 for d in self._df_divergencias if d["Status"] == "OK")
        n_div = sum(1 for d in self._df_divergencias if d["Status"] == "DIVERGENTE")
        n_so_base = sum(1 for d in self._df_divergencias if d["Status"] == "SÓ NA BASE (XLSX)")
        n_so_pdf = sum(1 for d in self._df_divergencias if d["Status"] == "SÓ NO PDF")

        self.nb.select(2)
        self.lbl_status.configure(
            text=(f"Comparação concluída — {n_ok} OK · {n_div} divergente(s) · "
                  f"{n_so_base} só na base · {n_so_pdf} só no PDF.")
        )

    def comparar_tres_vias(self):
        """
        Cruza, para a conta informada em 'Conta a analisar':
          - Base (xlsx): agregada por COD_EVENTO + lado (débito/crédito) em
            que a conta alvo aparece;
          - PDF (Folha): agregado por Cod (código de evento), ignorando
            linhas de base de cálculo ("B");
          - Razão: cada lançamento é casado (1-para-1, guloso) com a
            base/PDF por valor + lado, já que o razão não traz um código de
            evento equivalente ao COD_EVENTO — só a SEQUENCIA interna do
            lançamento contábil, que expomos como "Cód. Evento (Razão)".

        Linhas do razão que não encontram par em nenhum código (ex.: valores
        líquidos/consolidados que o Protheus lança de forma agregada) também
        aparecem no resultado, marcadas como "SÓ NO RAZÃO".
        """
        if self.df is None:
            messagebox.showwarning("Atenção", "Importe a base (.xlsx) antes de comparar.")
            return
        if not self.pdf_records:
            messagebox.showwarning("Atenção", "Importe o PDF da Folha antes de comparar.")
            return
        if self.df_razao is None:
            messagebox.showwarning("Atenção", "Importe o Razão (.xlsx/.ods) antes de comparar.")
            return

        conta = self.entry_conta.get().strip()
        if not conta:
            messagebox.showwarning("Atenção", "Informe a conta que deseja analisar.")
            return

        df = self.df
        if "COD_EVENTO" not in df.columns:
            messagebox.showerror(
                "Coluna não encontrada",
                "A base importada não possui a coluna 'COD_EVENTO', necessária para "
                "cruzar com o PDF e com o Razão."
            )
            return

        tolerancia = 0.01

        # ---- 1) Base: lançamentos da conta alvo, agregados por cod + lado ----
        sub = df[(df[COL_DEBITO] == conta) | (df[COL_CREDITO] == conta)].copy()
        sub = sub.dropna(subset=["COD_EVENTO"])
        sub["_cod4"] = sub["COD_EVENTO"].apply(lambda v: str(int(v)).zfill(4))
        sub["_lado"] = sub.apply(lambda r: "DEBITO" if r[COL_DEBITO] == conta else "CREDITO", axis=1)
        base_group = sub.groupby(["_cod4", "_lado"]).agg(
            valor=(COL_VALOR, "sum"),
            descricao=(COL_DESC, "first"),
        )

        # ---- 2) PDF agregado por cod (ignora linhas "B" de base de cálculo) ----
        pdf_group = {}
        for r in self.pdf_records:
            if r["is_base"]:
                continue
            cod = r["cod"]
            if cod not in pdf_group:
                pdf_group[cod] = {"valor": 0.0, "descricao": r["descricao"]}
            pdf_group[cod]["valor"] += r["valor"]

        # ---- 3) Razão: agrega por EVENTO_COD + lado para comparação correta ----
        # Quando o razão vem de múltiplas filiais (TOTAL), o mesmo evento aparece
        # N vezes (uma por filial). Agrupar por EVENTO_COD soma todos os valores,
        # tornando a comparação equivalente à agregação já feita na Base e no PDF.
        # Quando só há uma filial, o resultado é idêntico ao comportamento anterior.
        raz = self.df_razao

        def _agregar_razao_por_evento(df_raz, col_valor):
            """
            Agrega o razão por EVENTO_COD somando col_valor.
            Retorna DataFrame com: EVENTO_COD, col_valor (soma), HISTORICO (primeiro),
            RUBRICA (primeiro não-vazio), SEQUENCIA (lista como string).
            """
            if "EVENTO_COD" not in df_raz.columns:
                # Sem EVENTO_COD não é possível agrupar — retorna pool individual
                sub = df_raz[df_raz[col_valor] > 0].copy()
                sub["used"] = False
                return sub

            sub = df_raz[df_raz[col_valor] > 0].copy()
            if sub.empty:
                return sub

            grp = sub.groupby("EVENTO_COD", sort=False).agg(
                **{
                    col_valor: (col_valor, "sum"),
                    "HISTORICO": ("HISTORICO", "first"),
                    "RUBRICA": ("RUBRICA", lambda s: next((v for v in s if v), "")),
                    "SEQUENCIA": ("SEQUENCIA", lambda s: ", ".join(str(v) for v in s)),
                }
            ).reset_index()
            grp["used"] = False
            return grp

        deb_pool = _agregar_razao_por_evento(raz, "DEBITO")
        cred_pool = _agregar_razao_por_evento(raz, "CREDITO")
        pools = {"DEBITO": (deb_pool, "DEBITO"), "CREDITO": (cred_pool, "CREDITO")}

        # ---- 3b) Crosswalk EVENTO_COD (razão) -> COD_EVENTO (base/pdf) ----
        # Ponte via nome da rubrica (pemazaweb + descrição da base). Resolve
        # a maior parte dos casos sem depender de casamento por valor.
        mapa_pema = carregar_mapeamento_pemazaweb(self.pemazaweb_path) if self.pemazaweb_path else {}
        crosswalk, rubricas_base = montar_crosswalk_evento_cod(df, mapa_pema)

        def _cod4_para_linha_razao(row):
            evt = str(row.get("EVENTO_COD", "")).strip()
            cod4 = crosswalk.get(evt)
            if cod4:
                return cod4
            # fallback: similaridade de texto entre HISTORICO e nomes de rubrica da base
            return resolver_cod4_por_similaridade(row.get("HISTORICO", ""), rubricas_base)

        for pool, _ in pools.values():
            if not pool.empty and "EVENTO_COD" in pool.columns:
                pool["COD4_CROSSWALK"] = pool.apply(_cod4_para_linha_razao, axis=1)
            elif not pool.empty:
                pool["COD4_CROSSWALK"] = None

        cods_base = set(cod for cod, _ in base_group.index)
        todos_cods = sorted(cods_base | set(pdf_group.keys()))

        resultados = []
        for cod in todos_cods:
            lados = set(l for c, l in base_group.index if c == cod) or {"DEBITO", "CREDITO"}
            # só considera lados em que exista base OU pdf (evita duplicar quando
            # o cod não existe na base, caso em que testamos os dois lados)
            for lado in lados:
                tem_base = (cod, lado) in base_group.index
                valor_base = float(base_group.loc[(cod, lado), "valor"]) if tem_base else None
                valor_pdf = pdf_group.get(cod, {}).get("valor")
                if not tem_base and valor_pdf is None:
                    continue
                if not tem_base and lado not in ("DEBITO", "CREDITO"):
                    continue
                descricao = (base_group.loc[(cod, lado), "descricao"] if tem_base
                             else pdf_group.get(cod, {}).get("descricao", ""))

                pool, col = pools[lado]

                # Casamento primário: por EVENTO_COD (código da conta partida).
                # O pool já está agregado por EVENTO_COD, então um evento que
                # aparece em N filiais aparece como UMA linha com a soma total —
                # isso garante que o comparativo multi-filial some tudo corretamente.
                # Fallback por valor mantido para pools sem EVENTO_COD.
                historico_razao, seq_razao, valor_razao = None, None, None
                evento_cod_razao, rubrica_razao = None, None
                metodo_match = None

                cod_num = str(int(cod)) if cod.isdigit() else cod

                # Critério 1 (preferencial): crosswalk EVENTO_COD -> COD_EVENTO
                # via rubrica. Resolve a diferença de numeração entre Razão e
                # Base/PDF, que é a causa raiz de sobras em "SÓ NO RAZÃO".
                if "COD4_CROSSWALK" in pool.columns:
                    cand = pool[(~pool["used"]) & (pool["COD4_CROSSWALK"] == cod)]
                    if not cand.empty:
                        idx = cand.index[0]
                        pool.loc[idx, "used"] = True
                        historico_razao = pool.loc[idx, "HISTORICO"]
                        seq_razao = str(pool.loc[idx, "SEQUENCIA"])
                        valor_razao = float(pool.loc[idx, col])
                        evento_cod_razao = pool.loc[idx, "EVENTO_COD"]
                        rubrica_razao = pool.loc[idx, "RUBRICA"] if "RUBRICA" in pool.columns else None
                        metodo_match = "crosswalk"

                # Critério 2: EVENTO_COD cru bate direto com o cod (casos raros
                # em que as duas numerações coincidem)
                if valor_razao is None and "EVENTO_COD" in pool.columns:
                    cand = pool[(~pool["used"]) & (pool["EVENTO_COD"].astype(str) == cod_num)]
                    if not cand.empty:
                        idx = cand.index[0]
                        pool.loc[idx, "used"] = True
                        historico_razao = pool.loc[idx, "HISTORICO"]
                        seq_razao = str(pool.loc[idx, "SEQUENCIA"])   # pode ser "seq1, seq2, ..."
                        valor_razao = float(pool.loc[idx, col])
                        evento_cod_razao = pool.loc[idx, "EVENTO_COD"]
                        rubrica_razao = pool.loc[idx, "RUBRICA"] if "RUBRICA" in pool.columns else None
                        metodo_match = "evento_cod_direto"

                # Fallback: casamento por valor (quando pool não tem EVENTO_COD)
                if valor_razao is None:
                    for alvo in [v for v in (valor_base, valor_pdf) if v is not None]:
                        cand = pool[(~pool["used"]) & (pool[col].round(2) == round(alvo, 2))]
                        if not cand.empty:
                            idx = cand.index[0]
                            pool.loc[idx, "used"] = True
                            historico_razao = pool.loc[idx, "HISTORICO"]
                            seq_razao = str(pool.loc[idx, "SEQUENCIA"])
                            valor_razao = float(pool.loc[idx, col])
                            evento_cod_razao = pool.loc[idx, "EVENTO_COD"] if "EVENTO_COD" in pool.columns else None
                            rubrica_razao = pool.loc[idx, "RUBRICA"] if "RUBRICA" in pool.columns else None
                            metodo_match = "valor"
                            break

                status = self._status_3vias(valor_base, valor_pdf, valor_razao, tolerancia)
                resultados.append({
                    "Cód. Evento": cod,
                    "Descrição": descricao,
                    "Lado": lado,
                    "Valor Base": valor_base,
                    "Valor PDF": valor_pdf,
                    "Valor Razão": valor_razao,
                    "Histórico (Razão)": historico_razao,
                    "Cód. Conta Partida": evento_cod_razao,
                    "Rubrica PemazaWeb": rubrica_razao,
                    "Seq. Razão": seq_razao,
                    "Status": status,
                    "Método Match": metodo_match,
                })

        # ---- 4) Sobras do razão (lançamentos que não casaram com nada) ----
        for lado, (pool, col) in pools.items():
            for _, r in pool[~pool["used"]].iterrows():
                resultados.append({
                    "Cód. Evento": "-",
                    "Descrição": r["HISTORICO"],
                    "Lado": lado,
                    "Valor Base": None,
                    "Valor PDF": None,
                    "Valor Razão": float(r[col]),
                    "Histórico (Razão)": r["HISTORICO"],
                    "Cód. Conta Partida": r["EVENTO_COD"] if "EVENTO_COD" in r.index else None,
                    "Rubrica PemazaWeb": r["RUBRICA"] if "RUBRICA" in r.index else None,
                    "Seq. Razão": _seq_para_exibicao(r["SEQUENCIA"]),
                    "Status": "SÓ NO RAZÃO",
                    "Método Match": None,
                })

        self._df_tres_vias = resultados

        for row in self.tree_3vias.get_children():
            self.tree_3vias.delete(row)

        for d in resultados:
            tag = {"OK": "ok", "DIVERGENTE": "divergente"}.get(d["Status"], "parcial")
            self.tree_3vias.insert("", "end", values=(
                d["Cód. Evento"],
                d["Descrição"],
                d["Lado"],
                self._fmt(d["Valor Base"]) if d["Valor Base"] is not None else "—",
                self._fmt(d["Valor PDF"]) if d["Valor PDF"] is not None else "—",
                self._fmt(d["Valor Razão"]) if d["Valor Razão"] is not None else "—",
                d["Histórico (Razão)"] or "—",
                d.get("Cód. Conta Partida") or "—",
                d.get("Rubrica PemazaWeb") or "—",
                d.get("Seq. Razão") if d.get("Seq. Razão") is not None else "—",
                d["Status"],
            ), tags=(tag,))

        n_ok = sum(1 for d in resultados if d["Status"] == "OK")
        n_div = sum(1 for d in resultados if d["Status"] == "DIVERGENTE")
        n_parcial = len(resultados) - n_ok - n_div

        self.nb.select(3)
        self.lbl_status.configure(
            text=(f"Comparação 3 vias concluída para a conta {conta} — "
                  f"{n_ok} OK · {n_div} divergente(s) · {n_parcial} parcial(is)/só em uma fonte.")
        )

    @staticmethod
    def _status_3vias(valor_base, valor_pdf, valor_razao, tolerancia):
        presentes = [v for v in (valor_base, valor_pdf, valor_razao) if v is not None]
        if len(presentes) < 2:
            if valor_base is not None:
                return "SÓ NA BASE"
            if valor_pdf is not None:
                return "SÓ NO PDF"
            return "SÓ NO RAZÃO"
        if max(presentes) - min(presentes) <= tolerancia:
            if len(presentes) == 3:
                return "OK"
            faltando = "RAZÃO" if valor_razao is None else ("PDF" if valor_pdf is None else "BASE")
            return f"OK (SEM {faltando})"
        return "DIVERGENTE"

    def _calcular_visao_por_conta(self, conta_alvo):
        df = self.df

        deb_group = df.groupby(COL_DEBITO)[COL_VALOR].agg(["sum", "count"])
        cred_group = df.groupby(COL_CREDITO)[COL_VALOR].agg(["sum", "count"])

        todas_contas = sorted(set(deb_group.index) | set(cred_group.index))

        for row in self.tree_contas.get_children():
            self.tree_contas.delete(row)

        self._df_visao_contas = []  # para exportação

        for conta in todas_contas:
            total_deb = deb_group["sum"].get(conta, 0.0)
            qtd_deb = int(deb_group["count"].get(conta, 0))
            total_cred = cred_group["sum"].get(conta, 0.0)
            qtd_cred = int(cred_group["count"].get(conta, 0))
            saldo = total_deb - total_cred
            qtd_total = qtd_deb + qtd_cred

            tag = "pos" if saldo >= 0 else "neg"
            tags = (tag,)
            if conta == conta_alvo:
                tags = (tag, "alvo")

            self.tree_contas.insert("", "end", values=(
                conta,
                self._fmt(total_deb),
                self._fmt(total_cred),
                self._fmt(saldo),
                qtd_total,
            ), tags=tags)

            self._df_visao_contas.append({
                "Conta": conta,
                "Total Débito": total_deb,
                "Total Crédito": total_cred,
                "Saldo (Débito - Crédito)": saldo,
                "Qtd. Lançamentos": qtd_total,
            })

    def exportar_txt_folha(self):
        if self.df is None:
            messagebox.showwarning("Atenção", "Importe a base (.xlsx) antes de exportar.")
            return

        colunas_necessarias = [COL_DEBITO, COL_CREDITO, COL_VALOR]
        faltando = [c for c in colunas_necessarias if c not in self.df.columns]
        if faltando:
            messagebox.showerror(
                "Colunas não encontradas",
                f"A base importada não possui as colunas necessárias para exportar o TXT: {faltando}"
            )
            return

        path = filedialog.asksaveasfilename(
            title="Salvar TXT da Folha",
            defaultextension=".txt",
            filetypes=[("Arquivo de texto", "*.txt")],
            initialfile="FOLHA.txt",
        )
        if not path:
            return

        try:
            linhas = gerar_linhas_txt_folha(self.df)
            with open(path, "w", encoding="ascii", errors="replace", newline="\r\n") as f:
                f.write("\n".join(linhas))
                if linhas:
                    f.write("\n")
            messagebox.showinfo("Exportado", f"Arquivo TXT salvo em:\n{path}\n\n{len(linhas)} linha(s) gerada(s).")
            self.lbl_status.configure(text=f"TXT exportado com sucesso — {len(linhas)} linha(s).")
        except Exception as e:
            messagebox.showerror("Erro ao exportar TXT", str(e))

    def exportar_excel(self):
        if not getattr(self, "_df_visao_contas", None):
            messagebox.showwarning("Atenção", "Calcule a análise antes de exportar.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar visão por conta",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile="visao_por_conta.xlsx",
        )
        if not path:
            return

        try:
            out_df = pd.DataFrame(self._df_visao_contas)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                out_df.to_excel(writer, sheet_name="Visao_por_Conta", index=False)
                if self._df_divergencias:
                    pd.DataFrame(self._df_divergencias).to_excel(
                        writer, sheet_name="Divergencias_x_PDF", index=False)
                if self._df_tres_vias:
                    pd.DataFrame(self._df_tres_vias).to_excel(
                        writer, sheet_name="Comparacao_3_Vias", index=False)
            messagebox.showinfo("Exportado", f"Arquivo salvo em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    @staticmethod
    def _fmt(valor):
        try:
            v = float(valor)
        except (TypeError, ValueError):
            return "R$ 0,00"
        sinal = "-" if v < 0 else ""
        v = abs(v)
        return f"{sinal}R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _build_content(win):
    """
    Ponto de entrada usado pelo menu principal:
        import fopag_v2
        fopag_v2._build_content(win)   # win = tk.Toplevel(root)
    """
    return _FopagFrame(win)


if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()