"""
protheus/excel_massa.py — Input em massa por planilha Excel.

  gerar_modelo(caminho)   grava o modelo (.xlsx) para preencher
  ler_planilha(caminho)   lê a planilha preenchida e devolve (lançamentos, erros)

Cada linha passa pelo mesmo validar() do lançamento digitado, então o TXT gerado
segue exatamente as mesmas regras. Sem Tkinter (testável).
"""
import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from protheus.txt_layout import validar

ABA = "Lançamentos"
ABA_AJUDA = "Instruções"

# (chave do lançamento, cabeçalho do modelo, largura, formato da célula)
COLUNAS = [
    ("lp",          "Lançamento Padrão",       18, "@"),
    ("data",        "Data",                    13, "DD/MM/YYYY"),
    ("conta_deb",   "Conta Débito",            15, "@"),
    ("conta_cred",  "Conta Crédito",           15, "@"),
    ("valor",       "Valor",                   16, "#,##0.00"),
    ("filial_deb",  "Filial Débito",           14, "@"),
    ("filial_cred", "Filial Crédito",          14, "@"),
    ("cc_deb",      "Centro de Custo Débito",  24, "@"),
    ("cc_cred",     "Centro de Custo Crédito", 24, "@"),
    ("item_deb",    "Item Contábil Débito",    22, "@"),
    ("item_cred",   "Item Contábil Crédito",   22, "@"),
    ("historico",   "Histórico",               70, "@"),
]
OBRIGATORIAS = ("lp", "data", "conta_deb", "conta_cred", "valor", "filial_deb", "filial_cred", "historico")
LINHAS_FORMATADAS = 5000      # linhas do modelo já formatadas (texto, data, valor)

# nomes aceitos no cabeçalho (sem acento, minúsculos, só letras e números)
_APELIDOS = {
    "lancamentopadrao": "lp", "lp": "lp", "lanctopadrao": "lp", "codigolp": "lp",
    "data": "data", "datalancamento": "data",
    "contadebito": "conta_deb", "contadeb": "conta_deb", "debito": "conta_deb",
    "contacredito": "conta_cred", "contacred": "conta_cred", "credito": "conta_cred",
    "valor": "valor", "valorrs": "valor",
    "filialdebito": "filial_deb", "filialdeb": "filial_deb",
    "filialcredito": "filial_cred", "filialcred": "filial_cred",
    "centrodecustodebito": "cc_deb", "centrocustodebito": "cc_deb", "ccdebito": "cc_deb", "ccdeb": "cc_deb",
    "ccustodebito": "cc_deb",
    "centrodecustocredito": "cc_cred", "centrocustocredito": "cc_cred", "cccredito": "cc_cred",
    "cccred": "cc_cred", "ccustocredito": "cc_cred",
    "itemcontabildebito": "item_deb", "itemcontabildeb": "item_deb", "itemdebito": "item_deb", "itemdeb": "item_deb",
    "itemcontabilcredito": "item_cred", "itemcontabilcred": "item_cred", "itemcredito": "item_cred", "itemcred": "item_cred",
    "historico": "historico", "hist": "historico",
}


def _norm(txt) -> str:
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())


# ───────────────────────────── modelo ─────────────────────────────
def gerar_modelo(caminho: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = ABA

    azul = PatternFill("solid", fgColor="294C60")
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    for c, (_k, titulo, larg, fmt) in enumerate(COLUNAS, 1):
        cel = ws.cell(1, c, titulo)
        cel.font = Font(name="Segoe UI", bold=True, color="FFFFFF")
        cel.fill, cel.border = azul, borda
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = larg
        for r in range(2, LINHAS_FORMATADAS + 2):
            ws.cell(r, c).number_format = fmt
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # validações de preenchimento (o Excel avisa na hora, antes de importar)
    letra = {k: get_column_letter(c) for c, (k, *_r) in enumerate(COLUNAS, 1)}

    def _dv(chave: str, **kw):
        dv = DataValidation(allow_blank=True, showErrorMessage=True, errorStyle="stop", **kw)
        ws.add_data_validation(dv)
        dv.add(f"{letra[chave]}2:{letra[chave]}{LINHAS_FORMATADAS + 1}")

    _dv("lp", type="list", formula1='"101,102,103"',
        errorTitle="Lançamento Padrão", error="Escolha 101, 102 ou 103.")
    _dv("conta_deb", type="textLength", operator="equal", formula1="8",
        errorTitle="Conta Débito", error="A conta deve ter 8 dígitos.")
    _dv("conta_cred", type="textLength", operator="equal", formula1="8",
        errorTitle="Conta Crédito", error="A conta deve ter 8 dígitos.")
    _dv("valor", type="decimal", operator="greaterThan", formula1="0",
        errorTitle="Valor", error="Informe um valor maior que zero.")
    _dv("filial_deb", type="textLength", operator="equal", formula1="7",
        errorTitle="Filial Débito", error="A filial deve ter 7 caracteres.")
    _dv("filial_cred", type="textLength", operator="equal", formula1="7",
        errorTitle="Filial Crédito", error="A filial deve ter 7 caracteres.")
    _dv("cc_deb", type="textLength", operator="lessThanOrEqual", formula1="7",
        errorTitle="Centro de Custo", error="No máximo 7 caracteres.")
    _dv("cc_cred", type="textLength", operator="lessThanOrEqual", formula1="7",
        errorTitle="Centro de Custo", error="No máximo 7 caracteres.")
    _dv("item_deb", type="textLength", operator="lessThanOrEqual", formula1="13",
        errorTitle="Item Contábil", error="No máximo 13 caracteres.")
    _dv("item_cred", type="textLength", operator="lessThanOrEqual", formula1="13",
        errorTitle="Item Contábil", error="No máximo 13 caracteres.")
    _dv("historico", type="textLength", operator="lessThanOrEqual", formula1="100",
        errorTitle="Histórico", error="No máximo 100 caracteres.")

    ajuda = wb.create_sheet(ABA_AJUDA)
    ajuda.column_dimensions["A"].width = 26
    for col in "BCDEFGHIJKL":
        ajuda.column_dimensions[col].width = 18
    ajuda["A1"] = "Como preencher"
    ajuda["A1"].font = Font(name="Segoe UI", bold=True, size=14)
    regras = [
        "Preencha a aba “Lançamentos”: uma linha por lançamento, a partir da linha 2. Não mude a ordem nem os nomes das colunas.",
        "Lançamento Padrão: obrigatório, escolha 101, 102 ou 103 na lista (não é preenchido automaticamente).",
        "Data: DD/MM/AAAA (pode ser data do Excel ou texto).",
        "Contas (8 dígitos, coluna em texto para preservar zeros à esquerda): LP 101 = Conta Débito e Conta Crédito "
        "obrigatórias | LP 102 = Conta Crédito VAZIA (só débito) | LP 103 = Conta Débito VAZIA (só crédito).",
        "Valor: maior que zero, até 99.999.999,99. Pode usar vírgula ou ponto decimal.",
        "Filiais (7 caracteres, ex.: 01AM001) seguem a mesma regra das contas: LP 101 = Filial Débito e Filial Crédito "
        "obrigatórias | LP 102 = Filial Crédito VAZIA | LP 103 = Filial Débito VAZIA.",
        "Centro de Custo: opcional, até 7 caracteres.",
        "Item Contábil (Cli/Forn): opcional, até 13 caracteres; segue o lado vazio dos LPs: "
        "LP 102 = Item Crédito VAZIO | LP 103 = Item Débito VAZIO.",
        "Histórico: obrigatório, até 100 caracteres (o app converte para MAIÚSCULAS e remove acentos).",
        "Linhas totalmente em branco são ignoradas. Linhas com erro não entram: o app lista o motivo de cada uma.",
    ]
    for i, texto in enumerate(regras, 3):
        ajuda.cell(i, 1, f"{i - 2}. {texto}").alignment = Alignment(wrap_text=False)

    ajuda.cell(13, 1, "Exemplo (não é importado — está só nesta aba):").font = Font(name="Segoe UI", bold=True)
    for c, (_k, titulo, _l, _f) in enumerate(COLUNAS, 1):
        cel = ajuda.cell(14, c, titulo)
        cel.font = Font(name="Segoe UI", bold=True, color="FFFFFF")
        cel.fill, cel.border = azul, borda
        cel.alignment = Alignment(horizontal="center", wrap_text=True)
    exemplos = [
        ["101", "30/09/2026", "11020001", "21010002", "1.234,56", "01AM001", "01AM001", "CC001", "CC002", "ITEM000000001", "ITEM000000002",
         "101: DEBITO E CREDITO PREENCHIDOS"],
        ["102", "30/09/2026", "11020001", "", "500,00", "01AM001", "", "CC001", "", "ITEM000000001", "",
         "102: SO DEBITO (CONTA, FILIAL E ITEM CREDITO VAZIOS)"],
        ["103", "30/09/2026", "", "21010002", "750,00", "", "01AM001", "", "CC002", "", "ITEM000000002",
         "103: SO CREDITO (CONTA, FILIAL E ITEM DEBITO VAZIOS)"],
    ]
    for i, exemplo in enumerate(exemplos, 15):
        for c, v in enumerate(exemplo, 1):
            ajuda.cell(i, c, v).border = borda

    wb.save(caminho)


# ───────────────────────────── leitura ─────────────────────────────
def _texto(v) -> str:
    return "" if v is None else str(v).strip()


def _celula(chave: str, v) -> str:
    """Converte o valor da célula para o texto que o validar() espera."""
    if v is None:
        return ""
    if chave == "data":
        if isinstance(v, (datetime, date)):
            return v.strftime("%d/%m/%Y")
        t = _texto(v)
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", t)            # "2026-09-30" -> "30/09/2026"
        return f"{m[3]}/{m[2]}/{m[1]}" if m else t
    if chave == "valor":
        if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
            return format(Decimal(str(v)).quantize(Decimal("0.01")), "f")   # 21.217 -> "21.22", nunca "milhar"
        return _texto(v)
    if isinstance(v, bool):
        return _texto(v)
    if isinstance(v, (int, float)):
        # Excel guardou como número: o zero à esquerda se perdeu (ex.: 01020001 -> 1020001)
        n = str(int(v)) if float(v).is_integer() else str(v)
        tam = {"conta_deb": 8, "conta_cred": 8, "filial_deb": 7, "filial_cred": 7}.get(chave)
        return n.zfill(tam) if tam and n.isdigit() else n
    return _texto(v)


def _achar_cabecalho(linhas: list[tuple]) -> tuple[int, dict[str, int]] | None:
    for i, row in enumerate(linhas[:15]):
        mapa = {}
        for j, v in enumerate(row):
            chave = _APELIDOS.get(_norm(v))
            if chave and chave not in mapa:
                mapa[chave] = j
        if len(mapa) >= 4:
            return i, mapa
    return None


def ler_planilha(caminho: str) -> tuple[list[dict], list[str]]:
    """
    Lê o modelo preenchido. Retorna (lançamentos válidos já normalizados, mensagens de erro).
    Levanta ValueError se não achar o cabeçalho ou faltar coluna obrigatória.
    """
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb[ABA] if ABA in wb.sheetnames else wb[wb.sheetnames[0]]
        linhas = [tuple(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    achado = _achar_cabecalho(linhas)
    if achado is None:
        raise ValueError("Não encontrei o cabeçalho da planilha. Use o modelo do app "
                         "(colunas Data, Conta Débito, Conta Crédito, Valor…).")
    i_cab, mapa = achado
    faltam = [t for k, t, _l, _f in COLUNAS if k in OBRIGATORIAS and k not in mapa]
    if faltam:
        raise ValueError("Coluna(s) obrigatória(s) não encontrada(s): " + ", ".join(faltam)
                         + ".\nUse o modelo atual (botão “Baixar modelo Excel”).")

    validos, erros = [], []
    for n, row in enumerate(linhas[i_cab + 1:], start=i_cab + 2):          # n = número da linha no Excel
        if all(_texto(v) == "" for v in row):
            continue
        bruto = {k: _celula(k, row[j] if j < len(row) else None) for k, j in mapa.items()}
        try:
            validos.append(validar(bruto))
        except ValueError as e:
            erros.append(f"Linha {n}: {e}")
        except Exception as e:                                              # data absurda etc.
            erros.append(f"Linha {n}: {e}")
    return validos, erros
