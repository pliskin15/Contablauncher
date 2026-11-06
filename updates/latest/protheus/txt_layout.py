"""
Gerador de TXT para importação de lançamentos contábeis em lote no Protheus.

Layout (posições fixas, TAB entre campos, linhas em CRLF, 204 caracteres):
  001-003 Lançamento padrão (101, 102 ou 103 — escolhido pelo usuário) | 005 "X"
  007-014 Conta Déb | 016-023 Conta Créd
  025-034 Valor (centavos, zeros à esquerda) | 036-135 Histórico
  137-144 Data DDMMAAAA | 146-152 Filial Déb | 154-160 Filial Créd
  162-168 C.Custo Déb | 170-176 C.Custo Créd
  178-190 Item Contábil Déb | 192-204 Item Contábil Créd  (Cli/Forn no Protheus)
"""

import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

ENCODING = "cp1252"

# Lançamentos padrão aceitos (posições 001-003). Não há valor automático: o usuário escolhe.
LPS = ("101", "102", "103")
# Conta que obrigatoriamente fica VAZIA em cada lançamento padrão (101: as duas são obrigatórias)
CONTA_VAZIA = {"102": "conta_cred", "103": "conta_deb"}
# Idem para a filial: fica vazia do mesmo lado da conta vazia
FILIAL_VAZIA = {"102": "filial_cred", "103": "filial_deb"}
# Item contábil segue o mesmo lado vazio da conta (102: crédito; 103: débito)
ITEM_VAZIO = {"102": "item_cred", "103": "item_deb"}

# (chave, rótulo, tamanho no TXT)
CAMPOS = [
    ("lp", "Lançamento Padrão", 3),
    ("data", "Data (DD/MM/AAAA)", 8),
    ("conta_deb", "Conta Débito", 8),
    ("conta_cred", "Conta Crédito", 8),
    ("valor", "Valor (R$)", 10),
    ("filial_deb", "Filial Débito", 7),
    ("filial_cred", "Filial Crédito", 7),
    ("cc_deb", "C. Custo Débito", 7),
    ("cc_cred", "C. Custo Crédito", 7),
    ("item_deb", "Item Contábil Débito", 13),
    ("item_cred", "Item Contábil Crédito", 13),
    ("historico", "Histórico", 100),
]
TAM = {k: t for k, _, t in CAMPOS}
# Campos limpos após adicionar (os demais ficam para agilizar a digitação)
LIMPAR_APOS_ADD = ("conta_deb", "valor", "historico", "cc_deb", "item_deb")


# ----------------------------------------------------------------- utilidades
def sem_acento(txt: str) -> str:
    nfkd = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper()


def ultimo_dia_mes_anterior() -> str:
    d = date.today().replace(day=1) - timedelta(days=1)
    return d.strftime("%d/%m/%Y")


def parse_valor(txt: str) -> Decimal:
    s = txt.strip().replace("R$", "").replace(" ", "")
    if not s:
        raise ValueError("Informe o valor.")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") == 1 and len(s.split(".")[1]) == 3:
        s = s.replace(".", "")  # 21.217 -> milhar
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        v = Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        raise ValueError(f"Valor inválido: {txt}")
    if v <= 0:
        raise ValueError("O valor deve ser maior que zero.")
    if v > Decimal("99999999.99"):
        raise ValueError("Valor excede 10 dígitos (máx. 99.999.999,99).")
    return v


def fmt_brl(v: Decimal) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def validar(dados: dict) -> dict:
    """Valida e normaliza um lançamento. Retorna dict limpo ou levanta ValueError."""
    d = {k: str(dados.get(k, "")).strip() for k in TAM}

    if d["lp"] not in LPS:
        raise ValueError("Escolha o lançamento padrão: 101, 102 ou 103.")

    try:
        datetime.strptime(d["data"], "%d/%m/%Y")
    except ValueError:
        raise ValueError("Data inválida. Use DD/MM/AAAA.")

    for k in ("conta_deb", "conta_cred"):
        rotulo = dict((c, r) for c, r, _ in CAMPOS)[k]
        if CONTA_VAZIA.get(d["lp"]) == k:
            if d[k]:
                raise ValueError(f"No lançamento padrão {d['lp']}, {rotulo} deve ficar vazia.")
            continue
        if not d[k].isdigit() or len(d[k]) != 8:
            raise ValueError(f"{rotulo} deve ter 8 dígitos"
                             f"{' (obrigatória no lançamento padrão 101)' if d['lp'] == '101' else ''}.")

    d["valor"] = str(parse_valor(d["valor"]))

    for k in ("filial_deb", "filial_cred"):
        rotulo = dict((c, r) for c, r, _ in CAMPOS)[k]
        d[k] = d[k].upper()
        if FILIAL_VAZIA.get(d["lp"]) == k:
            if d[k]:
                raise ValueError(f"No lançamento padrão {d['lp']}, {rotulo} deve ficar vazia.")
            continue
        if len(d[k]) != 7:
            raise ValueError(f"{rotulo} deve ter 7 caracteres"
                             f"{' (obrigatória no lançamento padrão 101)' if d['lp'] == '101' else ''}.")

    for k in ("cc_deb", "cc_cred"):
        d[k] = d[k].upper()
        if len(d[k]) > TAM[k]:
            raise ValueError(f"Centro de custo tem no máximo {TAM[k]} caracteres.")

    for k in ("item_deb", "item_cred"):
        rotulo = dict((c, r) for c, r, _ in CAMPOS)[k]
        d[k] = d[k].upper()
        if ITEM_VAZIO.get(d["lp"]) == k:
            if d[k]:
                raise ValueError(f"No lançamento padrão {d['lp']}, {rotulo} deve ficar vazio.")
            continue
        if len(d[k]) > TAM[k]:
            raise ValueError(f"{rotulo} tem no máximo {TAM[k]} caracteres.")

    d["historico"] = sem_acento(d["historico"])
    if not d["historico"]:
        raise ValueError("Informe o histórico.")
    if len(d["historico"]) > 100:
        raise ValueError(f"Histórico com {len(d['historico'])} caracteres (máx. 100).")
    return d


def montar_linha(d: dict) -> str:
    centavos = int(Decimal(d["valor"]) * 100)
    campos = [
        d["lp"],
        "X",
        d["conta_deb"].ljust(8),
        d["conta_cred"].ljust(8),
        str(centavos).zfill(10),
        d["historico"].ljust(100)[:100],
        d["data"].replace("/", ""),
        d["filial_deb"].ljust(7),
        d["filial_cred"].ljust(7),
        d["cc_deb"].ljust(7)[:7],
        d["cc_cred"].ljust(7)[:7],
        d["item_deb"].ljust(13)[:13],
        d["item_cred"].ljust(13)[:13],
    ]
    return "\t".join(campos)


def ler_linha(linha: str) -> dict:
    """Lê uma linha do TXT pelas posições do lançamento padrão (base 1)."""
    linha = linha.rstrip("\r\n").ljust(204)

    def pos(ini, tam):
        return linha[ini - 1:ini - 1 + tam].strip()

    dt = pos(137, 8)
    return {
        "lp": pos(1, 3),
        "conta_deb": pos(7, 8),
        "conta_cred": pos(16, 8),
        "valor": str(Decimal(int(pos(25, 10) or 0)) / 100),
        "historico": pos(36, 100),
        "data": f"{dt[:2]}/{dt[2:4]}/{dt[4:]}",
        "filial_deb": pos(146, 7),
        "filial_cred": pos(154, 7),
        "cc_deb": pos(162, 7),
        "cc_cred": pos(170, 7),
        "item_deb": pos(178, 13),
        "item_cred": pos(192, 13),
    }
