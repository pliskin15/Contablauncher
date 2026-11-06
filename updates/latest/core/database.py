"""
database.py
Camada de acesso a dados do Cofre — Supabase (compartilhado entre os PCs).

Mantém exatamente as mesmas funções públicas da versão SQLite
(core/database_local.py), então brinks/ e prossegur/ não mudam.

Tabelas: depositos_brinks, creditos_brinks, depositos_prossegur,
creditos_prossegur, conciliacao_ajustes — criadas por tools/supabase_cofre.sql.

Observações:
  • Datas continuam em texto DD/MM/YYYY (colunas com collate "C" no Postgres,
    para a comparação >= / <= se comportar como no SQLite).
  • Agregações (SUM/COUNT/GROUP BY) e ordenações por CAST(loja) são feitas
    aqui em Python, depois de buscar as linhas filtradas.
"""

import math
from datetime import datetime

from core.supabase_config import supabase

_PAGINA = 1000     # limite de linhas por requisição do PostgREST
_LOTE   = 500      # linhas por upsert

# ── Colunas por tabela (sem o id) ────────────────────────────────────────────
COLS_DEP_BRINKS = [
    "data_deposito", "data_inclusao", "nr_serial", "loja", "valor",
    "depositante", "nr_envelope", "sequencia", "identificador",
    "sigla_filial", "razao_social", "arquivo_origem", "importado_em",
]
COLS_CRED = [
    "data", "loja", "historico", "debito", "sequencia", "lote", "voucher",
    "doc_nro", "centro_custo", "conta_partida", "arquivo_origem", "importado_em",
]
COLS_DEP_PROSSEGUR = [
    "data_deposito", "data_hora", "nome_cofre", "loja", "valor", "tipo",
    "depositante", "cliente", "moeda", "arquivo_origem", "importado_em",
]
COLS_AJUSTES = [
    "transportadora", "loja", "data_deposito", "total_deposito", "total_credito",
    "memo_remessa", "dif_dep_memo", "dif_dep_cred", "status", "observacao",
    "data_credito_esperada", "editado_em",
]

CONFLITO_DEP_BRINKS    = "data_inclusao,nr_serial,identificador"
CONFLITO_CRED          = "data,loja,historico,sequencia,lote,voucher"
CONFLITO_DEP_PROSSEGUR = "data_hora,nome_cofre,valor"
CONFLITO_AJUSTES       = "transportadora,loja,data_deposito"


# ──────────────────────────────────────────────────────────────────────────── #
#  Utilitários                                                                  #
# ──────────────────────────────────────────────────────────────────────────── #

_verificado = False


def inicializar():
    """Confere (uma vez) que as tabelas existem no Supabase."""
    global _verificado
    if _verificado:
        return
    try:
        supabase.table("depositos_brinks").select("id").limit(1).execute()
    except Exception as e:
        raise RuntimeError(
            "Não foi possível acessar as tabelas do Cofre no Supabase.\n\n"
            "Se for a primeira vez, rode o script tools/supabase_cofre.sql no "
            "SQL Editor do Supabase. Se já rodou, verifique a conexão com a "
            f"internet e o arquivo .env.\n\nDetalhe: {e}"
        ) from e
    _verificado = True


def _agora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _limpar(v):
    """Converte tipos numpy/pandas para algo serializável em JSON; NaN -> None."""
    if v is None:
        return None
    if hasattr(v, "item"):          # numpy scalar
        v = v.item()
    if isinstance(v, float) and v != v:
        return None
    if hasattr(v, "isoformat") and not isinstance(v, str):
        return str(v)
    return v


def _int(s) -> int:
    """Equivalente ao CAST(x AS INTEGER) do SQLite (não numérico -> 0)."""
    try:
        return int(str(s).strip())
    except (TypeError, ValueError):
        return 0


def _normalizar(registros: list[dict], colunas: list[str]) -> list[dict]:
    return [{c: _limpar(r.get(c)) for c in colunas} for r in registros]


def _upsert_ignorar(tabela: str, colunas: list[str], conflito: str,
                    registros: list[dict]) -> tuple[int, int]:
    """INSERT ... ON CONFLICT DO NOTHING em lotes. Retorna (inseridos, ignorados)."""
    linhas = _normalizar(registros, colunas)
    inseridos = 0
    for i in range(0, len(linhas), _LOTE):
        lote = linhas[i:i + _LOTE]
        resp = (supabase.table(tabela)
                .upsert(lote, on_conflict=conflito, ignore_duplicates=True)
                .execute())
        inseridos += len(resp.data or [])
    return inseridos, len(linhas) - inseridos


def _buscar(tabela: str, colunas: str, col_data: str, data_inicio, data_fim,
            loja, extra=None) -> list[dict]:
    """SELECT paginado com os filtros de data/loja (e filtros extras eq)."""
    linhas, offset = [], 0
    while True:
        q = supabase.table(tabela).select(colunas)
        if data_inicio:
            q = q.gte(col_data, data_inicio)
        if data_fim:
            q = q.lte(col_data, data_fim)
        if loja:
            q = q.eq("loja", loja)
        for col, val in (extra or {}).items():
            q = q.eq(col, val)
        pagina = q.order("id").range(offset, offset + _PAGINA - 1).execute().data or []
        linhas.extend(pagina)
        if len(pagina) < _PAGINA:
            return linhas
        offset += _PAGINA


def _agrupar(linhas: list[dict], chave: list[str], soma: str) -> list[dict]:
    grupos: dict[tuple, dict] = {}
    for r in linhas:
        k = tuple(r[c] for c in chave)
        g = grupos.setdefault(k, {"valores": [], "qtd": 0})
        g["valores"].append(float(r[soma] or 0))
        g["qtd"] += 1
    # fsum evita ruído de ponto flutuante (3679.0699999999997 -> 3679.07)
    return [dict(zip(chave, k), total=round(math.fsum(g["valores"]), 2), qtd=g["qtd"])
            for k, g in grupos.items()]


# ──────────────────────────────────────────────────────────────────────────── #
#  Depósitos Brinks                                                             #
# ──────────────────────────────────────────────────────────────────────────── #

def insert_depositos_brinks(registros: list[dict], arquivo_origem: str) -> tuple[int, int]:
    inicializar()
    agora = _agora()
    for r in registros:
        r["arquivo_origem"] = arquivo_origem
        r["importado_em"]   = agora
    return _upsert_ignorar("depositos_brinks", COLS_DEP_BRINKS,
                           CONFLITO_DEP_BRINKS, registros)


def get_depositos_brinks(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    inicializar()
    rows = _buscar("depositos_brinks", "*", "data_deposito", data_inicio, data_fim, loja)
    rows.sort(key=lambda r: (r["data_deposito"], r["loja"], r["data_inclusao"]))
    return rows


def salvar_ajuste(transportadora: str, loja: str, data_deposito: str, dados: dict):
    """Insere ou atualiza um ajuste manual de conciliação."""
    inicializar()
    linha = {
        "transportadora":        transportadora,
        "loja":                  loja,
        "data_deposito":         data_deposito,
        "total_deposito":        dados.get("total_deposito"),
        "total_credito":         dados.get("total_credito"),
        "memo_remessa":          dados.get("memo_remessa"),
        "dif_dep_memo":          dados.get("dif_dep_memo"),
        "dif_dep_cred":          dados.get("dif_dep_cred"),
        "status":                dados.get("status"),
        "observacao":            dados.get("observacao", ""),
        "data_credito_esperada": dados.get("data_credito_esperada"),
        "editado_em":            _agora(),
    }
    (supabase.table("conciliacao_ajustes")
     .upsert(_normalizar([linha], COLS_AJUSTES), on_conflict=CONFLITO_AJUSTES)
     .execute())


def get_ajustes(transportadora: str) -> dict[tuple[str, str], dict]:
    """Retorna {(loja, data_deposito): dados} para a transportadora."""
    inicializar()
    cols = ("loja, data_deposito, total_deposito, total_credito, memo_remessa, "
            "dif_dep_memo, dif_dep_cred, status, observacao, data_credito_esperada")
    rows = _buscar("conciliacao_ajustes", cols, "data_deposito", None, None, None,
                   extra={"transportadora": transportadora})
    return {(r["loja"], r["data_deposito"]): r for r in rows}


def get_resumo_depositos_brinks(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    inicializar()
    rows = _buscar("depositos_brinks", "data_deposito, loja, nr_serial, valor",
                   "data_deposito", data_inicio, data_fim, loja)
    res = _agrupar(rows, ["data_deposito", "loja", "nr_serial"], "valor")
    res = [{"data_corte": r["data_deposito"], "loja": r["loja"],
            "nr_serial": r["nr_serial"], "total": r["total"], "qtd": r["qtd"]}
           for r in res]
    res.sort(key=lambda r: (r["data_corte"], _int(r["loja"]), r["nr_serial"]))
    return res


def get_lancamentos_dia_brinks(data_deposito: str, nr_serial: str) -> list[dict]:
    inicializar()
    linhas, offset = [], 0
    while True:
        pagina = (supabase.table("depositos_brinks")
                  .select("data_inclusao, valor, depositante")
                  .like("data_inclusao", f"{data_deposito}%")
                  .eq("nr_serial", nr_serial)
                  .order("id").range(offset, offset + _PAGINA - 1)
                  .execute().data or [])
        linhas.extend(pagina)
        if len(pagina) < _PAGINA:
            break
        offset += _PAGINA
    linhas.sort(key=lambda r: r["data_inclusao"])
    return linhas


# ──────────────────────────────────────────────────────────────────────────── #
#  Créditos Brinks                                                              #
# ──────────────────────────────────────────────────────────────────────────── #

def insert_creditos_brinks(registros: list[dict]) -> tuple[int, int]:
    inicializar()
    agora = _agora()
    for r in registros:
        r["importado_em"] = agora
    return _upsert_ignorar("creditos_brinks", COLS_CRED, CONFLITO_CRED, registros)


def get_creditos_brinks(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    inicializar()
    cols = ("data, loja, debito, historico, sequencia, lote, voucher, doc_nro, "
            "centro_custo, conta_partida, arquivo_origem")
    rows = _buscar("creditos_brinks", cols, "data", data_inicio, data_fim, loja)
    rows.sort(key=lambda r: (r["data"], _int(r["loja"])))
    return rows


# ──────────────────────────────────────────────────────────────────────────── #
#  Depósitos Prossegur                                                          #
# ──────────────────────────────────────────────────────────────────────────── #

def insert_depositos_prossegur(registros: list[dict], arquivo_origem: str) -> tuple[int, int]:
    inicializar()
    agora = _agora()
    for r in registros:
        r["arquivo_origem"] = arquivo_origem
        r["importado_em"]   = agora
    return _upsert_ignorar("depositos_prossegur", COLS_DEP_PROSSEGUR,
                           CONFLITO_DEP_PROSSEGUR, registros)


def get_resumo_depositos_prossegur(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    """Agrupa por (data_deposito, loja) — 1 linha por loja/dia com soma."""
    inicializar()
    rows = _buscar("depositos_prossegur", "data_deposito, loja, valor",
                   "data_deposito", data_inicio, data_fim, loja)
    res = _agrupar(rows, ["data_deposito", "loja"], "valor")
    res = [{"data_corte": r["data_deposito"], "loja": r["loja"],
            "total": r["total"], "qtd": r["qtd"]} for r in res]
    res.sort(key=lambda r: (r["data_corte"], _int(r["loja"])))
    return res


def get_resumo_depositos_prossegur_com_hora(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    """
    Retorna cada lançamento individual com data_hora, para que o motor
    de conciliação possa aplicar o corte de 16h (lançamentos após 16h
    são contabilizados no dia seguinte para comparação com créditos).
    """
    inicializar()
    rows = _buscar("depositos_prossegur", "data_deposito, data_hora, loja, valor",
                   "data_deposito", data_inicio, data_fim, loja)
    rows.sort(key=lambda r: (r["data_hora"], _int(r["loja"])))
    return rows


def get_lancamentos_dia_prossegur(data_deposito: str, loja: str) -> list[dict]:
    """Todos os lançamentos de uma loja em um dia (todos os cofres)."""
    inicializar()
    rows = _buscar("depositos_prossegur",
                   "data_hora, tipo, valor, nome_cofre, depositante, cliente",
                   "data_deposito", None, None, None,
                   extra={"data_deposito": data_deposito, "loja": loja})
    rows.sort(key=lambda r: r["data_hora"])
    return rows


def remapear_lojas_prossegur(mapa: dict) -> int:
    """
    Preenche `loja` nos depósitos sem loja, a partir de {nome_cofre: [lojas]}.
    Só remapeia nomes com exatamente uma loja possível. Retorna o nº de linhas.
    """
    inicializar()
    total = 0
    for nome, lojas in mapa.items():
        if len(lojas) != 1:
            continue
        resp = (supabase.table("depositos_prossegur")
                .update({"loja": str(lojas[0])})
                .eq("nome_cofre", nome).is_("loja", "null")
                .execute())
        total += len(resp.data or [])
    return total


# ──────────────────────────────────────────────────────────────────────────── #
#  Créditos Prossegur                                                           #
# ──────────────────────────────────────────────────────────────────────────── #

def insert_creditos_prossegur(registros: list[dict]) -> tuple[int, int]:
    inicializar()
    agora = _agora()
    for r in registros:
        r["importado_em"] = agora
    return _upsert_ignorar("creditos_prossegur", COLS_CRED, CONFLITO_CRED, registros)


def get_creditos_prossegur(
    data_inicio: str | None = None,
    data_fim:    str | None = None,
    loja:        str | None = None,
) -> list[dict]:
    """Agrupa por (data, loja) somando o débito — 1 linha por loja/dia."""
    inicializar()
    rows = _buscar("creditos_prossegur", "data, loja, debito", "data",
                   data_inicio, data_fim, loja)
    res = _agrupar(rows, ["data", "loja"], "debito")
    res = [{"data": r["data"], "loja": r["loja"],
            "debito": r["total"], "qtd": r["qtd"]} for r in res]
    res.sort(key=lambda r: (r["data"], _int(r["loja"])))
    return res
