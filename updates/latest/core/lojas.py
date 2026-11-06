"""
lojas.py
Módulo de Lojas — Conciliação Vendas × Contabilidade.

Estrutura:
  • Seleção de loja via Combobox
  • Importação de 2 bases Excel (Vendas e Contabilidade)
  • Tabela de conciliação na tela principal
  • Painel inferior com filtro Mês/Ano e cards-resumo por loja
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import re
import unicodedata
from collections import defaultdict

from core import theme, versao
from core.supabase_config import supabase

# ── Persistência local (apenas conciliações) ──────────────────────────────── #
_DATA_FILE = os.path.join(os.path.dirname(__file__), "lojas_data.json")


def _load_data() -> dict:
    """Carrega dados salvos em disco (apenas conciliações)."""
    if os.path.exists(_DATA_FILE):
        try:
            with open(_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"conciliacoes": {}}


def _save_data(data: dict):
    """Persiste dados em disco (apenas conciliações)."""
    with open(_DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── Supabase: busca lojas ──────────────────────────────────────────────────── #
def _fetch_lojas_supabase() -> list:
    """
    Retorna lista de dicts com as chaves: loja, cnpj, regional.
    Ex: [{"loja": "Loja Centro", "cnpj": "00.000.000/0001-00", "regional": "AM"}, ...]
    Retorna [] em caso de erro.
    """
    try:
        resp = supabase.table("lojas").select("loja, cnpj, regional").execute()
        return resp.data or []
    except Exception as e:
        messagebox.showerror("Erro Supabase",
                             f"Não foi possível carregar as lojas:\n{e}")
        return []


# ── Utilitários ────────────────────────────────────────────────────────────── #
def _parse_value(v) -> float:
    """Converte string monetária → float."""
    try:
        s = str(v).strip().replace("R$", "").replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        f = float(s)
        return 0.0 if f != f else f          # NaN → 0
    except Exception:
        return 0.0



# ══════════════════════════════════════════════════════════════════════════════
#   LEITURA DAS BASES  (Vendas Goodcard  ×  Razão modelo I da conta Goodcard)
# ══════════════════════════════════════════════════════════════════════════════
TOL_VALOR   = 0.015   # tolerância (R$) ao comparar valores
JANELA_DIAS = 7       # janela de datas ao casar por valor (quando não há NF)

_STATUS_TXT = {
    "verde":    "✔ OK",
    "amarelo":  "⚠ Atenção",
    "vermelho": "✖ Não encontrado",
    "cinza":    "— Fora do período",
    "roxo":     "↪ Lançada na 2157",
}


def _norm(s) -> str:
    """Maiúsculas, sem acento, espaços colapsados (para comparar cabeçalhos)."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def _find_col(df, *candidates, exact=False):
    """Localiza coluna por nome normalizado (exato; se exact=False, também 'contém')."""
    cols = {_norm(c): c for c in df.columns}
    for cand in candidates:
        n = _norm(cand)
        if n in cols:
            return cols[n]
    if not exact:
        for cand in candidates:
            n = _norm(cand)
            for k, c in cols.items():
                if n in k:
                    return c
    return None


def _clean_doc(v) -> str:
    s = str(v).strip()
    if s.lower() in ("", "nan", "none", "nat"):
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


def _brl(v) -> str:
    """1234.5 → '1.234,50'."""
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def _engine(path: str) -> str:
    return "xlrd" if path.lower().endswith(".xls") else "openpyxl"


def _meta_valores(raw, label: str, n: int = 1) -> list:
    """Acha `label` numa célula e devolve as n próximas células preenchidas da linha."""
    for i in range(len(raw)):
        vals = raw.iloc[i].tolist()
        for j, c in enumerate(vals):
            if _norm(c) == label:
                resto = [str(x).strip() for x in vals[j + 1:]
                         if str(x).strip().lower() not in ("", "nan", "none")]
                return resto[:n]
    return []


def _ler_razao(path: str):
    """
    Lê o Razão Modelo I (.xls/.xlsx) — cabeçalho detectado automaticamente
    (linha em que a 1ª coluna é 'DATA' e existe 'CONTA PARTIDA').

    Retorna (df, meta):
      df   — linhas de lançamento + coluna auxiliar `_filial`
      meta — {filial: {"ini": Timestamp, "fim": Timestamp, "conta": str}}
    """
    import pandas as pd
    xl = pd.ExcelFile(path, engine=_engine(path))
    frames, meta = [], {}
    for sh in xl.sheet_names:
        raw = xl.parse(sh, header=None, dtype=str)
        hdr = next((i for i in range(min(len(raw), 30))
                    if _norm(raw.iat[i, 0]) == "DATA"
                    and any(_norm(x) == "CONTA PARTIDA" for x in raw.iloc[i].tolist())),
                   None)
        if hdr is None:
            continue
        top = raw.iloc[:hdr]

        fil = _meta_valores(top, "FILIAL")
        m = re.search(r"\d+", fil[0]) if fil else None
        if not m:
            m = re.search(r"\d+", sh)
        filial = m.group(0) if m else sh

        per = _meta_valores(top, "PERIODO")
        datas = re.findall(r"\d{2}/\d{2}/\d{4}", per[0]) if per else []
        ini = pd.to_datetime(datas[0], format="%d/%m/%Y") if len(datas) >= 2 else None
        fim = pd.to_datetime(datas[1], format="%d/%m/%Y") if len(datas) >= 2 else None
        conta = " · ".join(_meta_valores(top, "CONTA", 2))

        d = raw.iloc[hdr + 1:].copy()
        d.columns = [str(c).strip() for c in raw.iloc[hdr]]
        d = d[d["DATA"].fillna("").str.match(r"^\d{2}/\d{2}/\d{4}")].copy()
        d["_filial"] = filial
        frames.append(d)
        meta[filial] = {"ini": ini, "fim": fim, "conta": conta}

    if not frames:
        raise ValueError("Não encontrei o cabeçalho do Razão (colunas DATA … CONTA PARTIDA) "
                         "em nenhuma aba do arquivo.")
    return pd.concat(frames, ignore_index=True), meta


def _resumo_razao(df, meta) -> str:
    partes = []
    for fil, m in meta.items():
        per = (f"{m['ini']:%d/%m/%Y}–{m['fim']:%d/%m/%Y}" if m.get("ini") is not None else "período ?")
        partes.append(f"Filial {fil} · {per}")
    return f"{len(df)} lançamentos · " + " | ".join(partes)


def _parse_dt(x: str, us: bool):
    """Interpreta a data/hora de uma venda."""
    import pandas as pd
    x = x.strip()
    if not x:
        return pd.NaT
    if re.match(r"^\d{4}-\d{2}-\d{2}", x):
        d = pd.to_datetime(x[:19], errors="coerce")
        if pd.isna(d):
            return pd.NaT
        # Excel (pt-BR) leu 'MM/DD/AAAA' como 'DD/MM/AAAA' e inverteu dia e mês
        if us and d.day <= 12:
            d = d.replace(month=d.day, day=d.month)
        return d
    fmts = (["%m/%d/%Y %H:%M", "%m/%d/%Y"] if us else ["%d/%m/%Y %H:%M", "%d/%m/%Y"])
    for f in fmts:
        d = pd.to_datetime(x, format=f, errors="coerce")
        if not pd.isna(d):
            return d
    return pd.NaT


def _datas_vendas(serie):
    """
    Normaliza a coluna 'Data / Hora' da base de Vendas.
    A base mistura texto 'MM/DD/AAAA hh:mm' (dia > 12) com datas reais que o Excel
    gerou invertendo dia/mês (dia ≤ 12). Só inverte de volta quando há texto
    inequivocamente em formato americano e nenhum em formato brasileiro.
    """
    s = serie.fillna("").astype(str).str.strip()
    us_n = br_n = 0
    for x in s:
        m = re.match(r"^(\d{1,2})/(\d{1,2})/\d{4}", x)
        if m:
            if int(m.group(2)) > 12:
                us_n += 1
            if int(m.group(1)) > 12:
                br_n += 1
    us = us_n > 0 and br_n == 0
    return [_parse_dt(x, us) for x in s]


def _ler_vendas(path: str):
    """
    Lê a base de Vendas Goodcard (todas as abas 'Filial NN').
    Cabeçalho detectado pela linha cuja 1ª coluna é 'Estabelecimento'.
    Adiciona colunas auxiliares: `_filial`, `_data` (Timestamp).
    """
    import pandas as pd
    xl = pd.ExcelFile(path, engine=_engine(path))
    frames = []
    for sh in xl.sheet_names:
        raw = xl.parse(sh, header=None, dtype=str)
        hdr = next((i for i in range(min(len(raw), 30))
                    if _norm(raw.iat[i, 0]) == "ESTABELECIMENTO"), None)
        if hdr is None:
            continue
        d = raw.iloc[hdr + 1:].copy()
        d.columns = [str(c).strip() for c in raw.iloc[hdr]]
        c_cod = _find_col(d, "Cód. de Venda", exact=True)
        if c_cod is None:
            continue
        d = d[d[c_cod].fillna("").str.strip().str.fullmatch(r"\d+")].copy()
        m = re.search(r"\d+", sh)
        d["_filial"] = m.group(0) if m else sh
        frames.append(d)

    if not frames:
        raise ValueError("Não encontrei o cabeçalho da base de Vendas "
                         "(linha com 'Estabelecimento') em nenhuma aba.")
    df = pd.concat(frames, ignore_index=True)
    c_dat = _find_col(df, "Data / Hora", "Data")
    if c_dat is None:
        raise ValueError("Base de Vendas sem a coluna 'Data / Hora'.")
    df["_data"] = _datas_vendas(df[c_dat])
    return df


def _resumo_vendas(df) -> str:
    cont = df["_filial"].value_counts().sort_index()
    return f"{len(df)} vendas · " + ", ".join(f"Filial {f} ({n})" for f, n in cont.items())


# ══════════════════════════════════════════════════════════════════════════════
#   CONCILIAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
def _conciliar_filial(sales: list, debs: list, creds: list, ini, fim) -> None:
    """
    Preenche s["tag"], s["lines"], s["obs"], s["crit"] em cada venda.

      1) Período   — venda fora do período do razão → 'cinza'
      2) Por NF    — Nota Fiscal (Peça/Serviço) da venda = DOC. NRO. do razão;
                     soma das vendas da NF = soma dos débitos da NF → 'verde',
                     senão → 'amarelo'
      3) Valor+Data— sem NF casada: débito de valor igual, dentro de ±JANELA_DIAS,
                     1 lançamento para 1 venda (mais próximo primeiro) → 'amarelo'
      4) Resto     — 'vermelho'
      5) Devolução — venda casada cuja NF tem crédito (DEV.) no razão → 'amarelo'
    """
    import pandas as pd

    deb_by_doc = defaultdict(list)
    for l in debs:
        if l["doc"]:
            deb_by_doc[l["doc"]].append(l)
    cred_by_nf = defaultdict(list)
    for c in creds:
        if c["nf"]:
            cred_by_nf[c["nf"]].append(c)

    for s in sales:
        s.update(tag=None, lines=[], obs="", crit="—")

    # 1) período
    if ini is not None and fim is not None:
        for s in sales:
            if pd.isna(s["data"]):
                continue
            d0 = s["data"].normalize()
            if d0 < ini or d0 > fim:
                s["tag"] = "cinza"
                s["obs"] = (f"Venda de {s['data']:%d/%m/%Y} fora do período do razão "
                            f"({ini:%d/%m/%Y} a {fim:%d/%m/%Y})")

    # 2) por NF
    grupos = defaultdict(list)
    for s in sales:
        if s["tag"]:
            continue
        chave = next((n for n in s["nfs"] if n in deb_by_doc), None)
        if chave:
            grupos[chave].append(s)
    for chave, gs in grupos.items():
        linhas = deb_by_doc[chave]
        for l in linhas:
            l["used"] = True
        tot_l = round(sum(l["valor"] for l in linhas), 2)
        tot_s = round(sum(s["valor"] for s in gs), 2)
        ok = abs(tot_l - tot_s) < TOL_VALOR
        for s in gs:
            s["lines"], s["crit"] = linhas, "NF"
            if ok:
                s["tag"] = "verde"
                s["obs"] = f"NF {chave} — {len(linhas)} lançamento(s) no razão"
            else:
                s["tag"] = "amarelo"
                s["obs"] = f"NF {chave}: vendas R$ {_brl(tot_s)} × razão R$ {_brl(tot_l)}"

    # 3) valor + data (1 × 1)
    pend = [s for s in sales if not s["tag"]]
    idx = defaultdict(list)
    for l in debs:
        if not l["used"]:
            idx[f"{l['valor']:.2f}"].append(l)
    pares = []
    for si, s in enumerate(pend):
        if pd.isna(s["data"]):
            continue
        for l in idx.get(f"{s['valor']:.2f}", []):
            if pd.isna(l["data"]):
                continue
            dd = abs((l["data"] - s["data"].normalize()).days)
            if dd <= JANELA_DIAS:
                pares.append((dd, si, l["i"], l))
    pares.sort(key=lambda p: (p[0], p[1], p[2]))
    for dd, si, _, l in pares:
        s = pend[si]
        if s["tag"] or l["used"]:
            continue
        l["used"] = True
        s["tag"], s["lines"], s["crit"] = "amarelo", [l], "Valor+Data"
        nf_txt = f"NF {s['nfs'][0]} não está no razão; " if s["nfs"] else "Sem NF; "
        s["obs"] = f"{nf_txt}casado por valor e data (Δ {dd} dia(s))"

    # 4) não encontrado
    for s in pend:
        if s["tag"]:
            continue
        s["tag"] = "vermelho"
        mesmos = [l for l in debs if abs(l["valor"] - s["valor"]) < TOL_VALOR
                  and not pd.isna(l["data"])]
        if mesmos and not pd.isna(s["data"]):
            l = min(mesmos, key=lambda l: abs((l["data"] - s["data"].normalize()).days))
            usado = " (já usado em outra venda)" if l["used"] else ""
            s["obs"] = (f"Valor igual no razão em {l['data']:%d/%m/%Y}, "
                        f"fora da janela de {JANELA_DIAS} dias{usado}")
        else:
            s["obs"] = "Nenhum débito de mesmo valor no razão"

    # 5) devolução
    for s in sales:
        if s["tag"] in ("verde", "amarelo") and s["lines"]:
            docs = {l["doc"] for l in s["lines"]}
            cs = [c for d in docs for c in cred_by_nf.get(d, [])]
            if cs:
                tot = round(sum(c["valor"] for c in cs), 2)
                tipo = "total" if any("TOTAL" in c["hist"].upper() for c in cs) else "parcial"
                s["tag"] = "amarelo"
                s["obs"] = (s["obs"] + " | " if s["obs"] else "") + \
                           f"Devolução {tipo} no razão: R$ {_brl(tot)}"
                s["devol"] = [c["raw"] for c in cs]


def _colunas_razao(df_c) -> dict:
    need = {k: _find_col(df_c, k, exact=True)
            for k in ("DATA", "DOC. NRO.", "CONTA PARTIDA", "DEBITO", "CREDITO")}
    faltam = [k for k, v in need.items() if v is None]
    if faltam:
        raise ValueError(f"Razão: coluna(s) não encontrada(s): {', '.join(faltam)}.\n\n"
                         f"Colunas: {list(df_c.columns)}")
    return need


def _extrair_lancamentos(sub_c, need, c_hist):
    """Separa as linhas do razão em débitos e créditos."""
    import pandas as pd
    debs, creds = [], []
    for i, (_, r) in enumerate(sub_c.iterrows()):
        d   = pd.to_datetime(str(r[need["DATA"]]).strip(), format="%d/%m/%Y", errors="coerce")
        deb = _parse_value(r[need["DEBITO"]])
        cre = _parse_value(r[need["CREDITO"]])
        raw = {str(k): ("" if str(v) in ("nan", "None") else str(v))
               for k, v in r.items() if not str(k).startswith("_")}
        hist = raw.get(c_hist, "") if c_hist else ""
        if deb > 0:
            debs.append(dict(i=i, data=d, valor=deb, used=False, raw=raw,
                             doc=_clean_doc(r[need["DOC. NRO."]]),
                             partida=raw.get(need["CONTA PARTIDA"], "")))
        if cre > 0:
            mnf = re.search(r"NF\.?\s*NR\.?\s*(\d+)", hist.upper())
            creds.append(dict(data=d, valor=cre, hist=hist, raw=raw,
                              nf=mnf.group(1) if mnf else ""))
    return debs, creds


def conciliar_bases(df_v, df_c, meta_c: dict, df_c2=None):
    """
    Concilia a base de Vendas Goodcard com o Razão (conta Goodcard: venda = débito,
    devolução = crédito em conta de partida de devolução).

    df_c2 (opcional) — Razão da conta 2157 (outra conta de cartão). As vendas que
    não forem encontradas no razão principal são procuradas nela; se casarem,
    ficam 'roxo' (venda lançada na conta errada).
    Retorna (rows, aviso). Levanta ValueError se faltar coluna.
    """
    import pandas as pd

    # ── colunas das Vendas ──
    c_val = _find_col(df_v, "Valor Bruto", exact=True)
    if c_val is None:
        raise ValueError("Base de Vendas: coluna 'Valor Bruto' não encontrada.\n\n"
                         f"Colunas: {list(df_v.columns)}")
    c_os  = _find_col(df_v, "Ordem de Servico", exact=True)
    c_pl  = _find_col(df_v, "Placa", exact=True)
    c_cli = _find_col(df_v, "Cliente", exact=True)
    c_cod = _find_col(df_v, "Cód. de Venda", exact=True)
    c_nfs = [c for c in (_find_col(df_v, "Nota Fiscal Peça", exact=True),
                         _find_col(df_v, "Nota Fiscal Serviço", exact=True),
                         _find_col(df_v, "Nota Fiscal", exact=True)) if c]

    # ── colunas do Razão ──
    need = _colunas_razao(df_c)
    c_hist = _find_col(df_c, "HISTORICO", exact=True)
    need2 = _colunas_razao(df_c2) if df_c2 is not None else None
    c_hist2 = _find_col(df_c2, "HISTORICO", exact=True) if df_c2 is not None else None

    if not meta_c:
        meta_c = {f: {"ini": None, "fim": None} for f in df_c["_filial"].unique()}

    rows, avisos = [], []
    for fil, m in meta_c.items():
        sub_v = df_v[df_v["_filial"] == fil]
        sub_c = df_c[df_c["_filial"] == fil]
        if sub_v.empty:
            avisos.append(f"Filial {fil} do razão não existe na base de Vendas")
            continue

        sales = []
        for _, r in sub_v.iterrows():
            valor = _parse_value(r[c_val])
            if valor <= 0:
                continue
            nfs = [n for n in (_clean_doc(r[c]) for c in c_nfs) if n]
            os_ = _clean_doc(r[c_os]) if c_os else ""
            sales.append(dict(
                data=r["_data"], valor=valor, nfs=nfs,
                doc=nfs[0] if nfs else (f"OS {os_}" if os_ else "—"),
                ordem=os_,
                placa=_clean_doc(r[c_pl]) if c_pl else "",
                cliente=_clean_doc(r[c_cli]) if c_cli else "",
                cod=_clean_doc(r[c_cod]) if c_cod else "",
            ))

        debs, creds = _extrair_lancamentos(sub_c, need, c_hist)

        _conciliar_filial(sales, debs, creds, m.get("ini"), m.get("fim"))

        # 2ª passada: vendas não encontradas → procura no razão da 2157
        if df_c2 is not None:
            sub_c2 = df_c2[df_c2["_filial"] == fil]
            reds = [x for x in sales if x["tag"] == "vermelho"]
            if reds and not sub_c2.empty:
                debs2, _ = _extrair_lancamentos(sub_c2, need2, c_hist2)
                obs0 = {id(x): x["obs"] for x in reds}
                _conciliar_filial(reds, debs2, [], None, None)
                for x in reds:
                    if x["tag"] in ("verde", "amarelo"):
                        x["tag"] = "roxo"
                        x["crit"] = "2157 · " + x["crit"]
                        x["obs"] = ("Encontrada na conta 2157 (possível lançamento na conta "
                                    "errada) — " + x["obs"])
                        x["conta2"] = True
                    else:
                        x["tag"], x["obs"] = "vermelho", obs0[id(x)]

        def _ordem(s):
            return (pd.isna(s["data"]), s["data"] if not pd.isna(s["data"]) else pd.Timestamp(0), s["doc"])

        for s in sorted(sales, key=_ordem):
            linhas = [l["raw"] for l in s["lines"]] + s.get("devol", [])
            partidas = sorted({l["partida"] for l in s["lines"] if l["partida"]})
            if s.get("conta2"):
                partidas = ["2157 ← " + p for p in partidas] or ["2157"]
            ok_dt = not pd.isna(s["data"])
            rows.append({
                "filial":       fil,
                "data":         f"{s['data']:%d/%m/%Y}" if ok_dt else "—",
                "hora":         f"{s['data']:%H:%M}" if ok_dt else "",
                "doc":          s["doc"],
                "valor_venda":  s["valor"],
                "conta_contab": ", ".join(partidas) if partidas else "—",
                "status":       _STATUS_TXT[s["tag"]],
                "status_tag":   s["tag"],
                "obs":          s["obs"],
                "criterio":     s["crit"],
                "cliente":      s["cliente"],
                "placa":        s["placa"],
                "ordem":        s["ordem"],
                "cod_venda":    s["cod"],
                "linhas_contab": linhas,
                "linha_contab":  linhas[0] if linhas else {},
            })

    outras = {}
    for f in df_v["_filial"]:
        if f not in meta_c:
            outras[f] = outras.get(f, 0) + 1
    if outras:
        avisos.append("Vendas de outras filiais ignoradas: " +
                      ", ".join(f"{f} ({n})" for f, n in sorted(outras.items())))
    if not rows:
        raise ValueError("Nenhuma venda pôde ser conciliada.\n" + "\n".join(avisos))
    return rows, " | ".join(avisos)



# ══════════════════════════════════════════════════════════════════════════════
#   EXPORTAÇÃO PARA EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def _to_date(s):
    import pandas as pd
    d = pd.to_datetime(str(s).strip(), format="%d/%m/%Y", errors="coerce")
    return None if pd.isna(d) else d.to_pydatetime()


def exportar_conciliacao_excel(path: str, loja: str, rows: list) -> None:
    """
    Gera o Excel da conciliação da loja:
      • aba 'Conciliação' — uma linha por venda (+ valor/data do razão e diferença)
      • aba 'Resumo'      — quantidade e valor por status (fórmulas sobre a aba Conciliação)
    """
    from collections import Counter
    from datetime import datetime
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    FONT = "Arial"
    f_norm = Font(name=FONT, size=10)
    f_bold = Font(name=FONT, size=10, bold=True)
    f_hdr  = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    fill_h = PatternFill("solid", fgColor="1F3A5F")
    fills  = {
        "verde":    PatternFill("solid", fgColor="C6EFCE"),
        "amarelo":  PatternFill("solid", fgColor="FFEB9C"),
        "vermelho": PatternFill("solid", fgColor="FFC7CE"),
        "cinza":    PatternFill("solid", fgColor="E7E6E6"),
        "roxo":     PatternFill("solid", fgColor="E4D4F4"),
    }
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    NUM = '#,##0.00;(#,##0.00);"-"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliação"

    headers = ["Filial", "Data Venda", "NF / Doc", "Cliente", "Placa", "Ordem Serviço",
               "Cód. Venda", "Valor Venda (R$)", "Conta Partida", "Data Razão",
               "Valor Razão (R$)", "Diferença (R$)", "Status", "Critério", "Observação"]
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill, c.border = f_hdr, fill_h, borda
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    # NF repartida entre várias vendas → valor do razão não é comparável linha a linha
    nf_cont = Counter(r.get("doc") for r in rows if r.get("criterio") == "NF")

    for r in rows:
        linhas = r.get("linhas_contab") or ([r["linha_contab"]] if r.get("linha_contab") else [])
        val_razao, dt_razao = None, None
        if r.get("ajuste_conta"):                              # ajuste manual
            val_razao = r["valor_venda"]
            dt_razao  = _to_date(r.get("ajuste_data", ""))
        elif linhas and not (r.get("criterio") == "NF" and nf_cont[r.get("doc")] > 1):
            deb = [l for l in linhas if _parse_value(l.get("DEBITO", "")) > 0]
            if deb:
                val_razao = round(sum(_parse_value(l["DEBITO"]) for l in deb), 2)
                ds = [d for d in (_to_date(l.get("DATA", "")) for l in deb) if d]
                dt_razao = min(ds) if ds else None

        dt_venda = _to_date(r.get("data", ""))
        n = ws.max_row + 1
        ws.append([
            r.get("filial", ""),
            dt_venda if dt_venda else r.get("data", ""),
            r.get("doc", ""),
            r.get("cliente", ""),
            r.get("placa", ""),
            r.get("ordem", ""),
            r.get("cod_venda", ""),
            r["valor_venda"],
            r.get("conta_contab", ""),
            dt_razao,
            val_razao,
            f'=IF(K{n}="","",H{n}-K{n})',
            r.get("status", ""),
            r.get("criterio", ""),
            r.get("obs", ""),
        ])
        ws.cell(n, 13).fill = fills.get(r.get("status_tag"), fills["cinza"])

    ultima = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=ultima):
        for c in row:
            c.font, c.border = f_norm, borda
            c.alignment = Alignment(vertical="top", wrap_text=(c.column == 15))
        row[1].number_format = row[9].number_format = "dd/mm/yyyy"
        for i in (7, 10, 11):
            row[i].number_format = NUM
    for col, w in zip("ABCDEFGHIJKLMNO", (7, 12, 14, 38, 10, 14, 12, 16, 14, 12, 16, 14, 18, 12, 70)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:O{ultima}"

    # ── Resumo ──
    rs = wb.create_sheet("Resumo")
    rs["A1"] = "Conciliação Goodcard — Vendas × Razão"
    rs["A1"].font = Font(name=FONT, size=14, bold=True)
    rs["A2"], rs["B2"] = "Loja:", loja
    rs["A3"], rs["B3"] = "Gerado em:", datetime.now().strftime("%d/%m/%Y %H:%M")
    for a in ("A2", "A3"):
        rs[a].font = f_bold
    for a in ("B2", "B3"):
        rs[a].font = f_norm

    for i, h in enumerate(["Status", "Qtde", "Valor (R$)", "% do valor"], start=1):
        c = rs.cell(5, i, h)
        c.font, c.fill, c.border = f_hdr, fill_h, borda
        c.alignment = Alignment(horizontal="center")
    grupos = [("✔ OK", "verde"), ("⚠ Atenção", "amarelo"),
              ("✖ Não encontrado", "vermelho"), ("— Fora do período", "cinza"),
              ("↪ Lançada na 2157", "roxo")]
    rng_s = f"'Conciliação'!$M$2:$M${ultima}"
    rng_v = f"'Conciliação'!$H$2:$H${ultima}"
    for k, (txt, tag) in enumerate(grupos):
        n = 6 + k
        rs.cell(n, 1, txt)
        rs.cell(n, 2, f'=COUNTIF({rng_s},A{n}&"*")')
        rs.cell(n, 3, f'=SUMIF({rng_s},A{n}&"*",{rng_v})')
        rs.cell(n, 4, f"=IF($C$11=0,0,C{n}/$C$11)")
        rs.cell(n, 1).fill = fills[tag]
    rs.cell(11, 1, "Total")
    rs.cell(11, 2, "=SUM(B6:B10)")
    rs.cell(11, 3, "=SUM(C6:C10)")
    rs.cell(11, 4, "=SUM(D6:D10)")
    for row in rs.iter_rows(min_row=6, max_row=11, max_col=4):
        for c in row:
            c.font = f_bold if row[0].row == 11 else f_norm
            c.border = borda
        row[1].number_format = "#,##0"
        row[2].number_format = NUM
        row[3].number_format = "0.0%"
    rs["A13"] = ("Status e valores calculados por fórmula sobre a aba 'Conciliação'. "
                 "'Fora do período' = venda fora do período coberto pelo razão importado. "
                 "'Lançada na 2157' = não está na conta principal, mas foi achada na 2157.")
    rs["A13"].font = Font(name=FONT, size=9, italic=True)
    rs.column_dimensions["A"].width = 22
    for col in "BCD":
        rs.column_dimensions[col].width = 16

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ── Diálogo de importação de bases ────────────────────────────────────────── #
class _ImportDialog(tk.Toplevel):
    """
    Janelinha modal para importar as duas bases (Vendas + Contabilidade).
    Retorna os DataFrames via self.df_vendas / self.df_contab após confirmar.
    """

    def __init__(self, parent, loja_nome: str):
        super().__init__(parent)
        self.title(f"Importar Bases — {loja_nome}")
        self.geometry("560x390")
        self.resizable(False, False)
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)
        self.grab_set()
        self.transient(parent)

        self.df_vendas  = None
        self.df_contab  = None
        self.meta_contab = {}
        self.df_contab2 = None
        self._path_c2   = tk.StringVar(value="Nenhum arquivo selecionado (opcional)")
        self._path_v    = tk.StringVar(value="Nenhum arquivo selecionado")
        self._path_c    = tk.StringVar(value="Nenhum arquivo selecionado")
        self._confirmed = False

        self._build()
        self._center(parent)

    # ── Layout ── #
    def _build(self):
        # Toolbar interna
        bar = tk.Frame(self, bg=theme.PRIMARY, height=44)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="📂  Importar Bases Excel",
                 bg=theme.PRIMARY, fg="white",
                 font=("Segoe UI Semibold", 12)).pack(side="left", padx=16, pady=10)

        body = tk.Frame(self, bg=theme.BG_APP, padx=24, pady=16)
        body.pack(fill="both", expand=True)

        # Vendas
        self._row_import(body, "📈  Base de Vendas", self._path_v,
                         self._browse_vendas, row=0)
        # Contabilidade
        self._row_import(body, "📒  Base de Contabilidade", self._path_c,
                         self._browse_contab, row=1)
        # Razão da conta 2157 (opcional)
        self._row_import(body, "🔎  Razão conta 2157 (opcional — vendas lançadas por engano)",
                         self._path_c2, self._browse_contab2, row=2)

        # Botões
        btn_frame = tk.Frame(self, bg=theme.BG_APP, padx=24)
        btn_frame.pack(fill="x", pady=(0, 16))

        tk.Button(btn_frame, text="✔  Confirmar",
                  bg=theme.PRIMARY, fg="white",
                  activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                  font=("Segoe UI Semibold", 11), relief="flat", bd=0,
                  cursor="hand2", padx=18, pady=7,
                  command=self._confirm).pack(side="right", padx=(8, 0))

        tk.Button(btn_frame, text="Cancelar",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 11), relief="flat", bd=0,
                  cursor="hand2", padx=14, pady=7,
                  command=self.destroy).pack(side="right")

    def _row_import(self, parent, label: str, var: tk.StringVar, cmd, row: int):
        tk.Label(parent, text=label, bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 10)).grid(
            row=row*2, column=0, sticky="w", pady=(8, 2))

        frm = tk.Frame(parent, bg=theme.BG_APP)
        frm.grid(row=row*2+1, column=0, sticky="ew", pady=(0, 6))
        parent.columnconfigure(0, weight=1)

        tk.Label(frm, textvariable=var, bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9), anchor="w", padx=8, pady=5,
                 width=48, relief="flat", wraplength=400, justify="left").pack(side="left", fill="x", expand=True)

        tk.Button(frm, text="Procurar",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 10), relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=5,
                  command=cmd).pack(side="left", padx=(6, 0))

    # ── Lógica ── #
    def _browse_vendas(self):
        p = filedialog.askopenfilename(
            title="Selecionar Base de Vendas (Goodcard)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not p:
            return
        try:
            df = _ler_vendas(p)
        except Exception as e:
            messagebox.showerror("Erro ao ler Vendas", str(e), parent=self)
            return
        self.df_vendas = df
        self._path_v.set(f"{os.path.basename(p)}\n{_resumo_vendas(df)}")

    def _browse_contab(self):
        p = filedialog.askopenfilename(
            title="Selecionar Razão (conta Goodcard)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not p:
            return
        try:
            df, meta = _ler_razao(p)
        except Exception as e:
            messagebox.showerror("Erro ao ler Razão", str(e), parent=self)
            return
        self.df_contab, self.meta_contab = df, meta
        self._path_c.set(f"{os.path.basename(p)}\n{_resumo_razao(df, meta)}")

    def _browse_contab2(self):
        p = filedialog.askopenfilename(
            title="Selecionar Razão da conta 2157",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not p:
            return
        try:
            df, meta = _ler_razao(p)
        except Exception as e:
            messagebox.showerror("Erro ao ler Razão 2157", str(e), parent=self)
            return
        self.df_contab2 = df
        self._path_c2.set(f"{os.path.basename(p)}\n{_resumo_razao(df, meta)}")

    def _confirm(self):
        if self.df_vendas is None or self.df_contab is None:
            messagebox.showwarning("Atenção",
                "Importe as duas bases antes de confirmar.", parent=self)
            return
        self._confirmed = True
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width()  // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w}x{h}+{pw - w//2}+{ph - h//2}")


# ── Frame principal do módulo ──────────────────────────────────────────────── #
class LojasFrame(tk.Frame):
    """
    Frame completo do módulo Lojas.
    Integra-se ao MenuPrincipal exatamente como _CofreFrame.
    """

    def __init__(self, master):
        super().__init__(master, bg=theme.BG_APP)
        self._data        = _load_data()        # {"conciliacoes": {...}}
        self._lojas_db    = []                  # lista de dicts {loja, cnpj, regional}
        self._loja_atual  = None                # nome da loja selecionada
        self._df_vendas   = None
        self._df_contab   = None
        self._concil_rows = []                  # linhas exibidas na tabela

        self._build_ui()
        self._carregar_lojas_db()               # carrega do Supabase após UI pronta

    # ══════════════════════════════════════════════════════════════════════════ #
    #   LAYOUT PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════ #
    def _build_ui(self):
        # ── Toolbar ── #
        toolbar = tk.Frame(self, bg=theme.PRIMARY, height=52)
        toolbar.pack(fill="x", side="top")
        toolbar.pack_propagate(False)
        tk.Label(toolbar, text="🏪  GOODCARD",
                 bg=theme.PRIMARY, fg="white",
                 font=("Segoe UI Semibold", 14)).pack(side="left", padx=20, pady=10)
        tk.Label(toolbar, text=versao.texto(),
                 bg=theme.PRIMARY, fg="white",
                 font=("Segoe UI", 10)).pack(side="right", padx=20)

        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")

        # ── Área de seleção de loja + importação ── #
        self._build_selector()

        # ── Separador ── #
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x", padx=24)

        # ── Tabela de conciliação ── #
        self._build_concil_table()

        # ── Separador ── #
        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x", padx=24, pady=(8, 0))

        # ── Painel de resumo por loja ── #
        self._build_summary_panel()

        # ── Status bar ── #
        status = tk.Frame(self, bg=theme.BG_APP, height=26)
        status.pack(fill="x", side="bottom")
        status.pack_propagate(False)
        self._status_var = tk.StringVar(value="Selecione uma loja para começar.")
        tk.Label(status, textvariable=self._status_var,
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=12, pady=4)

    # ── Seletor de loja + botões de ação ── #
    def _build_selector(self):
        bar = tk.Frame(self, bg=theme.BG_APP, padx=24, pady=14)
        bar.pack(fill="x")

        # Filtro: Regional
        tk.Label(bar, text="Regional:",
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 11)).pack(side="left")

        self._regional_var = tk.StringVar(value="Todas")
        self._combo_regional = ttk.Combobox(bar, textvariable=self._regional_var,
                                            state="readonly", width=14,
                                            font=("Segoe UI", 11))
        self._combo_regional.pack(side="left", padx=(8, 0))
        self._combo_regional.bind("<<ComboboxSelected>>", self._on_regional_selected)

        # Seletor: Loja
        tk.Label(bar, text="Loja:",
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 11)).pack(side="left", padx=(16, 0))

        self._loja_var = tk.StringVar()
        self._combo = ttk.Combobox(bar, textvariable=self._loja_var,
                                   state="readonly", width=30,
                                   font=("Segoe UI", 11))
        self._combo.pack(side="left", padx=(8, 0))
        self._combo.bind("<<ComboboxSelected>>", self._on_loja_selected)

        # Label CNPJ (aparece ao selecionar loja)
        self._cnpj_var = tk.StringVar(value="")
        tk.Label(bar, textvariable=self._cnpj_var,
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(10, 0))

        # Botão recarregar do banco
        tk.Button(bar, text="🔄",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 11), relief="flat", bd=0,
                  cursor="hand2", padx=8, pady=5,
                  command=self._carregar_lojas_db).pack(side="left", padx=(10, 0))

        # Botão: importar bases
        self._btn_import = tk.Button(bar, text="📂  Importar Bases",
                  bg=theme.PRIMARY, fg="white",
                  activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                  font=("Segoe UI Semibold", 10), relief="flat", bd=0,
                  cursor="hand2", padx=12, pady=5,
                  state="disabled",
                  command=self._importar_bases)
        self._btn_import.pack(side="left", padx=(8, 0))

        # Botão: exportar conciliação para Excel
        self._btn_export = tk.Button(bar, text="📤  Exportar Excel",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI Semibold", 10), relief="flat", bd=0,
                  cursor="hand2", padx=12, pady=5,
                  state="disabled",
                  command=self._exportar_excel)
        self._btn_export.pack(side="left", padx=(8, 0))

    # ── Tabela de conciliação ── #
    def _build_concil_table(self):
        hdr = tk.Frame(self, bg=theme.BG_APP, padx=24)
        hdr.pack(fill="x", pady=(10, 4))
        tk.Label(hdr, text="📊  Conciliação — Vendas × Contabilidade",
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 12)).pack(side="left")

        self._loja_label = tk.Label(hdr, text="",
                 bg=theme.BG_APP, fg=theme.PRIMARY,
                 font=("Segoe UI Semibold", 12))
        self._loja_label.pack(side="left", padx=(8, 0))

        # Frame da tabela com scroll
        tbl_wrap = tk.Frame(self, bg=theme.CARD_BG, padx=2, pady=2)
        tbl_wrap.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        cols = ("data", "doc", "valor_venda", "conta_contab", "status", "obs")
        headers = ("Data", "NF / Doc", "Valor Venda (R$)", "Conta Partida", "Status", "Observação")
        col_w   = (90, 100, 120, 100, 130, 420)

        self._tree = ttk.Treeview(tbl_wrap, columns=cols, show="headings",
                                  style="Treeview", selectmode="browse")
        for col, hdr_txt, w in zip(cols, headers, col_w):
            self._tree.heading(col, text=hdr_txt)
            self._tree.column(col, width=w, minwidth=80, anchor="center")

        vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<Double-1>", self._on_row_double_click)
        self._tree_placeholder()

    def _tree_placeholder(self):
        self._tree.delete(*self._tree.get_children())
        self._tree.insert("", "end", values=(
            "—", "—", "—", "—", "Selecione uma loja e importe as bases", ""))

    # ── Painel de resumo ── #
    def _build_summary_panel(self):
        # Cabeçalho + filtro
        hdr = tk.Frame(self, bg=theme.BG_APP, padx=24)
        hdr.pack(fill="x", pady=(10, 6))

        tk.Label(hdr, text="📋  Resumo das Lojas",
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 12)).pack(side="left")

        # Filtro Mês
        tk.Label(hdr, text="Mês:",
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10)).pack(side="left", padx=(24, 4))
        meses = ["Todos", "01 - Jan", "02 - Fev", "03 - Mar", "04 - Abr",
                 "05 - Mai", "06 - Jun", "07 - Jul", "08 - Ago",
                 "09 - Set", "10 - Out", "11 - Nov", "12 - Dez"]
        self._filter_mes = tk.StringVar(value="Todos")
        ttk.Combobox(hdr, textvariable=self._filter_mes, values=meses,
                     state="readonly", width=10,
                     font=("Segoe UI", 10)).pack(side="left")

        # Filtro Ano
        tk.Label(hdr, text="Ano:",
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10)).pack(side="left", padx=(12, 4))
        anos = ["Todos"] + [str(a) for a in range(2020, 2031)]
        self._filter_ano = tk.StringVar(value="Todos")
        ttk.Combobox(hdr, textvariable=self._filter_ano, values=anos,
                     state="readonly", width=8,
                     font=("Segoe UI", 10)).pack(side="left")

        tk.Button(hdr, text="🔍  Filtrar",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 10), relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=4,
                  command=self._refresh_cards).pack(side="left", padx=(10, 0))

        # Área de cards (com scroll horizontal via Canvas)
        canvas_wrap = tk.Frame(self, bg=theme.BG_APP)
        canvas_wrap.pack(fill="x", padx=24, pady=(0, 12))

        self._cards_canvas = tk.Canvas(canvas_wrap, bg=theme.BG_APP,
                                       highlightthickness=0, height=148)
        hsb = ttk.Scrollbar(canvas_wrap, orient="horizontal",
                             command=self._cards_canvas.xview)
        self._cards_canvas.configure(xscrollcommand=hsb.set)
        self._cards_canvas.pack(fill="x", expand=True)
        hsb.pack(fill="x")

        self._cards_inner = tk.Frame(self._cards_canvas, bg=theme.BG_APP)
        self._cards_win = self._cards_canvas.create_window(
            (0, 0), window=self._cards_inner, anchor="nw")
        self._cards_inner.bind("<Configure>", self._on_cards_resize)

        self._refresh_cards()

    def _on_cards_resize(self, _event=None):
        self._cards_canvas.configure(
            scrollregion=self._cards_canvas.bbox("all"))

    # ══════════════════════════════════════════════════════════════════════════ #
    #   LÓGICA
    # ══════════════════════════════════════════════════════════════════════════ #

    # ── Carregamento do Supabase ── #
    def _carregar_lojas_db(self):
        """Busca lojas no Supabase e atualiza combos."""
        self._status_var.set("Carregando lojas do banco de dados...")
        self.update_idletasks()
        self._lojas_db = _fetch_lojas_supabase()

        # Monta lista de regionais únicas (ordenadas)
        regionais = sorted({r["regional"] for r in self._lojas_db if r.get("regional")})
        self._combo_regional["values"] = ["Todas"] + regionais
        self._regional_var.set("Todas")

        self._refresh_combo()
        n = len(self._lojas_db)
        self._status_var.set(f"{n} loja(s) carregada(s) do banco. Selecione uma loja para começar.")
        self._refresh_cards()

    # ── Combo de lojas ── #
    def _refresh_combo(self):
        """Filtra lojas pela regional selecionada e atualiza o combobox de lojas."""
        regional = self._regional_var.get() if hasattr(self, "_regional_var") else "Todas"
        if regional == "Todas":
            lojas_filtradas = self._lojas_db
        else:
            lojas_filtradas = [l for l in self._lojas_db if l.get("regional") == regional]

        nomes = [l["loja"] for l in lojas_filtradas]
        self._combo["values"] = nomes if nomes else [""]
        # Limpa seleção se a loja atual sumiu do filtro
        if self._loja_var.get() not in nomes:
            self._loja_var.set("")
            self._cnpj_var.set("")

    def _on_regional_selected(self, _event=None):
        """Ao trocar regional, refiltra o combo de lojas."""
        self._loja_var.set("")
        self._cnpj_var.set("")
        self._loja_atual = None
        self._loja_label.config(text="")
        self._btn_import.config(state="disabled")
        self._concil_rows = []
        self._atualizar_btn_export()
        self._tree_placeholder()
        self._refresh_combo()

    def _on_loja_selected(self, _event=None):
        nome = self._loja_var.get()
        if not nome:
            return
        self._loja_atual = nome
        self._loja_label.config(text=f"›  {nome}")
        self._btn_import.config(state="normal")

        # Exibe CNPJ da loja selecionada
        loja_info = next((l for l in self._lojas_db if l["loja"] == nome), {})
        cnpj = loja_info.get("cnpj", "")
        self._cnpj_var.set(f"CNPJ: {cnpj}" if cnpj else "")

        # Carrega conciliação existente se houver
        rows = self._data.get("conciliacoes", {}).get(nome, [])
        self._concil_rows = rows
        self._refresh_table(rows)
        self._status_var.set(
            f"Loja: {nome}  |  "
            f"{len(rows)} registro(s) de conciliação carregado(s).")
        self._atualizar_btn_export()

    # ── Exportação ── #
    def _atualizar_btn_export(self):
        ok = bool(self._loja_atual and self._concil_rows)
        self._btn_export.config(state="normal" if ok else "disabled")

    def _exportar_excel(self):
        if not self._loja_atual or not self._concil_rows:
            return
        from datetime import datetime
        nome_seguro = re.sub(r'[\\/:*?"<>|]+', "_", self._loja_atual).strip()
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Exportar conciliação para Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"Conciliacao_Goodcard_{nome_seguro}_{datetime.now():%Y%m%d}.xlsx")
        if not path:
            return
        try:
            exportar_conciliacao_excel(path, self._loja_atual, self._concil_rows)
        except PermissionError:
            messagebox.showerror(
                "Exportar Excel",
                "Não foi possível gravar o arquivo.\nSe ele estiver aberto no Excel, feche-o e tente novamente.",
                parent=self)
            return
        except Exception as e:
            messagebox.showerror("Exportar Excel", f"Erro ao exportar:\n{e}", parent=self)
            return
        self._status_var.set(f"Exportado: {path}")
        if messagebox.askyesno("Exportar Excel", "Exportação concluída.\nDeseja abrir o arquivo?",
                               parent=self):
            try:
                os.startfile(path)                    # Windows
            except AttributeError:
                pass
            except Exception as e:
                messagebox.showwarning("Exportar Excel", f"Não foi possível abrir:\n{e}", parent=self)

    # ── Importação ── #
    def _importar_bases(self):
        if not self._loja_atual:
            return
        dlg = _ImportDialog(self, self._loja_atual)
        self.wait_window(dlg)
        if not dlg._confirmed:
            return

        self._df_vendas = dlg.df_vendas
        self._df_contab = dlg.df_contab
        rows = self._conciliar(self._df_vendas, self._df_contab, dlg.meta_contab,
                               dlg.df_contab2)
        if not rows:
            self._status_var.set("Conciliação não gerada — verifique as bases importadas.")
            return
        self._concil_rows = rows

        # Persiste
        if "conciliacoes" not in self._data:
            self._data["conciliacoes"] = {}
        self._data["conciliacoes"][self._loja_atual] = rows
        _save_data(self._data)

        self._refresh_table(rows)
        self._refresh_cards()
        self._atualizar_btn_export()
        n = lambda tag: sum(1 for r in rows if r.get("status_tag") == tag)
        self._status_var.set(
            f"Loja: {self._loja_atual}  |  "
            f"{len(rows)} venda(s)  —  "
            f"✔ {n('verde')} OK  |  "
            f"⚠ {n('amarelo')} Atenção  |  "
            f"✖ {n('vermelho')} Não encontrado  |  "
            f"— {n('cinza')} Fora do período"
            + (f"  |  ↪ {n('roxo')} na 2157" if n('roxo') else "")
            + (f"  |  {self._aviso}" if getattr(self, "_aviso", "") else ""))

    # ── Conciliação ── #
    def _conciliar(self, df_v, df_c, meta_c=None, df_c2=None) -> list:
        """Chama a conciliação (módulo-level `conciliar_bases`) e trata erros na UI."""
        self._aviso = ""
        try:
            rows, aviso = conciliar_bases(df_v, df_c, meta_c or {}, df_c2)
        except ValueError as e:
            messagebox.showerror("Erro na conciliação", str(e), parent=self)
            return []
        self._aviso = aviso
        return rows

    # ── Tabela ── #
    def _refresh_table(self, rows: list):
        self._tree.delete(*self._tree.get_children())
        if not rows:
            self._tree_placeholder()
            return
        for r in rows:
            tag = r.get("status_tag", "vermelho")
            self._tree.insert("", "end", values=(
                r.get("data", "—"),
                r.get("doc",  "—"),
                f"R$ {_brl(r['valor_venda'])}",
                r.get("conta_contab", "—"),
                r.get("status", "—"),
                r.get("obs", ""),
            ), tags=(tag,))

        self._tree.tag_configure("verde",    foreground=theme.ACCENT_GREEN)
        self._tree.tag_configure("amarelo",  foreground=theme.ACCENT_ORANGE)
        self._tree.tag_configure("vermelho", foreground=theme.ACCENT_RED)
        self._tree.tag_configure("cinza",    foreground=theme.FG_MUTED)
        self._tree.tag_configure("roxo",     foreground="#b58cf0")

    # ── Duplo clique → detalhes ── #
    # ── Duplo clique → detalhes ou ajuste manual ── #
    def _on_row_double_click(self, event):
        item = self._tree.focus()
        if not item:
            return
        idx = self._tree.index(item)
        if idx >= len(self._concil_rows):
            return
        r = self._concil_rows[idx]
        if r.get("status_tag") == "vermelho":
            self._abrir_ajuste_manual(idx, r)
        else:
            self._abrir_detalhes(r)

    def _abrir_detalhes(self, r):
        linhas = r.get("linhas_contab") or ([r["linha_contab"]] if r.get("linha_contab") else [])

        win = tk.Toplevel(self)
        win.title("Detalhes da Linha")
        win.geometry("600x640")
        win.resizable(True, True)
        win.configure(bg=theme.BG_APP)
        theme.apply_theme(win)
        win.grab_set()
        win.transient(self)

        bar = tk.Frame(win, bg=theme.PRIMARY, height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="🔍  Detalhes da Conciliação",
                 bg=theme.PRIMARY, fg="white",
                 font=("Segoe UI Semibold", 12)).pack(side="left", padx=16, pady=8)

        tk.Button(win, text="Fechar",
                  bg=theme.PRIMARY, fg="white",
                  activebackground=theme.PRIMARY_HOVER,
                  font=("Segoe UI Semibold", 10), relief="flat", bd=0,
                  cursor="hand2", padx=16, pady=6,
                  command=win.destroy).pack(side="bottom", pady=12)

        canvas = tk.Canvas(win, bg=theme.BG_APP, highlightthickness=0)
        vsb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        body = tk.Frame(canvas, bg=theme.BG_APP, padx=24, pady=12)
        canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _linha(row, lbl, val, wrap=380):
            tk.Label(body, text=lbl, bg=theme.BG_APP, fg=theme.FG_MUTED,
                     font=("Segoe UI", 10), width=16, anchor="nw").grid(
                row=row, column=0, sticky="nw", pady=2)
            tk.Label(body, text=str(val), bg=theme.BG_APP, fg=theme.FG_TEXT,
                     font=("Segoe UI Semibold", 10), anchor="w", justify="left",
                     wraplength=wrap).grid(row=row, column=1, sticky="w", pady=2)

        tk.Label(body, text="📈  Venda", bg=theme.BG_APP, fg=theme.PRIMARY,
                 font=("Segoe UI Semibold", 11)).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        venda_info = [
            ("Data:",         f"{r.get('data', '—')} {r.get('hora', '')}".strip()),
            ("NF / Doc:",     r.get("doc", "—")),
            ("Cliente:",      r.get("cliente") or "—"),
            ("Placa / OS:",   f"{r.get('placa') or '—'}  /  {r.get('ordem') or '—'}"),
            ("Cód. de Venda:", r.get("cod_venda") or "—"),
            ("Valor Venda:",  f"R$ {_brl(r['valor_venda'])}"),
            ("Status:",       r.get("status", "—")),
            ("Observação:",   r.get("obs") or "—"),
        ]
        row = 1
        for lbl, val in venda_info:
            _linha(row, lbl, val)
            row += 1

        tk.Frame(body, bg=theme.GRID_LINE, height=1).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10)
        row += 1
        tk.Label(body, text="📒  Contabilidade", bg=theme.BG_APP, fg=theme.PRIMARY,
                 font=("Segoe UI Semibold", 11)).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
        row += 1

        if linhas:
            for k, lc in enumerate(linhas, start=1):
                if len(linhas) > 1:
                    tk.Label(body, text=f"Lançamento {k}", bg=theme.BG_APP, fg=theme.ACCENT_ORANGE,
                             font=("Segoe UI Semibold", 10)).grid(
                        row=row, column=0, columnspan=2, sticky="w", pady=(6, 2))
                    row += 1
                for col, val in lc.items():
                    if str(val).strip() == "":
                        continue
                    _linha(row, f"{col}:", val)
                    row += 1
        else:
            tk.Label(body, text="Nenhum lançamento contábil encontrado.",
                     bg=theme.BG_APP, fg=theme.ACCENT_RED,
                     font=("Segoe UI", 10)).grid(
                row=row, column=0, columnspan=2, sticky="w")

    def _abrir_ajuste_manual(self, idx: int, r: dict):
        win = tk.Toplevel(self)
        win.title("Ajuste Manual")
        win.geometry("480x320")
        win.resizable(False, False)
        win.configure(bg=theme.BG_APP)
        theme.apply_theme(win)
        win.grab_set()
        win.transient(self)

        bar = tk.Frame(win, bg="#7a2020", height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="✏️  Ajuste Manual — Lançamento Não Encontrado",
                 bg="#7a2020", fg="white",
                 font=("Segoe UI Semibold", 11)).pack(side="left", padx=16, pady=8)

        body = tk.Frame(win, bg=theme.BG_APP, padx=24, pady=16)
        body.pack(fill="both", expand=True)

        # Info da venda (somente leitura)
        tk.Label(body, text="Venda:", bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), width=16, anchor="w").grid(
            row=0, column=0, sticky="w", pady=2)
        tk.Label(body, text=f"{r.get('data','—')}  |  R$ {_brl(r['valor_venda'])}",
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI Semibold", 10)).grid(row=0, column=1, sticky="w", pady=2)

        tk.Frame(body, bg=theme.GRID_LINE, height=1).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=10)

        # Campos editáveis
        campos = [
            ("Data Contábil:",  r.get("ajuste_data",  "")),
            ("Conta Partida:",  r.get("ajuste_conta", "")),
            ("Valor:",          r.get("ajuste_valor", "")),
        ]
        entries = {}
        for i, (lbl, val) in enumerate(campos, start=2):
            tk.Label(body, text=lbl, bg=theme.BG_APP, fg=theme.FG_MUTED,
                     font=("Segoe UI", 10), width=16, anchor="w").grid(
                row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=val)
            e = tk.Entry(body, textvariable=var,
                         bg=theme.CARD_BG, fg=theme.FG_TEXT,
                         insertbackground=theme.FG_TEXT,
                         font=("Segoe UI", 10), relief="flat", bd=6, width=26)
            e.grid(row=i, column=1, sticky="w", pady=4)
            entries[lbl] = var

        def _salvar():
            data_aj  = entries["Data Contábil:"].get().strip()
            conta_aj = entries["Conta Partida:"].get().strip()
            valor_aj = entries["Valor:"].get().strip()

            if not conta_aj:
                messagebox.showwarning("Atenção", "Informe a Conta Partida.", parent=win)
                return

            # Lançamento informado manualmente → considerado conciliado
            novo_status     = "✔ OK (Manual)"
            novo_status_tag = "verde"

            valor_num = _parse_value(valor_aj) if valor_aj else r["valor_venda"]

            # Atualiza o row
            self._concil_rows[idx]["ajuste_data"]  = data_aj
            self._concil_rows[idx]["ajuste_conta"] = conta_aj
            self._concil_rows[idx]["ajuste_valor"] = valor_aj
            self._concil_rows[idx]["status"]       = novo_status
            self._concil_rows[idx]["status_tag"]   = novo_status_tag
            self._concil_rows[idx]["conta_contab"] = conta_aj
            self._concil_rows[idx]["obs"]          = "Ajuste manual"
            if data_aj:
                self._concil_rows[idx]["data"] = data_aj
            self._concil_rows[idx]["valor_venda"]  = valor_num

            # Persiste
            self._data["conciliacoes"][self._loja_atual] = self._concil_rows
            _save_data(self._data)

            self._refresh_table(self._concil_rows)
            self._refresh_cards()
            win.destroy()

        btn_f = tk.Frame(win, bg=theme.BG_APP, padx=24)
        btn_f.pack(fill="x", pady=(0, 16))
        tk.Button(btn_f, text="💾  Salvar Ajuste",
                  bg=theme.PRIMARY, fg="white",
                  activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                  font=("Segoe UI Semibold", 10), relief="flat", bd=0,
                  cursor="hand2", padx=16, pady=7,
                  command=_salvar).pack(side="right", padx=(8, 0))
        tk.Button(btn_f, text="Cancelar",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 10), relief="flat", bd=0,
                  cursor="hand2", padx=12, pady=7,
                  command=win.destroy).pack(side="right")



    # ── Cards de resumo ── #
    def _refresh_cards(self):
        for w in self._cards_inner.winfo_children():
            w.destroy()

        mes_filtro = self._filter_mes.get()
        ano_filtro = self._filter_ano.get()
        concils    = self._data.get("conciliacoes", {})
        lojas      = [l["loja"] for l in self._lojas_db]

        if not lojas:
            tk.Label(self._cards_inner, text="Nenhuma loja cadastrada.",
                     bg=theme.BG_APP, fg=theme.FG_MUTED,
                     font=("Segoe UI", 10)).pack(padx=16, pady=20)
            self._on_cards_resize()
            return

        for i, loja in enumerate(lojas):
            rows = concils.get(loja, [])

            # Aplica filtro
            def _match(r):
                data = r.get("data", "")
                try:
                    import pandas as pd
                    dt = pd.to_datetime(data, dayfirst=True, errors="coerce")
                    if pd.isna(dt):
                        return True
                    if mes_filtro != "Todos":
                        m_num = mes_filtro.split(" - ")[0]
                        if str(dt.month).zfill(2) != m_num:
                            return False
                    if ano_filtro != "Todos":
                        if str(dt.year) != ano_filtro:
                            return False
                except Exception:
                    pass
                return True

            filtered = [r for r in rows if _match(r)]

            conc      = [r for r in filtered if r.get("status_tag") != "cinza"]
            n_fora    = len(filtered) - len(conc)
            total_v   = sum(r.get("valor_venda", 0)  for r in conc)
            n_div     = sum(1 for r in conc if r.get("status_tag") != "verde")
            status_cor = theme.ACCENT_GREEN if n_div == 0 and conc else \
                         (theme.ACCENT_ORANGE if conc else theme.FG_MUTED)

            card = tk.Frame(self._cards_inner, bg=theme.CARD_BG,
                            relief="flat", bd=0, cursor="hand2")
            card.pack(side="left", padx=8, pady=6, ipadx=0, ipady=0)

            # Borda superior colorida
            tk.Frame(card, bg=status_cor, height=3).pack(fill="x")

            body = tk.Frame(card, bg=theme.CARD_BG, padx=14, pady=10)
            body.pack()

            tk.Label(body, text=f"🏪  {loja}",
                     bg=theme.CARD_BG, fg=theme.FG_TEXT,
                     font=("Segoe UI Semibold", 11),
                     wraplength=160, justify="left").pack(anchor="w")

            tk.Frame(body, bg=theme.GRID_LINE, height=1).pack(
                fill="x", pady=(6, 6))

            def _info_row(parent, label, value, color=theme.FG_TEXT):
                f = tk.Frame(parent, bg=theme.CARD_BG)
                f.pack(fill="x", pady=1)
                tk.Label(f, text=label, bg=theme.CARD_BG, fg=theme.FG_MUTED,
                         font=("Segoe UI", 9), width=13, anchor="w").pack(side="left")
                tk.Label(f, text=value, bg=theme.CARD_BG, fg=color,
                         font=("Segoe UI Semibold", 9), anchor="e").pack(side="right")

            if conc:
                _info_row(body, "Vendas:", f"R$ {_brl(total_v)}", theme.FG_TEXT)
                st_txt  = "✔  OK" if n_div == 0 else f"⚠  {n_div} diverg."
                _info_row(body, "Status:", st_txt, status_cor)
                if n_fora:
                    _info_row(body, "Fora período:", str(n_fora), theme.FG_MUTED)
            else:
                tk.Label(body,
                         text=(f"{n_fora} venda(s) fora do período" if n_fora
                               else "Sem dados no período"),
                         bg=theme.CARD_BG, fg=theme.FG_MUTED,
                         font=("Segoe UI", 9)).pack(anchor="w", pady=4)

        self._on_cards_resize()


# ── Diálogo de texto simples ────────────────────────────────────────────────── #
class _InputDialog(tk.Toplevel):
    """Mini-diálogo para capturar um texto simples."""

    def __init__(self, parent, title: str, prompt: str):
        super().__init__(parent)
        self.title(title)
        self.geometry("340x140")
        self.resizable(False, False)
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)
        self.grab_set()
        self.transient(parent)
        self.result = ""

        tk.Label(self, text=prompt,
                 bg=theme.BG_APP, fg=theme.FG_TEXT,
                 font=("Segoe UI", 11)).pack(padx=24, pady=(18, 6), anchor="w")

        self._var = tk.StringVar()
        entry = tk.Entry(self, textvariable=self._var,
                         bg=theme.CARD_BG, fg=theme.FG_TEXT,
                         insertbackground=theme.FG_TEXT,
                         font=("Segoe UI", 11), relief="flat", bd=6)
        entry.pack(fill="x", padx=24)
        entry.focus_set()
        entry.bind("<Return>", lambda _: self._ok())

        bf = tk.Frame(self, bg=theme.BG_APP)
        bf.pack(fill="x", padx=24, pady=12)
        tk.Button(bf, text="OK",
                  bg=theme.PRIMARY, fg="white",
                  activebackground=theme.PRIMARY_HOVER, activeforeground="white",
                  font=("Segoe UI Semibold", 10), relief="flat", bd=0,
                  cursor="hand2", padx=16, pady=5,
                  command=self._ok).pack(side="right")
        tk.Button(bf, text="Cancelar",
                  bg=theme.SURFACE2, fg=theme.FG_TEXT,
                  activebackground=theme.GRID_LINE, activeforeground="white",
                  font=("Segoe UI", 10), relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=5,
                  command=self.destroy).pack(side="right", padx=(0, 6))

    def _ok(self):
        self.result = self._var.get()
        self.destroy()


# ── Factory para o MenuPrincipal ──────────────────────────────────────────── #
def _open_lojas(root: tk.Tk, on_close):
    """
    Factory compatível com o padrão do MenuPrincipal.
    Cria a Toplevel, injeta o LojasFrame e retorna a janela.
    """
    win = tk.Toplevel(root)
    win.title("")
    win.geometry("1100x680")
    win.minsize(900, 560)
    win.resizable(True, True)
    win.configure(bg=theme.BG_APP)
    theme.apply_theme(win)

    frame = LojasFrame(win)
    frame.pack(fill="both", expand=True)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.protocol("WM_DELETE_WINDOW", lambda: on_close(win))
    return win


# ── Entrada no registro de módulos (cole em MODULES do menu_principal.py) ─── #
LOJAS_MODULE = {
    "emoji":    "🏪",
    "title":    "Módulo de Lojas",
    "subtitle": "Conciliação Vendas × Contabilidade por loja",
    "factory":  _open_lojas,
}


# ── Execução standalone (teste direto) ── #
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Lojas — Standalone")
    root.geometry("1100x680")
    root.minsize(900, 560)
    theme.apply_theme(root)

    LojasFrame(root).pack(fill="both", expand=True)
    root.mainloop()