"""
prossegur/painel.py
Painel de conciliação — Prossegur
Grade visual interativa: eixo Y = lojas, eixo X = datas de depósito.
Filtros: status, regional (UF), mês/ano.
"""

import json
import os
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta
import calendar

from core.memorando_service import get_memorandos_por_loja_dia
from core import ui_theme as theme
from core import database as db
import holidays
from collections import defaultdict
from datetime import datetime, timedelta


HORA_CORTE = 16


def obter_data_credito_esperada(data_doc):
    return _proximo_util(data_doc)


def obter_janela_operacional(data_doc):

    inicio = data_doc.replace(
        hour=HORA_CORTE,
        minute=0,
        second=0,
        microsecond=0,
    )

    prox_util = obter_data_credito_esperada(data_doc)

    fim = prox_util.replace(
        hour=HORA_CORTE,
        minute=0,
        second=0,
        microsecond=0,
    )

    return inicio, fim, prox_util

FERIADOS_BR = holidays.Brazil(years=range(2024, 2030))

# ──────────────────────────────────────────────────────────────────────────── #
#  Carrega mapeamento loja → UF do JSON                                        #
# ──────────────────────────────────────────────────────────────────────────── #

import sys

def _base_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _load_loja_uf_map() -> dict[str, str]:
    json_path = os.path.join(_base_dir(), "codigos_cofres.json")
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        return {str(entry["loja"]): entry["uf"] for entry in data.get("lojas", [])}
    except Exception:
        return {}

LOJA_UF: dict[str, str] = _load_loja_uf_map()

# ──────────────────────────────────────────────────────────────────────────── #
#  Helpers de data                                                              #
# ──────────────────────────────────────────────────────────────────────────── #

def _parse(s: str) -> datetime | None:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y",
                "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None

def _proximo_util(dt: datetime) -> datetime:
    dt = dt + timedelta(days=1)
    while dt.weekday() >= 5 or dt.date() in FERIADOS_BR:
        dt = dt + timedelta(days=1)
    return dt

def _credito_esperado(dt: datetime) -> datetime:
    return _proximo_util(dt)

def _fmt_brl(v: float) -> str:
    return (
        f"R$ {abs(v):,.2f}"
        .replace(",", "X").replace(".", ",").replace("X", ".")
    )

def _fmt_date(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

HORA_CORTE_CRED = 16  # lançamentos após 16h contam no próximo dia útil (crédito)

# ──────────────────────────────────────────────────────────────────────────── #
#  Motor de conciliação — Prossegur                                             #
# ──────────────────────────────────────────────────────────────────────────── #

def _carregar_memorandos_para(depositos: list[dict]) -> dict[tuple[str, str], float]:
    meses: set[tuple[int, int]] = set()
    lojas: set[str] = set()
    for d in depositos:
        dt = _parse(d.get("data_corte", ""))
        if dt:
            meses.add((dt.year, dt.month))
            lojas.add(str(int(str(d["loja"]))))
    mem_idx: dict[tuple[str, str], float] = {}
    for (ano, mes) in meses:
        parcial = get_memorandos_por_loja_dia(ano, mes, list(lojas))
        mem_idx.update(parcial)

    # ── Enriquece com ajustes manuais que têm memo_remessa salvo ──────────────
    # Isso garante que g1_ok seja calculado corretamente ANTES do bloco 4,
    # em vez de depender do patch tardio de sobrescrita de status.
    ajustes = db.get_ajustes("prossegur")   # {(loja, data_deposito): dados}
    for (loja, data_dep), aj in ajustes.items():
        memo_manual = aj.get("memo_remessa")
        if memo_manual is not None:
            # Só sobrescreve se o ajuste tiver um valor real de memorando
            mem_idx[(loja, data_dep)] = float(memo_manual)
    # ─────────────────────────────────────────────────────────────────────────

    return mem_idx

def conciliar():
    """
    Nova lógica de conciliação — Depósito Ajust. × Créditos esperados.

    Por loja, ordena os depósitos cronologicamente (data_corte) e os
    créditos cronologicamente. Vai acumulando depósitos até igualar o
    valor de cada crédito (respeitando data_deposito <= data_credito).
    Quando a soma igualar dentro da tolerância → grupo OK.
    Se a soma ultrapassar → grupo divergente.
    Cada grupo gera uma linha no painel; a linha carrega a lista de
    depósitos que participaram da soma ("depositos_ajust").
    """

    depositos_resumo = db.get_resumo_depositos_prossegur()
    lancamentos      = db.get_resumo_depositos_prossegur_com_hora()
    creditos         = db.get_creditos_prossegur()
    mem_idx          = _carregar_memorandos_para(depositos_resumo)

    TOL = 0.01

    def _iguais(a, b):
        return abs(float(a or 0) - float(b or 0)) <= TOL

    # =========================================================
    # 1. ÍNDICE DOCUMENTAL  (data_doc, loja) → total diário
    #    usado para comparação com memorando
    # =========================================================

    dep_doc_idx: dict[tuple[str, str], float] = defaultdict(float)

    for d in depositos_resumo:
        loja_d = str(int(str(d["loja"])))
        dt_d   = _parse(d["data_corte"])
        if not dt_d:
            continue
        data_doc = _fmt_date(dt_d)
        dep_doc_idx[(data_doc, loja_d)] += float(d["total"])

    # ─────────────────────────────────────────────────────────────────

    # =========================================================
    # 1b. FILA DE LANÇAMENTOS INDIVIDUAIS por loja
    #     ordenada por data_hora — acumula lançamento a lançamento
    #     até igualar cada crédito (sem agrupar por dia)
    # =========================================================

    deps_por_loja: dict[str, list] = defaultdict(list)

    for lanc in lancamentos:
        loja = str(int(str(lanc["loja"])))
        dt   = _parse(lanc.get("data_hora") or lanc.get("data_deposito") or "")
        if not dt:
            continue
        data_doc = _fmt_date(dt)
        deps_por_loja[loja].append((dt, data_doc, float(lanc["valor"])))

    for loja in deps_por_loja:
        deps_por_loja[loja].sort(key=lambda x: x[0])

    # =========================================================
    # 2. ÍNDICE DE CRÉDITOS  (data_cred, loja) → total
    #    e lista ordenada de créditos por loja
    # =========================================================

    cred_idx: dict[tuple[str, str], float] = defaultdict(float)
    creds_por_loja: dict[str, list] = defaultdict(list)

    for c in creditos:
        loja      = str(int(str(c["loja"])))
        dt        = _parse(str(c["data"]))
        if not dt:
            continue
        data_cred = _fmt_date(dt)
        valor     = float(c["debito"])
        cred_idx[(data_cred, loja)] += valor
        creds_por_loja[loja].append((dt, data_cred, valor))

    for loja in creds_por_loja:
        creds_por_loja[loja].sort(key=lambda x: x[0])

    # =========================================================
    # 3. AGRUPAMENTO  depósitos → créditos  (por loja)
    # =========================================================

    linhas = []

    todas_lojas = set(deps_por_loja.keys()) | set(creds_por_loja.keys())

    for loja in sorted(todas_lojas, key=lambda x: int(x) if x.isdigit() else x):

        fila_deps  = list(deps_por_loja.get(loja, []))   # (dt, data_str, valor)
        fila_creds = list(creds_por_loja.get(loja, []))  # (dt, data_str, valor)

        dep_cursor = 0   # próximo depósito a consumir

        for (dt_cred, data_cred_str, valor_cred) in fila_creds:

            acum          = 0.0
            deps_do_grupo: list[dict] = []

            # Avança depósitos enquanto:
            #   - ainda há depósitos disponíveis
            #   - a data do depósito não ultrapassa a data do crédito
            #   - a soma acumulada ainda não igualou/ultrapassou o crédito

            i = dep_cursor
            while i < len(fila_deps):
                dt_dep, data_dep_str, valor_dep = fila_deps[i]

                # Depósitos após a hora de corte (ex: 17:25) pertencem ao
                # próximo dia útil — tratamos a data efetiva como o próximo util.
                if dt_dep.hour >= HORA_CORTE_CRED:
                    data_efetiva = _proximo_util(dt_dep).date()
                else:
                    data_efetiva = dt_dep.date()

                # restrição: data efetiva do depósito não pode ser posterior ao crédito
                if data_efetiva > dt_cred.date():
                    break

                proximo_acum = acum + valor_dep

                if proximo_acum > valor_cred + TOL:
                    # ultrapassaria — para aqui (divergente)
                    break

                acum += valor_dep
                deps_do_grupo.append({
                    # data_efetiva: usada para colorir o quadrinho correto na grade
                    "data_deposito":     data_efetiva.strftime("%d/%m/%Y"),
                    # data_doc: data documental original, usada para buscar dep_doc_idx e memorando
                    "data_doc":          data_dep_str,
                    "valor":             valor_dep,
                })
                i += 1

            # avança o cursor global de depósitos
            dep_cursor = i

            # ─── Grupo 1: Documental × Memorando ─────────────────────
            # dep×memo é verificado DIA A DIA — cada data documental
            # tem seu próprio memorando e não deve ser somada com outras.
            # g1_ok só é True se TODOS os dias do grupo passarem.
            if deps_do_grupo:
                data_ref   = deps_do_grupo[-1]["data_deposito"]
                dias_grupo = set(d["data_deposito"] for d in deps_do_grupo)
                dias_doc   = set(d["data_doc"]       for d in deps_do_grupo)
            else:
                data_ref   = data_cred_str
                dias_grupo = set()
                dias_doc   = set()

            # Dias documentais cujo depósito migrou para o dia seguinte (hora >= corte):
            # data_doc != data_efetiva → o memorando pertence ao dia efetivo, não ao doc.
            # Esses dias NÃO devem ser exigidos no mem_idx pelo data_doc original.
            dias_doc_migrados = {
                d["data_doc"]
                for d in deps_do_grupo
                if d["data_doc"] != d["data_deposito"]
            }

            # Para exibição no modal do dia de fechamento, usa o total do dia documental
            # correspondente à data de referência (não soma todos os dias do grupo)
            data_doc_ref = deps_do_grupo[-1]["data_doc"] if deps_do_grupo else data_ref
            total_dep    = dep_doc_idx.get((data_doc_ref, loja), 0.0)
            total_memo   = float(mem_idx.get((loja, data_doc_ref), 0.0))
            tem_memo      = (loja, data_doc_ref) in mem_idx
            dif_dep_memo  = total_dep - total_memo

            # g1_ok: todos os dias documentais do grupo precisam bater com o memorando,
            # EXCETO os dias cujo depósito migrou de dia (hora >= corte) — esses não
            # têm memorando próprio no data_doc original e não devem penalizar o grupo.
            g1_ok = tem_memo and _iguais(total_dep, total_memo) and all(
                d in dias_doc_migrados  # migrado → isento de verificação
                or (
                    (loja, d) in mem_idx
                    and _iguais(dep_doc_idx.get((d, loja), 0.0), float(mem_idx.get((loja, d), 0.0)))
                )
                for d in dias_doc
            )

            # ─── Status por dia (bloco 1 individual) ──────────────────
            # Cada quadrinho do grupo recebe a cor do seu próprio dep×memo,
            # exceto o último dia que carrega também o resultado do bloco 2.
            # Constrói mapeamento data_efetiva → data_doc para a busca correta
            status_por_dia: dict[str, str] = {}
            data_efetiva_para_doc: dict[str, str] = {
                d["data_deposito"]: d["data_doc"] for d in deps_do_grupo
            }
            for dia_ef in dias_grupo:
                dia_doc      = data_efetiva_para_doc.get(dia_ef, dia_ef)
                dep_dia      = dep_doc_idx.get((dia_doc, loja), 0.0)
                memo_dia     = float(mem_idx.get((loja, dia_doc), 0.0))
                tem_memo_dia = (loja, dia_doc) in mem_idx
                # Dia migrado (hora >= corte): data_doc != data_efetiva.
                # O memorando desse dia existe sob data_efetiva, não data_doc,
                # portanto não penalizamos o quadrinho com "divergente" por ausência.
                dia_migrado = (dia_doc != dia_ef)
                if dia_migrado:
                    # Para dias migrados usamos a data efetiva para buscar o memo
                    memo_ef  = float(mem_idx.get((loja, dia_ef), 0.0))
                    tem_ef   = (loja, dia_ef) in mem_idx
                    dep_ef   = dep_doc_idx.get((dia_ef, loja), 0.0)
                    if not tem_ef:
                        # Memorando do dia efetivo ainda não existe → pendente, não divergente
                        status_por_dia[dia_ef] = "divergente"
                    elif _iguais(dep_ef, memo_ef) and dep_ef > 0:
                        status_por_dia[dia_ef] = "ok"
                    else:
                        status_por_dia[dia_ef] = "divergente"
                elif not tem_memo_dia:
                    status_por_dia[dia_ef] = "divergente"
                elif _iguais(dep_dia, memo_dia) and dep_dia > 0:
                    status_por_dia[dia_ef] = "ok"
                else:
                    status_por_dia[dia_ef] = "divergente"

            # ─── Grupo 2: Depósito Ajust. × Crédito ──────────────────
            total_dep_ajust = acum
            total_cred      = valor_cred
            dif_dep_cred    = total_dep_ajust - total_cred
            g2_ok           = _iguais(total_dep_ajust, total_cred) and total_dep_ajust > 0

            # ─── Status ───────────────────────────────────────────────
            if total_dep_ajust == 0.0:
                status = "sem_credito"   # crédito sem depósito
            elif not tem_memo:
                status = "divergente"
            elif g1_ok and g2_ok:
                status       = "ok"
                dif_dep_memo = 0.0
                dif_dep_cred = 0.0
            else:
                status = "divergente"

            # O dia de referência (último do grupo) carrega o status do grupo inteiro
            status_por_dia[data_ref] = status

            linhas.append({
                # data de referência da linha no painel = data do último dep do grupo
                "data_deposito":         data_ref,
                "data_credito_esperada": data_cred_str,
                "loja":                  loja,

                # todas as datas que compõem o grupo (para colorir todos os quadrinhos)
                "datas_grupo": sorted(dias_grupo) if dias_grupo else [data_ref],

                # status individual por dia — usado para colorir cada quadrinho
                "status_por_dia": status_por_dia,

                # grupo 1
                "total_deposito":  total_dep,
                "memo_remessa":    total_memo,
                "dif_dep_memo":    dif_dep_memo,

                # grupo 2
                "total_deposito_cred": total_dep_ajust,
                "total_credito":       total_cred,
                "dif_dep_cred":        dif_dep_cred,

                # lista dos depósitos que compõem o ajuste
                "depositos_ajust": deps_do_grupo,

                "status":     status,
                "observacao": "",
            })

        # ─── Depósitos que sobraram (não casaram com nenhum crédito) ──
        for (dt_dep, data_dep_str, valor_dep) in fila_deps[dep_cursor:]:

            memo_val   = mem_idx.get((loja, data_dep_str))
            total_memo = float(memo_val or 0.0)
            tem_memo   = memo_val is not None
            dif_dep_memo = valor_dep - total_memo

            linhas.append({
                "data_deposito":         data_dep_str,
                "data_credito_esperada": "",
                "loja":                  loja,

                "total_deposito":  valor_dep,
                "memo_remessa":    total_memo,
                "dif_dep_memo":    dif_dep_memo,

                "total_deposito_cred": valor_dep,
                "total_credito":       0.0,
                "dif_dep_cred":        valor_dep,

                "depositos_ajust": [{
                    "data_deposito": data_dep_str,
                    "valor":         valor_dep,
                }],

                "status":     "sem_credito",
                "observacao": "",
            })

    # =========================================================
    # 4. Aplica ajustes manuais salvos em conciliacao_ajustes
    #    (sobrescreve o que o motor calculou automaticamente)
    # =========================================================

    ajustes = db.get_ajustes("prossegur")  # {(loja, data_deposito): dados}

    for linha in linhas:
        chave = (str(linha["loja"]), linha["data_deposito"])
        aj = ajustes.get(chave)
        if not aj:
            continue
        # Sobrescreve apenas os campos que foram editados manualmente
        if aj.get("status"):
            linha["status"] = aj["status"]
        if aj.get("observacao") is not None:
            linha["observacao"] = aj["observacao"]
        if aj.get("data_credito_esperada"):
            linha["data_credito_esperada"] = aj["data_credito_esperada"]
        if aj.get("total_deposito") is not None:
            linha["total_deposito"] = aj["total_deposito"]
        if aj.get("total_credito") is not None:
            linha["total_credito"] = aj["total_credito"]
        if aj.get("memo_remessa") is not None:
            linha["memo_remessa"] = aj["memo_remessa"]
        if aj.get("dif_dep_memo") is not None:
            linha["dif_dep_memo"] = aj["dif_dep_memo"]
        if aj.get("dif_dep_cred") is not None:
            linha["dif_dep_cred"] = aj["dif_dep_cred"]

    # =========================================================
    # 5. Ordena linhas finais: data_deposito → loja
    # =========================================================

    def _sort_key(l):
        d = _parse(l["data_deposito"]) if l["data_deposito"] != "—" else datetime.max
        n = int(l["loja"]) if str(l["loja"]).isdigit() else 0
        return (d, n)

    linhas.sort(key=_sort_key)

    return linhas

# ──────────────────────────────────────────────────────────────────────────── #
#  Constantes visuais                                                           #
# ──────────────────────────────────────────────────────────────────────────── #

STATUS_LABEL = {
    "ok":           "✔  OK",
    "divergente":   "⚠  Divergente",
    "sem_credito":  "✘  Sem crédito",
    "sem_deposito": "?  Sem depósito",
}
STATUS_COLOR = {
    "ok":           theme.ACCENT_GREEN,
    "divergente":   theme.ACCENT_ORANGE,
    "sem_credito":  theme.ACCENT_RED,
    "sem_deposito": "#9b9b9b",
}

CHIP_W     = 30
CHIP_H     = 20
CHIP_PAD_X = 7
CHIP_PAD_Y = 5
CHIP_VAZIO = theme.SURFACE2
HEADER_X   = 32
HEADER_Y   = 115


# ──────────────────────────────────────────────────────────────────────────── #
#  Janela de detalhe                                                            #
# ──────────────────────────────────────────────────────────────────────────── #
class DepositosWindow(tk.Toplevel):

    def __init__(self, parent, loja, data_ini, data_fim=None):
        super().__init__(parent)

        self.title("Depósitos")
        self.geometry("900x500")

        self.configure(bg=theme.BG_APP)

        tree = ttk.Treeview(
            self,
            columns=("hora", "valor"),
            show="headings"
        )

        tree.heading("hora", text="Hora")
        tree.heading("valor", text="Valor")

        tree.column("hora", width=180)
        tree.column("valor", width=140, anchor="e")

        tree.pack(fill="both", expand=True)

        lancamentos = db.get_resumo_depositos_prossegur_com_hora()

        for l in lancamentos:

            loja_l = str(l["loja"])

            if loja_l != str(loja):
                continue

            dt = _parse(l["data_hora"])

            if not dt:
                continue

            data_l = dt.strftime("%d/%m/%Y")

            if data_fim:
                if not (data_ini <= data_l <= data_fim):
                    continue
            else:
                if data_l != data_ini:
                    continue

            tree.insert(
                "",
                "end",
                values=(
                    dt.strftime("%d/%m/%Y %H:%M:%S"),
                    _fmt_brl(l["valor"])
                )
            )

class DepositosAjustWindow(tk.Toplevel):
    """
    Exibe os depósitos que compõem o grupo 'Depósito ajust.' de uma linha.
    Recebe a lista depositos_ajust: [{data_deposito, valor}, ...]
    """

    def __init__(self, parent, loja, depositos: list[dict]):
        super().__init__(parent)
        self.title(f"Depósitos Ajust. — Loja {loja}")
        self.geometry("520x400")
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)

        header = tk.Frame(self, bg=theme.PRIMARY)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"  Loja {loja}  ·  Depósitos que compõem o grupo",
            bg=theme.PRIMARY, fg="white",
            font=("Segoe UI Semibold", 11), pady=9, anchor="w",
        ).pack(fill="x", padx=12)

        tree = ttk.Treeview(
            self,
            columns=("data", "valor"),
            show="headings",
        )
        tree.heading("data",  text="Data depósito")
        tree.heading("valor", text="Valor")
        tree.column("data",  width=200)
        tree.column("valor", width=180, anchor="e")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        total = 0.0
        for d in depositos:
            val = float(d["valor"])
            total += val
            tree.insert(
                "", "end",
                values=(d["data_deposito"], _fmt_brl(val)),
            )

        # rodapé com total
        footer = tk.Frame(self, bg=theme.CARD_BG, pady=6)
        footer.pack(fill="x")
        tk.Label(
            footer,
            text=f"Total:  {_fmt_brl(total)}",
            bg=theme.CARD_BG, fg=theme.FG_TEXT,
            font=("Segoe UI Semibold", 10),
            anchor="e",
        ).pack(side="right", padx=16)
        tk.Label(
            footer,
            text=f"{len(depositos)} depósito(s)",
            bg=theme.CARD_BG, fg=theme.FG_MUTED,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(side="left", padx=16)


class DetalheCell(tk.Toplevel):
    def __init__(self, parent, linha: dict, data_clicada: str | None = None, on_save=None):
        super().__init__(parent)

        # Se a linha representa um grupo multi-dia, usa a data clicada para
        # mostrar o total documental e o memorando só daquele dia.
        # O bloco Depósito (ajust.) × Crédito sempre exibe o total do grupo.
        data_exibida = data_clicada or linha["data_deposito"]

        # Busca total_deposito e memo do dia clicado diretamente no banco
        if data_clicada and data_clicada != linha.get("data_deposito"):
            loja_str = str(int(str(linha["loja"])))
            lancamentos = db.get_resumo_depositos_prossegur_com_hora()
            total_dep_dia = sum(
                float(l["valor"])
                for l in lancamentos
                if str(l["loja"]) == loja_str
                and l.get("data_deposito") == data_clicada   # usa data documental, igual ao dep_doc_idx
            )
            from core.memorando_service import get_memorandos_por_loja_dia
            dt_click = _parse(data_clicada)
            mem_parcial = get_memorandos_por_loja_dia(
                dt_click.year, dt_click.month, [loja_str]
            ) if dt_click else {}
            memo_dia = float(mem_parcial.get((loja_str, data_clicada), 0.0))
        else:
            total_dep_dia = linha.get("total_deposito", 0.0)
            memo_dia      = linha.get("memo_remessa") or 0.0

        self.title(f"Loja {linha['loja']}  ·  {data_exibida}")
        self.geometry("480x580")
        self.resizable(False, False)
        self.configure(bg=theme.BG_APP)
        theme.apply_theme(self)

        self._dep       = tk.DoubleVar(value=total_dep_dia)
        self._dep_cred  = tk.DoubleVar(value=linha.get("total_deposito_cred", linha.get("total_deposito", 0.0)))
        self._cred      = tk.DoubleVar(value=linha.get("total_credito", 0.0))
        self._memo      = tk.DoubleVar(value=memo_dia)
        self._data_cred = tk.StringVar(value=linha.get("data_credito_esperada", ""))
        self._linha        = linha
        self._data_exibida = data_exibida
        self._on_save      = on_save
        # True quando estamos vendo um dia intermediário do grupo (não o dia de fechamento)
        self._dia_intermediario = (
            data_clicada is not None
            and data_clicada != linha.get("data_deposito")
        )

        self._build(linha, data_exibida)
        self._recalcular()   # recalcula status com dep/memo do dia clicado
        self._center(parent)

    def _fmt(self, v: float) -> str:
        return f"R$ {abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _recalcular(self, *_):
        dep      = self._dep.get()
        dep_cred = self._dep_cred.get()
        cred     = self._cred.get()
        memo     = self._memo.get()

        def _iguais(a, b, tol=0.01):
            return abs(float(a or 0) - float(b or 0)) <= tol

        g1_ok = memo > 0 and _iguais(dep, memo)
        g2_ok = cred > 0 and _iguais(dep_cred, cred)

        # Dia intermediário de grupo: o bloco 2 (ajust. × crédito) representa
        # o grupo inteiro e não deve puxar o status deste dia para divergente.
        # O status aqui reflete apenas se o bloco 1 (dep×memo) deste dia está OK.
        if getattr(self, "_dia_intermediario", False):
            if memo == 0.0:
                status = "divergente"
            elif g1_ok:
                status = "ok"
            else:
                status = "divergente"
        else:
            if cred == 0.0:
                status = "sem_credito"
            elif memo == 0.0:
                status = "divergente"
            elif g1_ok and g2_ok:
                status = "ok"
            else:
                status = "divergente"

        dif_memo = 0.0 if g1_ok else abs(dep - memo)
        dif_cred = 0.0 if g2_ok else abs(dep_cred - cred)

        def _fmt_dif(v):
            return "Zerado" if abs(v) < 0.02 else self._fmt(v)

        self._lbl_dif_memo.configure(text=_fmt_dif(dif_memo))
        self._lbl_dif_cred.configure(text=_fmt_dif(dif_cred))
        self._lbl_status.configure(text=STATUS_LABEL[status], fg=STATUS_COLOR[status])
        self._cur_status   = status
        self._cur_dif_memo = dif_memo
        self._cur_dif_cred = dif_cred

    def _atualizar_credito_por_data(self):
        data = self._data_cred.get().strip()
        loja = str(self._linha["loja"])

        try:
            datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
            return

        creditos = db.get_creditos_prossegur()

        total = 0.0

        for c in creditos:
            data_c = _fmt_date(_parse(str(c["data"])))
            loja_c = str(c["loja"])

            if data_c == data and loja_c == loja:
                total += float(c["debito"])

        self._cred.set(total)
        self._recalcular()    

    def _make_edit_row(self, body, label: str, var: tk.DoubleVar, bg: str):
        row = tk.Frame(body, bg=bg)
        row.pack(fill="x")
        tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=24, pady=7).pack(side="left", padx=10)
        right = tk.Frame(row, bg=bg)
        right.pack(side="right", padx=10)
        lbl = tk.Label(right, text=self._fmt(var.get()),
                       bg=bg, fg=theme.FG_TEXT, font=("Segoe UI Semibold", 10),cursor="hand2")
        lbl.pack(side="left")
        pen = tk.Label(right, text=" ✏", bg=bg, fg=theme.FG_MUTED,
                       font=("Segoe UI", 9), cursor="hand2")
        pen.pack(side="left")

        def _start_edit(e=None):
            lbl.pack_forget()
            pen.pack_forget()
            entry = tk.Entry(right, textvariable=var, width=12,
                             font=("Segoe UI Semibold", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
            entry.pack(side="left")
            entry.focus_set()
            entry.select_range(0, "end")

            def _finish(e=None):
                try:
                    raw = entry.get().strip().replace("R$", "").strip()
                    if "," in raw:
                        raw = raw.replace(".", "").replace(",", ".")
                    var.set(float(raw))
                except ValueError:
                    pass
                lbl.configure(text=self._fmt(var.get()))
                entry.destroy()
                lbl.pack(side="left")
                pen.pack(side="left")
                self._recalcular()

            entry.bind("<Return>", _finish)
            entry.bind("<FocusOut>", _finish)

        pen.bind("<Button-1>", _start_edit)
        lbl.bind("<Double-1>", _start_edit)
        
        def _abrir_depositos(e=None):

            if "depósito" not in label.lower():
                return

            loja = self._linha["loja"]
            data = self._linha["data_deposito"]

            if "ajust" in label.lower():
                # Abre janela mostrando exatamente os depósitos que compõem o grupo
                deps = self._linha.get("depositos_ajust") or []
                DepositosAjustWindow(self, loja, deps)
            else:
                DepositosWindow(self, loja, data)

        lbl.bind("<Button-1>", _abrir_depositos)
        

    def _make_date_edit_row(self, body, label: str, var: tk.StringVar, bg: str):
        row = tk.Frame(body, bg=bg)
        row.pack(fill="x")
        tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=24, pady=7).pack(side="left", padx=10)
        right = tk.Frame(row, bg=bg)
        right.pack(side="right", padx=10)
        lbl = tk.Label(right, text=var.get(),
                       bg=bg, fg=theme.FG_TEXT, font=("Segoe UI Semibold", 10))
        lbl.pack(side="left")
        pen = tk.Label(right, text=" ✏", bg=bg, fg=theme.FG_MUTED,
                       font=("Segoe UI", 9), cursor="hand2")
        pen.pack(side="left")

        def _start_edit(e=None):
            lbl.pack_forget()
            pen.pack_forget()
            entry = tk.Entry(right, textvariable=var, width=12,
                             font=("Segoe UI Semibold", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
            entry.pack(side="left")
            entry.focus_set()
            entry.select_range(0, "end")

            def _finish(e=None):
                txt = entry.get().strip()
                try:
                    datetime.strptime(txt, "%d/%m/%Y")
                    var.set(txt)
                    self._atualizar_credito_por_data()
                except ValueError:
                    pass
                lbl.configure(text=var.get())
                entry.destroy()
                lbl.pack(side="left")
                pen.pack(side="left")

            entry.bind("<Return>", _finish)
            entry.bind("<FocusOut>", _finish)

        pen.bind("<Button-1>", _start_edit)
        lbl.bind("<Double-1>", _start_edit)

    def _build(self, l: dict, data_exibida: str | None = None):
        data_exibida = data_exibida or l["data_deposito"]
        cor = STATUS_COLOR[l["status"]]

        header = tk.Frame(self, bg=cor)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"  Loja {l['loja']}   ·   {data_exibida}",
            bg=cor, fg="white",
            font=("Segoe UI Semibold", 11), pady=10, anchor="w",
        ).pack(fill="x", padx=12)

        body = tk.Frame(self, bg=theme.CARD_BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        def _static_row(label, valor, idx):
            bg = theme.CARD_BG if idx % 2 == 0 else theme.SURFACE2
            row = tk.Frame(body, bg=bg)
            row.pack(fill="x")
            tk.Label(row, text=label, bg=bg, fg=theme.FG_MUTED,
                     font=("Segoe UI", 10), anchor="w",
                     width=24, pady=7).pack(side="left", padx=10)
            tk.Label(row, text=valor, bg=bg, fg=theme.FG_TEXT,
                     font=("Segoe UI Semibold", 10),
                     anchor="e").pack(side="right", padx=10)

        _static_row("Data depósito", data_exibida, 0)
        self._make_date_edit_row(body, "Data crédito esperada", self._data_cred, theme.SURFACE2)

        row_idx = 2

        # ── Grupo 1: Depósito × Memorando ─────────────────────────────────
        sep1_bg = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        sep1 = tk.Frame(body, bg=sep1_bg)
        sep1.pack(fill="x")
        tk.Label(sep1, text="— Depósito × Memorando —",
                 bg=sep1_bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8, "italic"), pady=3).pack(side="left", padx=10)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Total depósito (dia)", self._dep, bg_e)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Memorando (remessa)", self._memo, bg_e)
        row_idx += 1

        bg_d = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_dm = tk.Frame(body, bg=bg_d)
        row_dm.pack(fill="x")
        tk.Label(row_dm, text="Diferença dep×memo", bg=bg_d, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=24, pady=7).pack(side="left", padx=10)
        dif_memo_init = l.get("dif_dep_memo", 0.0) or 0.0
        self._lbl_dif_memo = tk.Label(row_dm,
            text="Zerado" if abs(dif_memo_init) < 0.02 else self._fmt(dif_memo_init),
            bg=bg_d, fg=theme.FG_TEXT, font=("Segoe UI Semibold", 10))
        self._lbl_dif_memo.pack(side="right", padx=10)
        row_idx += 1

        # ── Grupo 2: Depósito (c/ corte 16h) × Crédito ────────────────────
        sep2_bg = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        sep2 = tk.Frame(body, bg=sep2_bg)
        sep2.pack(fill="x")
        tk.Label(sep2, text="— Depósito (corte 16h) × Crédito —",
                 bg=sep2_bg, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8, "italic"), pady=3).pack(side="left", padx=10)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Depósito (ajust. 16h)", self._dep_cred, bg_e)
        row_idx += 1

        bg_e = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        self._make_edit_row(body, "Total crédito", self._cred, bg_e)
        row_idx += 1

        bg_d2 = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_dc = tk.Frame(body, bg=bg_d2)
        row_dc.pack(fill="x")
        tk.Label(row_dc, text="Diferença dep×crédito", bg=bg_d2, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=24, pady=7).pack(side="left", padx=10)
        dif_cred_init = l.get("dif_dep_cred", 0.0) or 0.0
        self._lbl_dif_cred = tk.Label(row_dc,
            text="Zerado" if abs(dif_cred_init) < 0.02 else self._fmt(dif_cred_init),
            bg=bg_d2, fg=theme.FG_TEXT, font=("Segoe UI Semibold", 10))
        self._lbl_dif_cred.pack(side="right", padx=10)
        row_idx += 1

        # ── Status geral ───────────────────────────────────────────────────
        bg_s = theme.CARD_BG if row_idx % 2 == 0 else theme.SURFACE2
        row_st = tk.Frame(body, bg=bg_s)
        row_st.pack(fill="x")
        tk.Label(row_st, text="Status", bg=bg_s, fg=theme.FG_MUTED,
                 font=("Segoe UI", 10), anchor="w",
                 width=24, pady=7).pack(side="left", padx=10)
        self._lbl_status = tk.Label(row_st,
                                     text=STATUS_LABEL[l["status"]],
                                     bg=bg_s, fg=STATUS_COLOR[l["status"]],
                                     font=("Segoe UI Semibold", 10))
        self._lbl_status.pack(side="right", padx=10)
        self._cur_status   = l["status"]
        self._cur_dif_memo = dif_memo_init
        self._cur_dif_cred = dif_cred_init

        # ── Observações ───────────────────────────────────────────────────
        obs_frame = tk.Frame(self, bg=theme.BG_APP)
        obs_frame.pack(fill="x", padx=14, pady=(6, 0))
        tk.Label(obs_frame, text="Observações:", bg=theme.BG_APP,
                 fg=theme.FG_MUTED, font=("Segoe UI", 9)).pack(anchor="w")

        self._obs_var = tk.StringVar(value=l.get("observacao", "") or "")
        obs_entry = tk.Entry(obs_frame, textvariable=self._obs_var,
                             font=("Segoe UI", 10),
                             bg=theme.SURFACE2, fg=theme.FG_TEXT,
                             insertbackground=theme.FG_TEXT,
                             relief="flat", bd=0)
        obs_entry.pack(fill="x", pady=4, ipady=5)

        self._lbl_chars = tk.Label(obs_frame, text="0/150",
                                    bg=theme.BG_APP, fg=theme.FG_MUTED,
                                    font=("Segoe UI", 8))
        self._lbl_chars.pack(anchor="e")

        def _on_obs(*_):
            txt = self._obs_var.get()
            if len(txt) > 150:
                self._obs_var.set(txt[:150])
            self._lbl_chars.configure(text=f"{len(self._obs_var.get())}/150")

        self._obs_var.trace_add("write", _on_obs)
        _on_obs()

        # ── Botão salvar ──────────────────────────────────────────────────
        tk.Button(
            self,
            text="Salvar alterações",
            command=self._salvar,
            bg=theme.PRIMARY, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=8,
            font=("Segoe UI Semibold", 10),
        ).pack(fill="x", padx=14, pady=10)

    def _salvar(self):
        self._linha["total_deposito"]        = self._dep.get()
        self._linha["total_deposito_cred"]   = self._dep_cred.get()
        self._linha["total_credito"]         = self._cred.get()
        self._linha["memo_remessa"]          = self._memo.get()   # ← remove o "or None"
        self._linha["dif_dep_memo"]          = self._cur_dif_memo
        self._linha["dif_dep_cred"]          = self._cur_dif_cred
        self._linha["status"]                = self._cur_status
        self._linha["observacao"]            = self._obs_var.get().strip()
        self._linha["data_credito_esperada"] = self._data_cred.get()

        # ← usa self._data_exibida como chave, não self._linha["data_deposito"]
        db.salvar_ajuste("prossegur", str(self._linha["loja"]), self._data_exibida, {
            "total_deposito":        self._linha["total_deposito"],
            "total_credito":         self._linha["total_credito"],
            "memo_remessa":          self._linha["memo_remessa"],
            "dif_dep_memo":          self._cur_dif_memo,
            "dif_dep_cred":          self._cur_dif_cred,
            "status":                self._cur_status,
            "observacao":            self._linha["observacao"],
            "data_credito_esperada": self._data_cred.get(),
        })

        if self._on_save:
            cb = self._on_save
            self.destroy()
            self.after(50, cb)
        else:
            self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx() + parent.winfo_width() // 2
        ph = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w}x{h}+{pw - w // 2}+{ph - h // 2}")


# ──────────────────────────────────────────────────────────────────────────── #
#  Frame principal                                                              #
# ──────────────────────────────────────────────────────────────────────────── #

class ProssegurPainel(ttk.Frame):

    def __init__(self, master, **kwargs):
        super().__init__(master, style="App.TFrame", **kwargs)
        db.inicializar()
        self._linhas: list[dict] = []
        self._build_shell()
        try:
            self._atualizar()

        except Exception as e:
            import traceback

            print("ERRO PAINEL:")
            print(e)

            traceback.print_exc()
       
    def _build_shell(self):
        toolbar1 = tk.Frame(self, bg=theme.CARD_BG, pady=6)
        toolbar1.pack(fill="x")

        tk.Button(
            toolbar1, text="Atualizar",
            command=self._atualizar,
            bg=theme.PRIMARY, fg="white",
            activebackground=theme.PRIMARY_HOVER, activeforeground="white",
            relief="flat", bd=0, cursor="hand2",
            padx=14, pady=6,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left", padx=10)

        tk.Label(toolbar1, text="Status:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        self._filtro_var = tk.StringVar(value="todos")
        for label, valor in [
            ("Todos", "todos"), ("✔ OK", "ok"),
            ("⚠ Divergente", "divergente"), ("✘ Sem crédito", "sem_credito"),
        ]:
            tk.Radiobutton(
                toolbar1, text=label,
                variable=self._filtro_var, value=valor,
                command=self._aplicar_filtro,
                bg=theme.CARD_BG, fg=theme.FG_TEXT,
                selectcolor=theme.PRIMARY,
                activebackground=theme.CARD_BG,
                font=("Segoe UI", 9), relief="flat", bd=0,
            ).pack(side="left", padx=4)

        self._lbl_resumo = tk.Label(
            toolbar1, text="",
            bg=theme.CARD_BG, fg=theme.FG_MUTED,
            font=("Segoe UI", 9),
        )
        self._lbl_resumo.pack(side="right", padx=16)

        toolbar2 = tk.Frame(self, bg=theme.CARD_BG, pady=4)
        toolbar2.pack(fill="x")

        tk.Label(toolbar2, text="Regional:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        ufs_disponiveis = sorted(set(LOJA_UF.values()))
        self._uf_var = tk.StringVar(value="Todas")
        self._cmb_uf = ttk.Combobox(
            toolbar2, textvariable=self._uf_var,
            values=["Todas"] + ufs_disponiveis,
            state="readonly", width=8, font=("Segoe UI", 9),
        )
        self._cmb_uf.pack(side="left", padx=4)
        self._cmb_uf.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Label(toolbar2, text="Mês:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(16, 4))

        self._mes_var = tk.StringVar(value="Todos")
        self._cmb_mes = ttk.Combobox(
            toolbar2, textvariable=self._mes_var,
            values=["Todos"] + [f"{n:02d} – {MESES_PT[n]}" for n in range(1, 13)],
            state="readonly", width=16, font=("Segoe UI", 9),
        )
        self._cmb_mes.pack(side="left", padx=4)
        self._cmb_mes.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Label(toolbar2, text="Ano:",
                 bg=theme.CARD_BG, fg=theme.FG_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))

        self._ano_var = tk.StringVar(value="Todos")
        self._cmb_ano = ttk.Combobox(
            toolbar2, textvariable=self._ano_var,
            values=["Todos"], state="readonly", width=8, font=("Segoe UI", 9),
        )
        self._cmb_ano.pack(side="left", padx=4)
        self._cmb_ano.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())

        tk.Frame(self, bg=theme.GRID_LINE, height=1).pack(fill="x")

        outer = ttk.Frame(self, style="App.TFrame")
        outer.pack(fill="both", expand=True)

        self._vsb = ttk.Scrollbar(outer, orient="vertical", style="Vertical.TScrollbar")
        self._hsb = ttk.Scrollbar(outer, orient="horizontal")
        self._canvas = tk.Canvas(
            outer, bg=theme.BG_APP, highlightthickness=0,
            yscrollcommand=self._vsb.set,
            xscrollcommand=self._hsb.set,
        )
        self._vsb.configure(command=self._canvas.yview)
        self._hsb.configure(command=self._canvas.xview)
        self._vsb.pack(side="right", fill="y")
        self._hsb.pack(side="bottom", fill="x")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self._canvas.bind("<Shift-MouseWheel>",
            lambda e: self._canvas.xview_scroll(-1 * (e.delta // 120), "units"))

        leg = tk.Frame(self, bg=theme.BG_APP, pady=5)
        leg.pack(fill="x", side="bottom")
        for status, cor in STATUS_COLOR.items():
            tk.Label(leg, text="●", fg=cor, bg=theme.BG_APP,
                     font=("Segoe UI", 11)).pack(side="left", padx=(12, 2))
            tk.Label(leg, text=STATUS_LABEL[status],
                     fg=theme.FG_MUTED, bg=theme.BG_APP,
                     font=("Segoe UI", 9)).pack(side="left", padx=(0, 10))

        tk.Label(leg, text="Clique em um quadrado para ver o detalhe",
                 bg=theme.BG_APP, fg=theme.FG_MUTED,
                 font=("Segoe UI", 8)).pack(side="right", padx=12)

    def _on_linha_salva(self, linha_atualizada: dict | None = None):
        self._atualizar()

    def _atualizar(self):
        self._linhas = conciliar()
        self._popular_anos()
        self._aplicar_filtro()

    def _popular_anos(self):
        anos = set()
        for l in self._linhas:
            if l["data_deposito"] == "—":
                continue
            dt = _parse(l["data_deposito"])
            if dt:
                anos.add(str(dt.year))
        opcoes = ["Todos"] + sorted(anos, reverse=True)
        self._cmb_ano.configure(values=opcoes)
        if self._ano_var.get() not in opcoes:
            self._ano_var.set("Todos")

    def _aplicar_filtro(self):

        f       = self._filtro_var.get()
        uf      = self._uf_var.get()
        mes_sel = self._mes_var.get()
        ano_sel = self._ano_var.get()

        mes_num = None

        if mes_sel.lower() != "todos":

            try:
                mes_num = int(mes_sel.split("–")[0].strip())
            except:
                mes_num = None

        ano_num = None

        if ano_sel != "Todos":

            try:
                ano_num = int(ano_sel)
            except:
                ano_num = None

        exibir = []


        for l in self._linhas:

            try:

                # STATUS
                if f != "todos" and l["status"] != f:
                    continue

                # UF
                if uf != "Todas":

                    loja_uf = LOJA_UF.get(str(l["loja"]), "")

                    if loja_uf != uf:
                        continue

                # DATA
                if mes_num is not None or ano_num is not None:

                    data_dep = l.get("data_deposito")

                    if not data_dep or data_dep == "—":
                        continue

                    dt = _parse(data_dep)

                    if not dt:
                        continue

                    if mes_num is not None and dt.month != mes_num:
                        continue

                    if ano_num is not None and dt.year != ano_num:
                        continue

                exibir.append(l)

            except Exception as e:

                print("ERRO FILTRO:", e)
                print(l)

        datas_extras = None

        if mes_num is None:

            hoje = datetime.now()

            mes_num = hoje.month
            ano_num = hoje.year

        datas_extras = self._gerar_datas_mes(mes_num, ano_num)

        self._desenhar_grade(exibir, datas_extras=datas_extras)

        # resumo usando EXIBIR
        total = len(exibir)

        ok  = sum(1 for l in exibir if l["status"] == "ok")
        div = sum(1 for l in exibir if l["status"] == "divergente")
        sc  = sum(1 for l in exibir if l["status"] == "sem_credito")

        self._lbl_resumo.configure(
            text=f"{total} registros  ·  ✔ {ok}  ⚠ {div}  ✘ {sc}"
        )

    def _gerar_datas_mes(self, mes: int, ano: int | None) -> list[str]:
        if ano is None:
            ano = datetime.now().year
        _, num_dias = calendar.monthrange(ano, mes)
        return [f"{d:02d}/{mes:02d}/{ano}" for d in range(1, num_dias + 1)]

    def _atualizar_resumo(self):
        total = len(self._linhas)
        ok  = sum(1 for l in self._linhas if l["status"] == "ok")
        div = sum(1 for l in self._linhas if l["status"] == "divergente")
        sc  = sum(1 for l in self._linhas if l["status"] == "sem_credito")
        self._lbl_resumo.configure(
            text=f"{total} registros  ·  ✔ {ok}  ⚠ {div}  ✘ {sc}"
        )

    def _desenhar_grade(self, linhas: list[dict], datas_extras: list[str] | None = None):
        c = self._canvas
        c.delete("all")

        if not linhas and not datas_extras:
            c.create_text(200, 80, text="Nenhum dado para exibir.",
                          fill=theme.FG_MUTED, font=("Segoe UI", 11))
            c.configure(scrollregion=(0, 0, 400, 160))
            return

        def _sort_date(s):
            try:    return datetime.strptime(s, "%d/%m/%Y")
            except: return datetime.max

        datas_dados = {l["data_deposito"] for l in linhas if l["data_deposito"] != "—"}
        if datas_extras:
            datas_todas = sorted(datas_dados | set(datas_extras), key=_sort_date)
        else:
            datas_todas = sorted(datas_dados, key=_sort_date)

        lojas = sorted({l["loja"] for l in linhas},
                       key=lambda x: int(x) if str(x).isdigit() else x)

        if not datas_todas or not lojas:
            c.create_text(200, 80, text="Nenhum dado para exibir.",
                          fill=theme.FG_MUTED, font=("Segoe UI", 11))
            c.configure(scrollregion=(0, 0, 400, 160))
            return

        idx = {}

        for l in linhas:

            datas = l.get("datas_grupo") or [l["data_deposito"]]

            for d in datas:

                idx[(d, l["loja"])] = l

        ROW_H   = CHIP_H + CHIP_PAD_Y * 2
        COL_W   = CHIP_W + CHIP_PAD_X * 2
        total_w = HEADER_Y + len(datas_todas) * COL_W + 20
        total_h = HEADER_X + len(lojas) * ROW_H + 20

        for j, data in enumerate(datas_todas):
            x = HEADER_Y + j * COL_W + COL_W // 2
            c.create_text(x, HEADER_X // 2, text=data[:5],
                          fill=theme.FG_MUTED, font=("Segoe UI", 8))

        c.create_line(0, HEADER_X, total_w, HEADER_X, fill=theme.GRID_LINE, width=1)

        for i, loja in enumerate(lojas):
            y_top  = HEADER_X + i * ROW_H
            bg_row = theme.CARD_BG if i % 2 == 0 else theme.SURFACE2

            c.create_rectangle(0, y_top, total_w, y_top + ROW_H, fill=bg_row, outline="")

            uf_loja    = LOJA_UF.get(str(loja), "")
            label_loja = f"L{loja}" + (f" ({uf_loja})" if uf_loja else "")
            c.create_text(6, y_top + ROW_H // 2, text=label_loja,
                          fill=theme.FG_MUTED, font=("Segoe UI", 9), anchor="w")

            for j, data in enumerate(datas_todas):
                linha = idx.get((data, loja))
                x0 = HEADER_Y + j * COL_W + CHIP_PAD_X
                y0 = y_top + CHIP_PAD_Y
                x1 = x0 + CHIP_W
                y1 = y0 + CHIP_H

                if linha is None:
                    c.create_rectangle(x0, y0, x1, y1, fill=CHIP_VAZIO, outline=theme.GRID_LINE)
                else:
                    # usa o status individual do dia se disponível,
                    # senão cai no status geral do grupo
                    status_dia = linha.get("status_por_dia", {}).get(data, linha["status"])
                    cor   = STATUS_COLOR[status_dia]
                    hover = self._darken(cor)
                    rid   = c.create_rectangle(x0, y0, x1, y1, fill=cor, outline=cor, tags=("chip",))

                    def _enter(e, r=rid, h=hover): c.itemconfig(r, fill=h, outline=h)
                    def _leave(e, r=rid, co=cor):  c.itemconfig(r, fill=co, outline=co)
                    def _click(e, l=linha, d=data): DetalheCell(self.winfo_toplevel(), l, data_clicada=d, on_save=self._on_linha_salva)

                    c.tag_bind(rid, "<Enter>",    _enter)
                    c.tag_bind(rid, "<Leave>",    _leave)
                    c.tag_bind(rid, "<Button-1>", _click)
                    c.tag_bind(rid, "<ButtonRelease-1>",
                               lambda e, r=rid, co=cor: c.itemconfig(r, fill=co, outline=co))

        c.configure(scrollregion=(0, 0, total_w, total_h))
        c.configure(cursor="hand2")

    @staticmethod
    def _darken(hex_color: str) -> str:
        h = hex_color.lstrip("#")
        r, g, b = (int(h[i:i+2], 16) for i in (0, 2, 4))
        return f"#{max(0,r-35):02x}{max(0,g-35):02x}{max(0,b-35):02x}"