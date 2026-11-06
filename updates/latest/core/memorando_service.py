"""
memorando_service.py
Serviço compartilhado para buscar valor_remessa_dinheiro da tabela
'memorandos' no Supabase, convertendo centavos → reais.

Usado por brinks/painel.py e prossegur/painel.py.
"""

import calendar
from datetime import datetime
from core.supabase_config import supabase


def get_memorandos_por_loja_dia(
    ano: int,
    mes_num: int,
    lojas: list[str] | None = None,
) -> dict[tuple[str, str], float]:

    resultado: dict[tuple[str, str], float] = {}
    try:
        last_day = calendar.monthrange(ano, mes_num)[1]
        ini = f"{ano}-{mes_num:02d}-01"
        fim = f"{ano}-{mes_num:02d}-{last_day:02d}"

        PAGE = 1000
        offset = 0
        rows = []

        while True:
            q = (
                supabase.table("memorandos")
                .select("data, loja, valor_remessa_dinheiro")
                .gte("data", ini)
                .lte("data", fim)
                .range(offset, offset + PAGE - 1)
            )
            if lojas:
                q = q.in_("loja", [str(l) for l in lojas])

            pagina = q.execute().data or []
            rows.extend(pagina)

            if len(pagina) < PAGE:
                break
            offset += PAGE

        print(f"[memorando_service] {ano}-{mes_num:02d}: {len(rows)} rows total")

        for r in rows:
            try:
                loja = str(int(str(r["loja"]).strip()))
                dt   = datetime.strptime(r["data"], "%Y-%m-%d")
                raw  = r.get("valor_remessa_dinheiro") or 0

                if isinstance(raw, str):
                    raw = raw.strip()
                    if "," in raw:
                        raw = raw.replace(".", "").replace(",", ".")

                val = float(raw)
                data_fmt = dt.strftime("%d/%m/%Y")
                key = (loja, data_fmt)
                resultado[key] = resultado.get(key, 0.0) + val
            except Exception:
                continue

    except Exception as e:
        print("[memorando_service] Erro ao buscar memorandos:", e)

    return resultado