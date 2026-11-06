import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import re
import os
import json
import pickle
from datetime import date, datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════
# PALETA E ESTILOS
# ══════════════════════════════════════════════════════════════

from core import theme as _t     # identidade visual do Prisma (core/theme.py)

COR = {
    "bg":           _t.BG_APP,
    "surface":      _t.CARD_BG,
    "border":       _t.GRID_LINE,
    "text":         _t.FG_TEXT,
    "muted":        _t.FG_MUTED,
    "azul":         "#3d63d6",       # cabeçalho/botão de débitos (texto branco por cima)
    "azul_light":   "#1a2440",
    "azul_dark":    "#9db6ff",
    "verde":        "#00a862",       # cabeçalho/botão de créditos (texto branco por cima)
    "verde_light":  "#10281f",
    "verde_dark":   _t.ACCENT_GREEN,
    "amber":        _t.ACCENT_ORANGE,
    "amber_light":  "#2b2410",
    "vermelho":     "#d63a56",
    "vermelho_lgt": "#2c1520",
    "deb_sel":      "#27407a",
    "cred_sel":     "#1c5a3d",
    "conciliado":   "#10281f",
    "detalhe_bg":   _t.CARD_BG,
    "detalhe_txt":  _t.FG_TEXT,
}

FONTE          = ("Helvetica", 10)
FONTE_BOLD     = ("Helvetica", 10, "bold")
FONTE_HUD      = ("Helvetica", 19, "bold")
FONTE_HUD_LBL  = ("Helvetica", 8)
FONTE_DET      = ("Helvetica", 9)
FONTE_DET_BOLD = ("Helvetica", 9, "bold")


# ══════════════════════════════════════════════════════════════
# PERSISTÊNCIA DE PERFIS (SALVAR/CARREGAR PROGRESSO POR EMPRESA)
# ══════════════════════════════════════════════════════════════

PERFIS_DIR = os.path.join(os.path.expanduser("~"), ".conciliador_989", "perfis")
os.makedirs(PERFIS_DIR, exist_ok=True)
INDEX_PATH = os.path.join(PERFIS_DIR, "_index.json")

CONTAS = ["989", "1677"]   # contas suportadas pelo seletor (expansível)


def slug_perfil(nome: str) -> str:
    """Converte o nome da empresa em um identificador de arquivo seguro."""
    s = re.sub(r'[^\w\-]+', '_', (nome or '').strip(), flags=re.UNICODE)
    s = s.strip('_').upper()
    return s or "PERFIL"


def caminho_perfil(slug: str, conta: str) -> str:
    """Perfis são armazenados em subpastas por conta, evitando colisão
    entre uma mesma empresa com progresso separado em contas diferentes."""
    pasta = os.path.join(PERFIS_DIR, conta)
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, f"{slug}.pkl")


def carregar_index() -> list:
    if not os.path.exists(INDEX_PATH):
        return []
    try:
        with open(INDEX_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def salvar_index(lista: list) -> None:
    with open(INDEX_PATH, 'w', encoding='utf-8') as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)


def upsert_index(entry: dict) -> None:
    """Chave composta (slug, conta) — a mesma empresa pode ter um perfil
    distinto em cada conta."""
    lista = [e for e in carregar_index()
             if not (e['slug'] == entry['slug'] and e.get('conta') == entry.get('conta'))]
    lista.append(entry)
    salvar_index(lista)


def remover_index(slug: str, conta: str) -> None:
    salvar_index([e for e in carregar_index()
                  if not (e['slug'] == slug and e.get('conta') == conta)])


# ══════════════════════════════════════════════════════════════
# MOTOR DE CONCILIAÇÃO
# ══════════════════════════════════════════════════════════════

def extract_lj(historico: str):
    m = re.search(r'LJ\s*0*(\d+)', str(historico), re.IGNORECASE)
    return m.group(1).lstrip('0') if m else None


def extract_lj_partida(conta_partida):
    """Extrai o número da loja do débito a partir de CONTA_PARTIDA (formato 'XX - YYY')."""
    m = re.match(r'^\s*(\d+)\s*[-–]', str(conta_partida))
    return m.group(1).lstrip('0') if m else None


def _normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nomes de colunas: strip, upper, espaços → underscore."""
    df.columns = [c.strip().upper().replace(' ', '_') for c in df.columns]
    return df


def load_and_prepare(path: str):
    """Lê planilha no novo formato (colunas DEBITO e CREDITO separadas)
    ou nos formatos legados (coluna TIPO ou duas abas)."""
    xl  = pd.ExcelFile(path)

    # ── Detecta formato ────────────────────────────────────────────────────
    if 'DEBITO' in xl.sheet_names and 'CREDITO' in xl.sheet_names:
        # Legado: duas abas separadas
        deb  = _normalizar_colunas(pd.read_excel(xl, sheet_name='DEBITO',  header=0))
        cred = _normalizar_colunas(pd.read_excel(xl, sheet_name='CREDITO', header=0))
        # Garante coluna de valor unificada
        for df in (deb, cred):
            if 'VALOR' in df.columns and 'VALOR_ORIG' not in df.columns:
                df.rename(columns={'VALOR': 'VALOR_ORIG'}, inplace=True)
    else:
        # Aba única — novo padrão ou legado com coluna TIPO
        raw = _normalizar_colunas(pd.read_excel(xl, header=0))

        if 'DEBITO' in raw.columns and 'CREDITO' in raw.columns:
            # ── Novo padrão: colunas DEBITO e CREDITO separadas ────────────
            # Linhas com valor em DEBITO → DataFrame de débitos
            deb_mask  = raw['DEBITO'].notna()  & (raw['DEBITO']  != 0)
            cred_mask = raw['CREDITO'].notna() & (raw['CREDITO'] != 0)

            deb  = raw[deb_mask].copy().reset_index(drop=True)
            cred = raw[cred_mask].copy().reset_index(drop=True)

            deb['VALOR_ORIG']  = deb['DEBITO'].abs().round(2)
            cred['VALOR_ORIG'] = cred['CREDITO'].abs().round(2)

        elif 'TIPO' in raw.columns:
            # Legado: aba única com coluna TIPO
            if 'VALOR' in raw.columns and 'VALOR_ORIG' not in raw.columns:
                raw.rename(columns={'VALOR': 'VALOR_ORIG'}, inplace=True)
            deb  = raw[raw['TIPO'].str.upper() == 'DEBITO'].copy().reset_index(drop=True)
            cred = raw[raw['TIPO'].str.upper() == 'CREDITO'].copy().reset_index(drop=True)
        else:
            raise ValueError(
                "Formato de planilha não reconhecido. "
                "Esperado: colunas DEBITO/CREDITO, coluna TIPO, ou abas DEBITO/CREDITO.")

    # ── Limpeza comum ──────────────────────────────────────────────────────
    for df in (deb, cred):
        df.dropna(subset=['VALOR_ORIG', 'DATA'], inplace=True)
        df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')
        df.dropna(subset=['DATA'], inplace=True)
        df['VALOR_ORIG'] = df['VALOR_ORIG'].abs().round(2)
        df.reset_index(drop=True, inplace=True)

    deb['_ID']      = ['D' + str(i) for i in deb.index]
    cred['_ID']     = ['C' + str(i) for i in cred.index]

    # Loja no crédito: extraída do HISTORICO (LJ XX)
    cred['LJ_NUM']  = cred['HISTORICO'].apply(extract_lj)

    # Loja no débito: FILIAL (novo modelo) ou CONTA_PARTIDA (legado)
    if 'FILIAL' in deb.columns:
        deb['LJ_DEB'] = deb['FILIAL'].apply(
            lambda v: str(int(v)).lstrip('0') if pd.notna(v) else None)
    elif 'CONTA_PARTIDA' in deb.columns:
        deb['LJ_DEB'] = deb['CONTA_PARTIDA'].apply(extract_lj_partida)
    else:
        deb['LJ_DEB'] = None

    deb['_STATUS']  = 'PENDENTE'
    cred['_STATUS'] = 'PENDENTE'
    deb['_GRUPO']   = None
    cred['_GRUPO']  = None
    return deb, cred


def _lojas_compativeis(lj_deb, lj_cred):
    """True se ambas as lojas forem None/desconhecidas OU iguais."""
    if lj_deb is None or lj_cred is None:
        return True   # sem informação → não restringe
    return str(lj_deb).strip() == str(lj_cred).strip()


MARGEM_ERRO = 1.00   # tolerância de ± R$ 1,00 na conciliação automática e manual


def auto_conciliar(deb: pd.DataFrame, cred: pd.DataFrame, janela_dias: int = 2,
                   ignorar_loja: bool = False):
    grupos = []
    gid_counter = [0]

    def novo_grupo(ids_deb, ids_cred, metodo):
        gid_counter[0] += 1
        gid = f"G{gid_counter[0]:05d}"
        for iid in ids_deb:
            i = deb.index[deb['_ID'] == iid][0]
            deb.at[i, '_STATUS'] = 'CONCILIADO'
            deb.at[i, '_GRUPO']  = gid
        for iid in ids_cred:
            i = cred.index[cred['_ID'] == iid][0]
            cred.at[i, '_STATUS'] = 'CONCILIADO'
            cred.at[i, '_GRUPO']  = gid
        grupos.append({'gid': gid, 'deb': ids_deb, 'cred': ids_cred, 'metodo': metodo})

    # ── Rodada 1 — 1 Débito × 1 Crédito (mesma loja, mesma data ±janela) ──
    # Usa lista plana para suportar margem de ± MARGEM_ERRO
    cred_pend1 = [dict(row) for _, row in cred[cred['_STATUS'] == 'PENDENTE'].iterrows()]

    for _, drow in deb[deb['_STATUS'] == 'PENDENTE'].iterrows():
        val      = drow['VALOR_ORIG']
        data_deb = drow['DATA'].date()
        lj_deb   = drow.get('LJ_DEB')

        melhor, menor_delta, menor_diff = None, janela_dias + 1, MARGEM_ERRO + 1
        for crow in cred_pend1:
            if crow['_STATUS'] == 'CONCILIADO':
                continue
            diff = abs(crow['VALOR_ORIG'] - val)
            if diff > MARGEM_ERRO:
                continue
            if not ignorar_loja and not _lojas_compativeis(lj_deb, crow.get('LJ_NUM')):
                continue
            delta = abs((pd.Timestamp(crow['DATA']).date() - data_deb).days)
            if delta > janela_dias:
                continue
            # Prefere menor diferença de valor; empata por menor delta de data
            if diff < menor_diff or (diff == menor_diff and delta < menor_delta):
                melhor, menor_delta, menor_diff = crow, delta, diff

        if melhor:
            i = cred.index[cred['_ID'] == melhor['_ID']][0]
            cred.at[i, '_STATUS'] = 'CONCILIADO'
            for c in cred_pend1:
                if c['_ID'] == melhor['_ID']:
                    c['_STATUS'] = 'CONCILIADO'
            novo_grupo([drow['_ID']], [melhor['_ID']], 'AUTO_1x1')

    # ── Rodada 2 — 1 Débito × N Créditos (mesma loja) ──────────────────
    cred_pend2 = cred[cred['_STATUS'] == 'PENDENTE'].copy()
    for _, drow in deb[deb['_STATUS'] == 'PENDENTE'].sort_values(
            'VALOR_ORIG', ascending=False).iterrows():
        val_alvo = drow['VALOR_ORIG']
        data_deb = drow['DATA'].date()
        lj_deb   = drow.get('LJ_DEB')

        candidatos = cred_pend2[
            (cred_pend2['_STATUS'] == 'PENDENTE') &
            (cred_pend2['DATA'].apply(
                lambda d: abs((d.date() - data_deb).days)) <= janela_dias) &
            (cred_pend2['VALOR_ORIG'] <= val_alvo + MARGEM_ERRO) &
            (cred_pend2['LJ_NUM'].apply(lambda lj: True if ignorar_loja else _lojas_compativeis(lj_deb, lj)))
        ].sort_values('VALOR_ORIG', ascending=False)

        acumulado, sel = 0.0, []
        for _, crow in candidatos.iterrows():
            if round(acumulado + crow['VALOR_ORIG'], 2) <= val_alvo + MARGEM_ERRO:
                acumulado = round(acumulado + crow['VALOR_ORIG'], 2)
                sel.append(crow['_ID'])
                if abs(acumulado - val_alvo) <= MARGEM_ERRO:
                    break

        if abs(round(acumulado, 2) - val_alvo) <= MARGEM_ERRO and sel:
            for iid in sel:
                idx = cred_pend2.index[cred_pend2['_ID'] == iid][0]
                cred_pend2.at[idx, '_STATUS'] = 'CONCILIADO'
            novo_grupo([drow['_ID']], sel, 'AUTO_Nx1')

    # ── Rodada 3 — M Débitos × N Créditos (subconjuntos, mesma loja) ──────
    # Busca subconjuntos de débitos e créditos cujas somas se encaixam dentro
    # da margem de erro. Limita o tamanho dos subconjuntos para manter
    # performance (MAX_SUBSET itens por lado).
    MAX_SUBSET = 6   # evita explosão combinatória em lotes grandes

    def _subsets(items, max_size):
        """Gera todos os subconjuntos não-vazios de tamanho ≤ max_size."""
        from itertools import combinations
        for size in range(1, min(max_size, len(items)) + 1):
            yield from combinations(items, size)

    deb_pend3  = deb[deb['_STATUS'] == 'PENDENTE'].copy()
    cred_pend3 = cred[cred['_STATUS'] == 'PENDENTE'].copy()

    if ignorar_loja:
        # Sem agrupamento por loja: trata tudo como um único "grupo universal"
        lojas_iter = [('__ALL__', slice(None), slice(None))]
    else:
        lojas = deb_pend3['LJ_DEB'].dropna().unique().tolist()
        lojas_sem = [None] if deb_pend3['LJ_DEB'].isna().any() else []
        lojas_iter = [(loja, None, None) for loja in list(lojas) + lojas_sem]

    for loja_entry in lojas_iter:
        loja = loja_entry[0]
        encontrou = True
        while encontrou:           # repete enquanto achar grupos nesta loja
            encontrou = False

            if loja == '__ALL__':
                debs_loja  = deb_pend3[deb_pend3['_STATUS'] == 'PENDENTE']
                creds_loja = cred_pend3[cred_pend3['_STATUS'] == 'PENDENTE']
            elif loja is not None:
                mask_d = deb_pend3['LJ_DEB'] == loja
                mask_c = cred_pend3['LJ_NUM'] == loja
                debs_loja  = deb_pend3[mask_d & (deb_pend3['_STATUS'] == 'PENDENTE')]
                creds_loja = cred_pend3[mask_c & (cred_pend3['_STATUS'] == 'PENDENTE')]
            else:
                mask_d = deb_pend3['LJ_DEB'].isna()
                mask_c = cred_pend3['LJ_NUM'].isna()
                debs_loja  = deb_pend3[mask_d & (deb_pend3['_STATUS'] == 'PENDENTE')]
                creds_loja = cred_pend3[mask_c & (cred_pend3['_STATUS'] == 'PENDENTE')]

            if debs_loja.empty or creds_loja.empty:
                break

            # Âncora: itera por cada débito pendente e define a janela de datas
            for _, drow in debs_loja.sort_values('DATA').iterrows():
                if drow['_STATUS'] == 'CONCILIADO':
                    continue
                data_ref = drow['DATA'].date()

                # Candidatos dentro da janela de dias
                cands_d = debs_loja[
                    debs_loja['DATA'].apply(
                        lambda d: abs((d.date() - data_ref).days)) <= janela_dias
                ][['_ID', 'VALOR_ORIG']].to_dict('records')

                cands_c = creds_loja[
                    creds_loja['DATA'].apply(
                        lambda d: abs((d.date() - data_ref).days)) <= janela_dias
                ][['_ID', 'VALOR_ORIG']].to_dict('records')

                if not cands_d or not cands_c:
                    continue

                melhor = None          # (ids_d, ids_c, diferença)
                melhor_diff = MARGEM_ERRO + 1

                # Busca subconjunto de débitos × subconjunto de créditos
                for sub_d in _subsets(cands_d, MAX_SUBSET):
                    soma_d = round(sum(r['VALOR_ORIG'] for r in sub_d), 2)
                    for sub_c in _subsets(cands_c, MAX_SUBSET):
                        soma_c = round(sum(r['VALOR_ORIG'] for r in sub_c), 2)
                        diff = abs(soma_d - soma_c)
                        if diff <= MARGEM_ERRO and diff < melhor_diff:
                            melhor = (
                                [r['_ID'] for r in sub_d],
                                [r['_ID'] for r in sub_c],
                                diff,
                            )
                            melhor_diff = diff
                            if diff == 0.0:
                                break   # perfeito, não precisa continuar
                    if melhor and melhor_diff == 0.0:
                        break

                if melhor:
                    ids_d, ids_c, _ = melhor
                    deb_pend3.loc[deb_pend3['_ID'].isin(ids_d), '_STATUS']  = 'CONCILIADO'
                    cred_pend3.loc[cred_pend3['_ID'].isin(ids_c), '_STATUS'] = 'CONCILIADO'
                    novo_grupo(ids_d, ids_c, 'AUTO_MxN')
                    encontrou = True
                    break   # reinicia o while para esta loja com listas atualizadas

    return grupos


# ══════════════════════════════════════════════════════════════
# EXPORTAÇÃO XLSX
# ══════════════════════════════════════════════════════════════

THIN  = Side(style='thin', color='CCCCCC')
BORDA = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _hdr(ws, row, cols, fill_hex):
    fill = PatternFill("solid", fgColor=fill_hex)
    for c, t in enumerate(cols, 1):
        cell = ws.cell(row=row, column=c, value=t)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = BORDA


def _drow(ws, row, values, fill_hex=None):
    fill = PatternFill("solid", fgColor=fill_hex) if fill_hex else None
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.border = BORDA
        cell.alignment = Alignment(vertical='center')
        if fill:
            cell.fill = fill
        if isinstance(v, float):
            cell.number_format = '#,##0.00'


def _autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 4, 50)


def exportar_xlsx(deb, cred, grupos, dest_path):
    wb = openpyxl.Workbook()

    def _get(row, *keys, default=''):
        """Tenta várias chaves (para compatibilidade legado/novo modelo)."""
        for k in keys:
            v = row.get(k)
            if v is not None and str(v).strip() not in ('', 'nan', 'None'):
                return v
        return default

    # CONCILIADOS
    ws1 = wb.active
    ws1.title = "CONCILIADOS"
    _hdr(ws1, 1, ['GRUPO', 'MÉTODO', 'TIPO', 'DATA', 'FILIAL', 'DOC/ID',
                  'HISTÓRICO', 'RAZÃO SOCIAL', 'CONTA_PARTIDA / LJ', 'VALOR (R$)'], '185FA5')
    row = 2
    for g in grupos:
        for iid in g['deb']:
            r = deb[deb['_ID'] == iid].iloc[0]
            lj_info = str(_get(r, 'CONTA_PARTIDA') or '')
            _drow(ws1, row, [
                g['gid'], g['metodo'], 'DÉBITO',
                r['DATA'].strftime('%d/%m/%Y'),
                str(_get(r, 'FILIAL')),
                str(_get(r, 'DOC_RNO')),
                str(_get(r, 'HISTORICO')),
                str(_get(r, 'RAZAO_SOCIAL')),
                lj_info,
                r['VALOR_ORIG'],
            ], 'DDEEFF')
            row += 1
        for iid in g['cred']:
            r = cred[cred['_ID'] == iid].iloc[0]
            lj_info = f"LJ {r.get('LJ_NUM', '')}" if r.get('LJ_NUM') else ''
            _drow(ws1, row, [
                g['gid'], g['metodo'], 'CRÉDITO',
                r['DATA'].strftime('%d/%m/%Y'),
                str(_get(r, 'FILIAL')),
                str(_get(r, 'DOC_RNO')),
                str(_get(r, 'HISTORICO')),
                str(_get(r, 'RAZAO_SOCIAL')),
                lj_info,
                r['VALOR_ORIG'],
            ], 'DDEEDD')
            row += 1
    _autofit(ws1)

    # DÉBITOS PENDENTES
    ws2 = wb.create_sheet("DEBITOS_PENDENTES")
    _hdr(ws2, 1, ['DATA', 'FILIAL', 'DOC_RNO', 'HISTÓRICO', 'RAZÃO SOCIAL',
                  'CONTA_PARTIDA', 'LOJA', 'VOUCHER', 'VALOR (R$)'], 'BA7517')
    for r, (_, rec) in enumerate(deb[deb['_STATUS'] == 'PENDENTE'].iterrows(), 2):
        _drow(ws2, r, [
            rec['DATA'].strftime('%d/%m/%Y'),
            str(_get(rec, 'FILIAL')),
            str(_get(rec, 'DOC_RNO')),
            str(_get(rec, 'HISTORICO')),
            str(_get(rec, 'RAZAO_SOCIAL')),
            str(_get(rec, 'CONTA_PARTIDA')),
            str(rec.get('LJ_DEB', '') or ''),
            str(_get(rec, 'VOUCHER')),
            rec['VALOR_ORIG'],
        ], 'FFF8EE')
    _autofit(ws2)

    # CRÉDITOS PENDENTES
    ws3 = wb.create_sheet("CREDITOS_PENDENTES")
    _hdr(ws3, 1, ['DATA', 'FILIAL', 'DOC_RNO', 'HISTÓRICO', 'RAZÃO SOCIAL',
                  'COD_FORNECEDOR', 'LJ', 'VALOR (R$)'], '3B6D11')
    for r, (_, rec) in enumerate(cred[cred['_STATUS'] == 'PENDENTE'].iterrows(), 2):
        _drow(ws3, r, [
            rec['DATA'].strftime('%d/%m/%Y'),
            str(_get(rec, 'FILIAL')),
            str(_get(rec, 'DOC_RNO')),
            str(_get(rec, 'HISTORICO')),
            str(_get(rec, 'RAZAO_SOCIAL')),
            str(_get(rec, 'COD_FORNECEDOR')),
            str(rec.get('LJ_NUM', '') or ''),
            rec['VALOR_ORIG'],
        ], 'F0FBF0')
    _autofit(ws3)

    # RESUMO
    ws4 = wb.create_sheet("RESUMO")
    dc = deb[deb['_STATUS'] == 'CONCILIADO']
    cc = cred[cred['_STATUS'] == 'CONCILIADO']
    resumo = [
        ('TOTAL DÉBITOS (qtd)',       len(deb)),
        ('TOTAL CRÉDITOS (qtd)',      len(cred)),
        ('DÉBITOS CONCILIADOS',       len(dc)),
        ('CRÉDITOS CONCILIADOS',      len(cc)),
        ('DÉBITOS PENDENTES',         len(deb) - len(dc)),
        ('CRÉDITOS PENDENTES',        len(cred) - len(cc)),
        ('GRUPOS DE CONCILIAÇÃO',     len(grupos)),
        ('', ''),
        ('VALOR DÉBITO TOTAL (R$)',   deb['VALOR_ORIG'].sum()),
        ('VALOR CRÉDITO TOTAL (R$)',  cred['VALOR_ORIG'].sum()),
        ('VALOR CONCILIADO DÉB (R$)', dc['VALOR_ORIG'].sum()),
        ('VALOR CONCILIADO CRÉ (R$)', cc['VALOR_ORIG'].sum()),
        ('VALOR PENDENTE DÉB (R$)',
         deb[deb['_STATUS'] == 'PENDENTE']['VALOR_ORIG'].sum()),
        ('VALOR PENDENTE CRÉ (R$)',
         cred[cred['_STATUS'] == 'PENDENTE']['VALOR_ORIG'].sum()),
        ('', ''),
        ('TAXA CONCILIAÇÃO DÉB (%)',
         round(len(dc) / len(deb) * 100, 1) if len(deb) else 0),
        ('TAXA CONCILIAÇÃO CRÉ (%)',
         round(len(cc) / len(cred) * 100, 1) if len(cred) else 0),
        ('', ''),
        ('Gerado em', pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')),
    ]
    ws4.column_dimensions['A'].width = 38
    ws4.column_dimensions['B'].width = 22
    _hdr(ws4, 1, ['MÉTRICA', 'VALOR'], '2C2C2A')
    for r, (k, v) in enumerate(resumo, 2):
        ck = ws4.cell(row=r, column=1, value=k)
        cv = ws4.cell(row=r, column=2, value=v)
        for cell in (ck, cv):
            cell.border = BORDA
            cell.alignment = Alignment(vertical='center')
        cv.alignment = Alignment(horizontal='right', vertical='center')
        if isinstance(v, float):
            cv.number_format = '#,##0.00'
    wb.save(dest_path)


# ══════════════════════════════════════════════════════════════
# WIDGET — FILTRO DE DATA INDEPENDENTE POR PAINEL
# ══════════════════════════════════════════════════════════════

MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun",
            "Jul","Ago","Set","Out","Nov","Dez"]


class FiltroDia(tk.Frame):
    """
    Filtro de data com dois níveis:
      • Linha 1 — navegação por mês: ◀ ANO ▶  +  botões Jan…Dez
      • Linha 2 — navegação por dia dentro do mês: ◀ DD/MM/YYYY ▶  + entrada livre
    """

    def __init__(self, parent, cor_acento, on_change, **kw):
        super().__init__(parent, bg=COR["bg"], **kw)
        self.on_change   = on_change
        self.cor         = cor_acento
        self.todas_datas: list[date] = []
        self._data: date | None      = None

        # Índice rápido: (ano, mes) → lista de datas disponíveis
        self._por_mes: dict = {}

        self._build_linha_mes()
        self._build_linha_dia()

    # ── construção ────────────────────────────────────────────

    def _build_linha_mes(self):
        linha = tk.Frame(self, bg=COR["bg"])
        linha.pack(fill='x', padx=6, pady=(4, 1))

        # Navegação de ano
        self._btn_ano_ant = tk.Button(
            linha, text="◀", command=self._ano_anterior,
            bg=COR["bg"], fg=self.cor, font=FONTE_BOLD,
            relief='flat', cursor='hand2', bd=0,
            activebackground=COR["bg"])
        self._btn_ano_ant.pack(side='left')

        self._lbl_ano = tk.Label(
            linha, text="——", bg=COR["bg"],
            fg=self.cor, font=("Helvetica", 10, "bold"), width=5)
        self._lbl_ano.pack(side='left')

        self._btn_ano_prox = tk.Button(
            linha, text="▶", command=self._ano_proximo,
            bg=COR["bg"], fg=self.cor, font=FONTE_BOLD,
            relief='flat', cursor='hand2', bd=0,
            activebackground=COR["bg"])
        self._btn_ano_prox.pack(side='left', padx=(0, 6))

        # Botões de mês Jan … Dez
        self._btns_mes = {}
        for i, nome in enumerate(MESES_PT, 1):
            btn = tk.Button(
                linha, text=nome,
                command=lambda m=i: self._ir_para_mes(m),
                bg=COR["bg"], fg=COR["muted"],
                font=("Helvetica", 8), relief='flat',
                cursor='hand2', bd=0, padx=3,
                activebackground=COR["bg"])
            btn.pack(side='left', padx=1)
            self._btns_mes[i] = btn

    def _build_linha_dia(self):
        linha = tk.Frame(self, bg=COR["bg"])
        linha.pack(fill='x', padx=6, pady=(1, 4))

        tk.Button(linha, text="◀", command=self._anterior,
                  bg=COR["bg"], fg=self.cor, font=FONTE_BOLD,
                  relief='flat', cursor='hand2', bd=0,
                  activebackground=COR["bg"]).pack(side='left')

        self._lbl_dia = tk.Label(
            linha, text="—", bg=COR["bg"],
            fg=COR["text"], font=("Helvetica", 10, "bold"),
            width=12, anchor='center')
        self._lbl_dia.pack(side='left', padx=2)

        tk.Button(linha, text="▶", command=self._proximo,
                  bg=COR["bg"], fg=self.cor, font=FONTE_BOLD,
                  relief='flat', cursor='hand2', bd=0,
                  activebackground=COR["bg"]).pack(side='left')

        self._entry = tk.Entry(linha, width=10, font=FONTE,
                               relief='solid', bd=1)
        self._entry.insert(0, "dd/mm/aa")
        self._entry.bind('<FocusIn>',  lambda e: self._entry.delete(0, 'end')
                         if self._entry.get() == "dd/mm/aa" else None)
        self._entry.bind('<FocusOut>', lambda e: self._entry.insert(0, "dd/mm/aa")
                         if self._entry.get() == "" else None)
        self._entry.bind('<Return>', self._ir_para_entry)
        self._entry.pack(side='left', padx=(8, 2))

        tk.Button(linha, text="OK", command=self._ir_para_entry,
                  bg=self.cor, fg="white", font=FONTE,
                  relief='flat', cursor='hand2', bd=0, padx=6,
                  activebackground=self.cor).pack(side='left')

    # ── API pública ───────────────────────────────────────────

    def set_datas(self, datas):
        self.todas_datas = sorted(datas)
        # Indexar por (ano, mês)
        self._por_mes = {}
        for d in self.todas_datas:
            self._por_mes.setdefault((d.year, d.month), []).append(d)
        if datas:
            self.set(self.todas_datas[0])

    def set(self, d: date):
        self._data = d
        self._lbl_dia.config(text=d.strftime("%d/%m/%Y") if d else "—")
        self._lbl_ano.config(text=str(d.year) if d else "——")
        self._colorir_btns_mes(d)
        self.on_change(d)

    @property
    def data(self):
        return self._data

    # ── navegação de ano ──────────────────────────────────────

    def _anos_disponiveis(self):
        return sorted({d.year for d in self.todas_datas})

    def _ano_anterior(self):
        if not self._data:
            return
        anos = self._anos_disponiveis()
        idx  = anos.index(self._data.year)
        if idx > 0:
            self._ir_para_ano(anos[idx - 1])

    def _ano_proximo(self):
        if not self._data:
            return
        anos = self._anos_disponiveis()
        idx  = anos.index(self._data.year)
        if idx < len(anos) - 1:
            self._ir_para_ano(anos[idx + 1])

    def _ir_para_ano(self, ano: int):
        # Vai para o primeiro dia disponível do mesmo mês naquele ano,
        # ou o primeiro dia do ano se o mês não existir.
        mes_atual = self._data.month if self._data else 1
        chave = (ano, mes_atual)
        if chave in self._por_mes:
            self.set(self._por_mes[chave][0])
        else:
            # Pega o primeiro dia disponível do ano
            datas_ano = [d for d in self.todas_datas if d.year == ano]
            if datas_ano:
                self.set(datas_ano[0])

    # ── navegação de mês ──────────────────────────────────────

    def _ir_para_mes(self, mes: int):
        if not self._data:
            return
        ano   = self._data.year
        chave = (ano, mes)
        if chave in self._por_mes:
            self.set(self._por_mes[chave][0])

    def _colorir_btns_mes(self, d: date | None):
        if d is None:
            return
        ano = d.year
        for m, btn in self._btns_mes.items():
            disponivel = (ano, m) in self._por_mes
            ativo      = (m == d.month)
            if ativo:
                btn.config(bg=self.cor, fg="white",
                           relief='flat', font=("Helvetica", 8, "bold"))
            elif disponivel:
                btn.config(bg=COR["bg"], fg=self.cor,
                           relief='flat', font=("Helvetica", 8))
            else:
                btn.config(bg=COR["bg"], fg=COR["border"],
                           relief='flat', font=("Helvetica", 8),
                           cursor='arrow')

    # ── navegação de dia ──────────────────────────────────────

    def _anterior(self):
        if not self.todas_datas or self._data is None:
            return
        idx = self.todas_datas.index(self._data)
        if idx > 0:
            self.set(self.todas_datas[idx - 1])

    def _proximo(self):
        if not self.todas_datas or self._data is None:
            return
        idx = self.todas_datas.index(self._data)
        if idx < len(self.todas_datas) - 1:
            self.set(self.todas_datas[idx + 1])

    def _ir_para_entry(self, _=None):
        txt = self._entry.get().strip()
        if not txt or txt == 'dd/mm/aa':
            return
        try:
            d = pd.to_datetime(txt, dayfirst=True).date()
        except Exception:
            messagebox.showerror("Data inválida", "Use o formato dd/mm/aa")
            return
        if self.todas_datas:
            self.set(min(self.todas_datas, key=lambda x: abs((x - d).days)))


# ══════════════════════════════════════════════════════════════
# APLICAÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════

class ConciliadorApp(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master, bg=COR["bg"])

        # Estado independente por conta (989 / 1677 / futuras) — permite
        # alternar sem reimportar, cada uma com seu próprio perfil salvo.
        self.conta_atual  = CONTAS[0]
        self.contas_dados = {c: self._estado_vazio() for c in CONTAS}
        self._aplicar_estado(self.contas_dados[self.conta_atual], atualizar_ui=False)

        self._build_ui()
        self._atualizar_lbl_perfil()
        self._atualizar_btns_conta()

    # ──────────────────────────────────────────────────────────
    # ESTADO POR CONTA
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def _estado_vazio() -> dict:
        return {
            'deb': None, 'cred': None, 'grupos': [], 'grupo_id_counter': 0,
            'sel_deb': [], 'sel_cred': [], 'data_deb': None, 'data_cred': None,
            'perfil_atual': None, 'perfil_slug': None, 'arquivo_origem': None,
            'perfil_data_importacao': None,
        }

    def _snapshot_estado_atual(self) -> dict:
        return {
            'deb': self.deb, 'cred': self.cred, 'grupos': self.grupos,
            'grupo_id_counter': self.grupo_id_counter,
            'sel_deb': self.sel_deb, 'sel_cred': self.sel_cred,
            'data_deb': self.data_deb, 'data_cred': self.data_cred,
            'perfil_atual': self.perfil_atual, 'perfil_slug': self.perfil_slug,
            'arquivo_origem': self.arquivo_origem,
            'perfil_data_importacao': self.perfil_data_importacao,
        }

    def _aplicar_estado(self, estado: dict, atualizar_ui: bool = True):
        """Carrega um estado (de outra conta ou de um perfil) para os
        atributos ativos e, opcionalmente, atualiza a UI já construída."""
        self.deb    = estado['deb']
        self.cred   = estado['cred']
        self.grupos = estado['grupos']
        self.grupo_id_counter = estado['grupo_id_counter']
        self.sel_deb  = estado['sel_deb']
        self.sel_cred = estado['sel_cred']
        self.data_deb  = estado['data_deb']
        self.data_cred = estado['data_cred']
        self.perfil_atual   = estado['perfil_atual']
        self.perfil_slug    = estado['perfil_slug']
        self.arquivo_origem = estado['arquivo_origem']
        self.perfil_data_importacao = estado['perfil_data_importacao']

        if not atualizar_ui:
            return

        self._atualizar_lbl_perfil()
        if self.deb is not None and self.cred is not None:
            datas_d = sorted(self.deb['DATA'].dt.date.unique())
            datas_c = sorted(self.cred['DATA'].dt.date.unique())
            todas   = sorted(set(datas_d) | set(datas_c))
            self.painel_deb['filtro'].set_datas(datas_d or todas)
            self.painel_cred['filtro'].set_datas(datas_c or todas)
            if self.data_deb in datas_d:
                self.painel_deb['filtro'].set(self.data_deb)
            if self.data_cred in datas_c:
                self.painel_cred['filtro'].set(self.data_cred)
            self._atualizar_hud()
        else:
            self._limpar_paineis_visual()
        self._esconder_detalhe()

    def _limpar_paineis_visual(self):
        for p in (self.painel_deb, self.painel_cred):
            p['tv'].delete(*p['tv'].get_children())
            p['lbl_count'].config(text="")
            p['lbl_sel'].config(text="")
        for card in self.hud_cards.values():
            card.config(text="—")

    def _selecionar_conta(self, nova_conta: str):
        if nova_conta == self.conta_atual:
            return
        # guarda o estado da conta que está saindo, em memória
        self.contas_dados[self.conta_atual] = self._snapshot_estado_atual()
        self.conta_atual = nova_conta
        self._aplicar_estado(self.contas_dados.get(nova_conta, self._estado_vazio()))
        self._atualizar_btns_conta()
        self.winfo_toplevel().title(f"Conciliador Financeiro — Conta {nova_conta}")
        if self.deb is not None:
            self._set_status(
                f"↔  Conta alterada para {nova_conta}"
                + (f" — perfil '{self.perfil_atual}'" if self.perfil_atual else ""))
        else:
            self._set_status(f"↔  Conta alterada para {nova_conta} — nenhum dado carregado ainda")

    def _confronto_em_breve(self):
        # mantido por compatibilidade; não é mais usado como comando do botão
        messagebox.showinfo(
            "Em breve",
            "O confronto entre as contas 989 e 1677 será implementado em uma próxima etapa.")

    # ──────────────────────────────────────────────────────────
    # CONFRONTO 989 × 1677
    # ──────────────────────────────────────────────────────────

    def _abrir_confronto(self):
        if (getattr(self, 'confronto_fontes', None) and self.conta_atual == 'CONFRONTO'
                and self.grupos):
            if not messagebox.askyesno(
                    "Confronto em andamento",
                    "Há conciliações do confronto atual ainda não salvas.\n"
                    "Deseja descartá-las e selecionar novos perfis?"):
                return

        entry989 = self._dialog_escolher_perfil(
            '989', "Confronto — selecione o perfil da conta 989")
        if not entry989:
            return
        entry1677 = self._dialog_escolher_perfil(
            '1677', "Confronto — selecione o perfil da conta 1677")
        if not entry1677:
            return
        self._montar_confronto(entry989, entry1677)

    def _dialog_escolher_perfil(self, conta: str, titulo: str):
        """Lista os perfis salvos de uma conta e retorna a entrada escolhida
        (bloqueia até o usuário confirmar ou cancelar)."""
        lista = [e for e in carregar_index() if e.get('conta') == conta]
        if not lista:
            messagebox.showinfo(
                "Perfis", f"Nenhum perfil salvo para a conta {conta} ainda.\n"
                          f"Salve o progresso da conta {conta} antes de confrontar.")
            return None

        resultado = {'entry': None}
        dlg = tk.Toplevel(self)
        dlg.title(titulo)
        dlg.configure(bg=COR["bg"])
        dlg.geometry("560x340")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        tk.Label(dlg, text=titulo, bg=COR["bg"], fg=COR["text"],
                 font=FONTE_BOLD).pack(anchor='w', padx=12, pady=(12, 6))

        frame_lb = tk.Frame(dlg, bg=COR["bg"])
        frame_lb.pack(fill='both', expand=True, padx=12)
        sb = ttk.Scrollbar(frame_lb, orient='vertical')
        lb = tk.Listbox(frame_lb, font=FONTE, yscrollcommand=sb.set,
                         activestyle='none', selectmode='browse',
                         relief='solid', bd=1)
        sb.config(command=lb.yview)
        lb.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        lista_ordenada = sorted(lista, key=lambda e: e.get('nome_empresa', '').lower())
        for e in lista_ordenada:
            lb.insert('end',
                f"{e.get('nome_empresa', '?')}   |   salvo em {e.get('data_ultima_gravacao', '—')}"
                f"   |   {e.get('n_deb', 0)} déb. / {e.get('n_cred', 0)} créd.")

        def _confirmar():
            sel = lb.curselection()
            if sel:
                resultado['entry'] = lista_ordenada[sel[0]]
            dlg.destroy()

        botoes = tk.Frame(dlg, bg=COR["bg"])
        botoes.pack(fill='x', padx=12, pady=10)
        tk.Button(botoes, text="Selecionar", command=_confirmar,
                  bg=COR["azul"], fg="white", font=FONTE_BOLD, relief='flat',
                  cursor='hand2', padx=12, pady=6, bd=0).pack(side='left')
        tk.Button(botoes, text="Cancelar", command=dlg.destroy,
                  bg=COR["bg"], fg=COR["muted"], font=FONTE, relief='flat',
                  cursor='hand2', padx=10, pady=6, bd=0).pack(side='right')
        lb.bind('<Double-Button-1>', lambda e: _confirmar())

        self.wait_window(dlg)
        return resultado['entry']

    def _montar_confronto(self, entry989: dict, entry1677: dict):
        try:
            with open(caminho_perfil(entry989['slug'], '989'), 'rb') as f:
                dados989 = pickle.load(f)
            with open(caminho_perfil(entry1677['slug'], '1677'), 'rb') as f:
                dados1677 = pickle.load(f)
        except Exception as e:
            messagebox.showerror("Erro ao carregar perfis", str(e))
            return

        # guarda o estado da conta ativa antes de trocar para o confronto
        self.contas_dados[self.conta_atual] = self._snapshot_estado_atual()

        self.confronto_fontes = {
            '989': {
                'slug': entry989['slug'],
                'nome': dados989.get('nome_empresa', entry989.get('nome_empresa')),
                'deb': dados989['deb'], 'cred': dados989['cred'],
                'grupos': dados989.get('grupos', []),
                'grupo_id_counter': dados989.get('grupo_id_counter', 0),
                'arquivo_origem': dados989.get('arquivo_origem'),
                'data_importacao': dados989.get('data_importacao'),
            },
            '1677': {
                'slug': entry1677['slug'],
                'nome': dados1677.get('nome_empresa', entry1677.get('nome_empresa')),
                'deb': dados1677['deb'], 'cred': dados1677['cred'],
                'grupos': dados1677.get('grupos', []),
                'grupo_id_counter': dados1677.get('grupo_id_counter', 0),
                'arquivo_origem': dados1677.get('arquivo_origem'),
                'data_importacao': dados1677.get('data_importacao'),
            },
        }
        self._seed_cx_counter()

        deb_pool  = self._montar_pool_confronto('deb')
        cred_pool = self._montar_pool_confronto('cred')

        self.conta_atual = 'CONFRONTO'
        estado = {
            'deb': deb_pool, 'cred': cred_pool, 'grupos': [], 'grupo_id_counter': 0,
            'sel_deb': [], 'sel_cred': [], 'data_deb': None, 'data_cred': None,
            'perfil_atual': (
                f"Confronto — {self.confronto_fontes['989']['nome']} (989) × "
                f"{self.confronto_fontes['1677']['nome']} (1677)"),
            'perfil_slug': None, 'arquivo_origem': None, 'perfil_data_importacao': None,
        }
        self.contas_dados['CONFRONTO'] = estado
        self._aplicar_estado(estado)
        self._atualizar_btns_conta()
        self.winfo_toplevel().title("Conciliador Financeiro — Confronto 989 × 1677")
        self._set_status(
            f"🔀  Confronto carregado — {len(deb_pool)} débito(s) pendente(s) | "
            f"{len(cred_pool)} crédito(s) pendente(s) entre as duas contas. "
            f"Conciliações aqui refletem no status das contas de origem.")

    def _montar_pool_confronto(self, tipo: str) -> pd.DataFrame:
        frames = []
        for conta, fonte in self.confronto_fontes.items():
            df = fonte[tipo]
            pend = df[df['_STATUS'] == 'PENDENTE'].copy()
            pend['_ID_ORIG']       = pend['_ID']
            pend['_CONTA_ORIGEM']  = conta
            pend['_ID']            = conta + '_' + pend['_ID_ORIG'].astype(str)
            pend['_STATUS']        = 'PENDENTE'
            pend['_GRUPO']         = None
            frames.append(pend)
        if frames and any(len(f) for f in frames):
            return pd.concat(frames, ignore_index=True)
        base_cols = list(self.confronto_fontes['989'][tipo].columns) + ['_ID_ORIG', '_CONTA_ORIGEM']
        return pd.DataFrame(columns=base_cols)

    def _seed_cx_counter(self):
        maxn = 0
        for conta in ('989', '1677'):
            for g in self.confronto_fontes[conta]['grupos']:
                gid = g.get('gid', '')
                if gid.startswith('CX'):
                    try:
                        maxn = max(maxn, int(gid[2:]))
                    except ValueError:
                        pass
        self.confronto_cx_counter = maxn

    def _propagar_confronto_grupo(self, grupo: dict):
        """Ao formar um grupo dentro do confronto, reflete o status
        CONCILIADO nas contas de origem (989 e/ou 1677)."""
        self.confronto_cx_counter += 1
        cx_gid = f"CX{self.confronto_cx_counter:05d}"
        grupo['gid_origem_manual'] = grupo['gid']
        grupo['gid'] = cx_gid   # o grupo passa a existir só com o gid CX

        for tipo, ids in (('deb', grupo['deb']), ('cred', grupo['cred'])):
            pool = self.deb if tipo == 'deb' else self.cred
            for iid in ids:
                row = pool[pool['_ID'] == iid]
                if row.empty:
                    continue
                conta_origem = row.iloc[0]['_CONTA_ORIGEM']
                id_orig      = row.iloc[0]['_ID_ORIG']
                contraparte  = '1677' if conta_origem == '989' else '989'

                fonte_df = self.confronto_fontes[conta_origem][tipo]
                m = fonte_df['_ID'] == id_orig
                if m.any():
                    fonte_df.loc[m, '_STATUS'] = 'CONCILIADO'
                    fonte_df.loc[m, '_GRUPO']  = cx_gid

                grupos_fonte = self.confronto_fontes[conta_origem]['grupos']
                entrada = next((g for g in grupos_fonte if g['gid'] == cx_gid), None)
                if entrada is None:
                    entrada = {
                        'gid': cx_gid, 'deb': [], 'cred': [],
                        'metodo': f'CONFRONTO_{contraparte}',
                        'nota': f"Conciliado via Confronto de contas com a conta {contraparte}.",
                    }
                    grupos_fonte.append(entrada)
                entrada[tipo].append(id_orig)

    def _reverter_confronto_grupo(self, grupo: dict):
        gid = grupo['gid']
        for tipo, ids in (('deb', grupo['deb']), ('cred', grupo['cred'])):
            pool = self.deb if tipo == 'deb' else self.cred
            for iid in ids:
                row = pool[pool['_ID'] == iid]
                if row.empty:
                    continue
                conta_origem = row.iloc[0]['_CONTA_ORIGEM']
                id_orig      = row.iloc[0]['_ID_ORIG']
                fonte_df = self.confronto_fontes[conta_origem][tipo]
                m = fonte_df['_ID'] == id_orig
                fonte_df.loc[m, '_STATUS'] = 'PENDENTE'
                fonte_df.loc[m, '_GRUPO']  = None
        for conta in ('989', '1677'):
            self.confronto_fontes[conta]['grupos'] = [
                g for g in self.confronto_fontes[conta]['grupos'] if g['gid'] != gid]

    def _salvar_confronto(self):
        if not getattr(self, 'confronto_fontes', None):
            messagebox.showwarning("Aviso", "Nenhum confronto carregado.")
            return
        agora = datetime.now().strftime('%d/%m/%Y %H:%M')
        for conta, fonte in self.confronto_fontes.items():
            dados = {
                'versao': 1,
                'nome_empresa': fonte['nome'],
                'conta': conta,
                'arquivo_origem': fonte.get('arquivo_origem'),
                'data_importacao': fonte.get('data_importacao') or agora,
                'data_ultima_gravacao': agora,
                'deb': fonte['deb'], 'cred': fonte['cred'],
                'grupos': fonte['grupos'],
                'grupo_id_counter': fonte.get('grupo_id_counter', 0),
            }
            try:
                with open(caminho_perfil(fonte['slug'], conta), 'wb') as f:
                    pickle.dump(dados, f)
                upsert_index({
                    'slug': fonte['slug'], 'conta': conta, 'nome_empresa': fonte['nome'],
                    'arquivo_origem': fonte.get('arquivo_origem'),
                    'data_importacao': dados['data_importacao'],
                    'data_ultima_gravacao': agora,
                    'n_deb': len(fonte['deb']), 'n_cred': len(fonte['cred']),
                    'n_grupos': len(fonte['grupos']),
                })
            except Exception as e:
                messagebox.showerror("Erro ao salvar confronto", str(e))
                return
        nome989  = self.confronto_fontes['989']['nome']
        nome1677 = self.confronto_fontes['1677']['nome']
        self._set_status(
            f"💾  Confronto salvo — perfil '{nome989}' (989) e "
            f"'{nome1677}' (1677) atualizados às {agora}")

    # ──────────────────────────────────────────────────────────
    # CONSTRUÇÃO DA UI
    # ──────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_topbar()
        self._build_hud()
        self._build_corpo()
        self._build_statusbar()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=_t.CARD_BG, height=48)
        bar.pack(fill='x')
        bar.pack_propagate(False)
        tk.Label(bar, text="⬡  CONCILIADOR", bg=_t.CARD_BG, fg="white",
                 font=("Helvetica", 13, "bold")).pack(side='left', padx=(20, 14))

        # ── Seletor de conta ────────────────────────────────────
        seletor = tk.Frame(bar, bg=_t.CARD_BG)
        seletor.pack(side='left')
        self.btns_conta = {}
        for conta in CONTAS:
            b = tk.Button(
                seletor, text=conta,
                command=lambda c=conta: self._selecionar_conta(c),
                font=("Helvetica", 9, "bold"), relief='flat', cursor='hand2',
                bd=0, padx=12, pady=6)
            b.pack(side='left', padx=2)
            self.btns_conta[conta] = b
        self.btn_confronto = tk.Button(
            seletor, text="🔀  Confronto 989×1677",
            command=self._abrir_confronto,
            bg=_t.SURFACE2, fg=_t.FG_MUTED, font=("Helvetica", 9, "bold"),
            relief='flat', cursor='hand2', bd=0, padx=10, pady=6)
        self.btn_confronto.pack(side='left', padx=(10, 0))

        self.lbl_perfil = tk.Label(
            bar, text="Perfil: —", bg=_t.CARD_BG, fg=_t.FG_MUTED,
            font=("Helvetica", 9, "italic"))
        self.lbl_perfil.pack(side='left', padx=(14, 0))
        bf = tk.Frame(bar, bg=_t.CARD_BG)
        bf.pack(side='right', padx=12)
        for txt, cmd, bg in [
            ("📂  Abrir Arquivo",     self._abrir_arquivo,     _t.SURFACE2),
            ("🗂  Abrir Perfil",      self._abrir_perfil,      _t.SURFACE2),
            ("⚡  Conciliar Tudo",    self._conciliar_tudo,    COR["azul"]),
            ("🔗  Conciliar Seleção", self._conciliar_selecao, COR["verde"]),
            ("✕  Limpar Seleção",    self._limpar_selecao,    "#7a5214"),
            ("💾  Salvar Progresso", self._salvar_progresso,  "#0e6b46"),
            ("📤  Exportar XLSX",    self._exportar,           COR["vermelho"]),
        ]:
            tk.Button(bf, text=txt, command=cmd, bg=bg, fg="white",
                      font=("Helvetica", 9, "bold"), relief='flat',
                      cursor='hand2', padx=10, pady=6, bd=0,
                      activebackground=bg, activeforeground="white"
                      ).pack(side='left', padx=3)

    def _atualizar_btns_conta(self):
        for conta, b in self.btns_conta.items():
            if conta == self.conta_atual:
                b.config(bg=COR["azul"], fg="white", activebackground=COR["azul"])
            else:
                b.config(bg=_t.SURFACE2, fg=_t.FG_MUTED, activebackground=_t.SURFACE2)
        if hasattr(self, 'btn_confronto'):
            if self.conta_atual == 'CONFRONTO':
                self.btn_confronto.config(bg=COR["amber"], fg=_t.CARD_BG,
                                           activebackground=COR["amber"])
            else:
                self.btn_confronto.config(bg=_t.SURFACE2, fg=_t.FG_MUTED,
                                           activebackground=_t.SURFACE2)

    def _build_hud(self):
        f = tk.Frame(self, bg=COR["bg"], pady=8)
        f.pack(fill='x', padx=16)
        self.hud_cards = {}
        specs = [
            ("total_deb",  "TOTAL DÉBITOS",     COR["azul"],  COR["azul_light"]),
            ("total_cred", "TOTAL CRÉDITOS",    COR["verde"], COR["verde_light"]),
            ("conc_deb",   "DÉB. CONCILIADOS",  COR["azul"],  COR["azul_light"]),
            ("conc_cred",  "CRÉ. CONCILIADOS",  COR["verde"], COR["verde_light"]),
            ("pend_deb",   "DÉB. PENDENTES",    COR["amber"], COR["amber_light"]),
            ("pend_cred",  "CRÉ. PENDENTES",    COR["amber"], COR["amber_light"]),
            ("valor_conc", "VALOR CONCILIADO",  COR["verde"], COR["verde_light"]),
            ("taxa",       "TAXA CONCILIAÇÃO",  COR["azul"],  COR["azul_light"]),
        ]
        for key, label, fg, bg in specs:
            card = tk.Frame(f, bg=bg, padx=12, pady=7)
            card.pack(side='left', expand=True, fill='both', padx=3)
            tk.Label(card, text=label, bg=bg, fg=fg, font=FONTE_HUD_LBL).pack(anchor='w')
            v = tk.Label(card, text="—", bg=bg, fg=fg, font=FONTE_HUD)
            v.pack(anchor='w')
            self.hud_cards[key] = v

    def _build_corpo(self):
        corpo = tk.Frame(self, bg=COR["bg"])
        corpo.pack(fill='both', expand=True, padx=16, pady=(4, 0))

        col = tk.Frame(corpo, bg=COR["bg"])
        col.pack(side='left', fill='both', expand=True)

        self.painel_deb = self._build_painel(
            col, "DÉBITOS (VENDAS)",
            ["Data", "Doc/NF", "Histórico", "Loja", "Valor (R$)", "Status"],
            COR["azul"], self._on_filtro_deb, self._atualizar_painel_deb)
        self.painel_deb['frame'].pack(side='left', fill='both', expand=True, padx=(0, 3))

        tk.Frame(col, bg=COR["border"], width=2).pack(side='left', fill='y')

        self.painel_cred = self._build_painel(
            col, "CRÉDITOS (PAGAMENTOS)",
            ["Data", "Doc/ID", "Histórico / Cliente", "LJ", "Valor (R$)", "Status"],
            COR["verde"], self._on_filtro_cred, self._atualizar_painel_cred)
        self.painel_cred['frame'].pack(side='left', fill='both', expand=True, padx=(3, 0))

        self._build_detalhe(corpo)

    def _build_painel(self, parent, titulo, colunas, cor, on_filtro, on_refresh):
        frame = tk.Frame(parent, bg=COR["surface"],
                         highlightbackground=COR["border"], highlightthickness=1)

        # Cabeçalho
        hdr = tk.Frame(frame, bg=cor, padx=10, pady=5)
        hdr.pack(fill='x')
        tk.Label(hdr, text=titulo, bg=cor, fg="white", font=FONTE_BOLD).pack(side='left')
        lbl_count = tk.Label(hdr, text="", bg=cor, fg="white", font=FONTE)
        lbl_count.pack(side='right')

        # Filtro de data
        filt = FiltroDia(frame, cor_acento=cor, on_change=on_filtro)
        filt.pack(fill='x', padx=8)

        # ── Checkbox "Ver todos os dias" ────────────────────────
        chk_frame = tk.Frame(frame, bg=COR["surface"])
        chk_frame.pack(fill='x', padx=8, pady=(0, 2))
        var_todos = tk.BooleanVar(value=False)
        chk_todos = tk.Checkbutton(
            chk_frame, text="Ver todos os dias",
            variable=var_todos,
            command=on_refresh,
            bg=COR["surface"], fg=COR["muted"],
            activebackground=COR["surface"],
            selectcolor=COR["surface"],
            font=FONTE, cursor='hand2',
            relief='flat', bd=0)
        chk_todos.pack(side='left')

        # ── Filtros extras: status + faixa de valor ─────────────
        extra = tk.Frame(frame, bg=COR["surface"])
        extra.pack(fill='x', padx=8, pady=(0, 4))

        tk.Label(extra, text="Status:", bg=COR["surface"], fg=COR["muted"],
                 font=FONTE).pack(side='left')
        cb_status = ttk.Combobox(
            extra, values=["Todos", "Pendente", "Conciliado"],
            state='readonly', width=10, font=FONTE)
        cb_status.set("Todos")
        cb_status.pack(side='left', padx=(4, 10))
        cb_status.bind('<<ComboboxSelected>>', lambda e: on_refresh())

        tk.Label(extra, text="Valor:", bg=COR["surface"], fg=COR["muted"],
                 font=FONTE).pack(side='left')
        ent_valor_min = tk.Entry(extra, width=7, font=FONTE, relief='solid', bd=1)
        ent_valor_min.pack(side='left', padx=(4, 2))
        tk.Label(extra, text="–", bg=COR["surface"], fg=COR["muted"],
                 font=FONTE).pack(side='left')
        ent_valor_max = tk.Entry(extra, width=7, font=FONTE, relief='solid', bd=1)
        ent_valor_max.pack(side='left', padx=(2, 6))
        ent_valor_min.bind('<Return>', lambda e: on_refresh())
        ent_valor_max.bind('<Return>', lambda e: on_refresh())

        tk.Button(extra, text="Aplicar", command=on_refresh,
                  bg=cor, fg="white", font=FONTE, relief='flat',
                  cursor='hand2', bd=0, padx=6,
                  activebackground=cor).pack(side='left', padx=2)

        def _limpar_extra():
            cb_status.set("Todos")
            ent_valor_min.delete(0, 'end')
            ent_valor_max.delete(0, 'end')
            var_todos.set(False)
            on_refresh()

        tk.Button(extra, text="✕ Limpar", command=_limpar_extra,
                  bg=COR["bg"], fg=COR["muted"], font=FONTE, relief='flat',
                  cursor='hand2', bd=0, padx=6).pack(side='left')

        # Label de seleção
        lbl_sel = tk.Label(frame, text="", bg=COR["surface"],
                           fg=cor, font=FONTE_BOLD, anchor='w')
        lbl_sel.pack(fill='x', padx=8)

        # Treeview
        tv_frame = tk.Frame(frame, bg=COR["surface"])
        tv_frame.pack(fill='both', expand=True)

        sn = titulo.split()[0].lower() + '.Treeview'
        st = ttk.Style()
        st.theme_use('clam')
        st.configure(sn, background=COR["surface"], foreground=COR["text"],
                     rowheight=24, fieldbackground=COR["surface"], font=FONTE)
        st.configure(sn + '.Heading', background=COR["bg"],
                     foreground=COR["muted"], font=FONTE_BOLD, relief='flat')
        st.map(sn, background=[('selected', cor)], foreground=[('selected', 'white')])

        vsb = ttk.Scrollbar(tv_frame, orient='vertical')
        hsb = ttk.Scrollbar(tv_frame, orient='horizontal')
        tv  = ttk.Treeview(tv_frame, columns=colunas, show='headings',
                           yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                           style=sn, selectmode='none')
        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        W = {'Data': 78, 'Doc/NF': 105, 'Doc/ID': 75, 'Histórico': 185,
             'Histórico / Cliente': 185, 'Loja': 48, 'LJ': 42,
             'Valor (R$)': 96, 'Status': 82}
        for col in colunas:
            tv.heading(col, text=col)
            tv.column(col, width=W.get(col, 100), minwidth=40,
                      anchor='e' if 'Valor' in col else 'w')

        tv.tag_configure('CONCILIADO', background=COR["conciliado"], foreground=_t.ACCENT_GREEN)
        tv.tag_configure('PENDENTE',   background=COR["surface"],    foreground=COR["text"])
        tv.tag_configure('SEL_DEB',    background=COR["deb_sel"],    foreground='#cfe0ff')
        tv.tag_configure('SEL_CRED',   background=COR["cred_sel"],   foreground='#c8f5de')

        tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        return {'tv': tv, 'lbl_count': lbl_count, 'frame': frame,
                'filtro': filt, 'lbl_sel': lbl_sel,
                'cb_status': cb_status,
                'ent_valor_min': ent_valor_min, 'ent_valor_max': ent_valor_max,
                'var_todos': var_todos}

    # ──────────────────────────────────────────────────────────
    # PAINEL DE DETALHE
    # ──────────────────────────────────────────────────────────

    def _build_detalhe(self, parent):
        det = tk.Frame(parent, bg=COR["detalhe_bg"], width=275,
                       highlightbackground=_t.GRID_LINE, highlightthickness=1)
        det.pack(side='left', fill='y', padx=(8, 0))
        det.pack_propagate(False)

        tk.Label(det, text="🔍  DETALHES DA CONCILIAÇÃO",
                 bg=COR["detalhe_bg"], fg=COR["detalhe_txt"],
                 font=("Helvetica", 9, "bold"), pady=8).pack(fill='x', padx=10)
        tk.Frame(det, bg=_t.GRID_LINE, height=1).pack(fill='x', padx=10)

        self.det_hint = tk.Label(det,
            text="\nClique em qualquer\nlinha para ver os\ndetalhes da\nconciliação.",
            bg=COR["detalhe_bg"], fg=_t.FG_MUTED,
            font=("Helvetica", 9, "italic"), justify='center')
        self.det_hint.pack(pady=20)

        self.det_content = tk.Frame(det, bg=COR["detalhe_bg"])
        # não empacotado ainda — aparece só ao clicar

        self._det_grupo  = self._det_field(self.det_content, "Grupo")
        self._det_metodo = self._det_field(self.det_content, "Método")

        tk.Frame(self.det_content, bg=_t.GRID_LINE, height=1
                 ).pack(fill='x', padx=10, pady=5)

        tk.Label(self.det_content, text="▸  DÉBITO",
                 bg=COR["detalhe_bg"], fg=_t.ACCENT_CYAN,
                 font=("Helvetica", 8, "bold")).pack(anchor='w', padx=12, pady=(2, 0))

        self._det_deb_data = self._det_field(self.det_content, "Data")
        self._det_deb_doc  = self._det_field(self.det_content, "Doc/NF")
        self._det_deb_hist = self._det_field(self.det_content, "Histórico", wrap=True)
        self._det_deb_loja = self._det_field(self.det_content, "Loja")
        self._det_deb_val  = self._det_field(self.det_content, "Valor")

        tk.Frame(self.det_content, bg=_t.GRID_LINE, height=1
                 ).pack(fill='x', padx=10, pady=5)

        tk.Label(self.det_content, text="▸  CRÉDITO(S)",
                 bg=COR["detalhe_bg"], fg=_t.ACCENT_GREEN,
                 font=("Helvetica", 8, "bold")).pack(anchor='w', padx=12, pady=(2, 0))

        # Frame scrollável para créditos
        cred_wrap = tk.Frame(self.det_content, bg=COR["detalhe_bg"])
        cred_wrap.pack(fill='both', expand=True, padx=6, pady=4)
        self._det_cred_frame = cred_wrap

    def _det_field(self, parent, label, wrap=False):
        f = tk.Frame(parent, bg=COR["detalhe_bg"])
        f.pack(fill='x', padx=10, pady=1)
        tk.Label(f, text=label + ":", bg=COR["detalhe_bg"],
                 fg=_t.FG_MUTED, font=FONTE_DET, width=9, anchor='w').pack(side='left')
        kw = dict(wraplength=155, justify='left') if wrap else {}
        lbl = tk.Label(f, text="—", bg=COR["detalhe_bg"],
                       fg=COR["detalhe_txt"],
                       font=FONTE_DET if wrap else FONTE_DET_BOLD,
                       anchor='w', **kw)
        lbl.pack(side='left', fill='x', expand=True)
        return lbl

    def _mostrar_detalhe(self, registro_id: str, tipo: str):
        df  = self.deb if tipo == 'deb' else self.cred
        rec = df[df['_ID'] == registro_id]
        if rec.empty:
            return
        rec = rec.iloc[0]

        self.det_hint.pack_forget()
        self.det_content.pack(fill='both', expand=True)

        # Esconde botão desfazer se existir
        if hasattr(self, '_btn_desfazer') and self._btn_desfazer.winfo_exists():
            self._btn_desfazer.destroy()

        # Pendente
        if rec['_STATUS'] != 'CONCILIADO' or not rec['_GRUPO']:
            self._det_grupo.config(text="PENDENTE", fg=COR["amber"])
            self._det_metodo.config(text="—", fg=COR["detalhe_txt"])
            for lbl in (self._det_deb_data, self._det_deb_doc,
                        self._det_deb_val):
                lbl.config(text="—", fg=COR["detalhe_txt"])
            self._det_deb_hist.config(text="Sem conciliação",
                                      fg=COR["detalhe_txt"])
            for w in self._det_cred_frame.winfo_children():
                w.destroy()
            tk.Label(self._det_cred_frame,
                     text="Sem créditos vinculados.",
                     bg=COR["detalhe_bg"], fg=_t.FG_MUTED,
                     font=("Helvetica", 8, "italic")).pack(anchor='w', pady=4)
            return

        grp = next((g for g in self.grupos if g['gid'] == rec['_GRUPO']), None)
        if not grp:
            return

        self._det_grupo.config(text=grp['gid'], fg=_t.FG_TEXT)
        self._det_metodo.config(
            text=grp['metodo'].replace('_', ' '), fg=COR["detalhe_txt"])

        # Botão desfazer — recriado a cada clique para capturar o gid correto
        if hasattr(self, '_btn_desfazer') and self._btn_desfazer.winfo_exists():
            self._btn_desfazer.destroy()
        self._btn_desfazer = tk.Button(
            self.det_content,
            text="↩  Desfazer Conciliação",
            command=lambda gid=grp['gid']: self._desfazer_grupo(gid),
            bg="#5a1f2c", fg="white",
            font=("Helvetica", 8, "bold"),
            relief='flat', cursor='hand2',
            padx=8, pady=5, bd=0,
            activebackground="#7a2a3c", activeforeground="white"
        )
        self._btn_desfazer.pack(fill='x', padx=10, pady=(4, 2))

        # Débito
        if grp['deb']:
            d0 = self.deb[self.deb['_ID'] == grp['deb'][0]].iloc[0]
            self._det_deb_data.config(text=d0['DATA'].strftime('%d/%m/%Y'))
            self._det_deb_doc.config(text=str(d0.get('DOC_RNO', '') or ''))
            self._det_deb_hist.config(
                text=str(d0.get('HISTORICO', '') or '')[:80])
            lj_txt = str(d0.get('LJ_DEB', '') or '')
            self._det_deb_loja.config(
                text=f"LJ {lj_txt}" if lj_txt else "—",
                fg=_t.ACCENT_CYAN)
            total = self.deb[self.deb['_ID'].isin(
                grp['deb'])]['VALOR_ORIG'].sum()
            self._det_deb_val.config(
                text=f"R$ {total:,.2f}", fg=_t.ACCENT_CYAN)
        else:
            for lbl in (self._det_deb_data, self._det_deb_doc, self._det_deb_val):
                lbl.config(text="—", fg=COR["detalhe_txt"])
            self._det_deb_hist.config(
                text=grp.get('nota', 'Sem débito nesta conta.'),
                fg=COR["detalhe_txt"])
            self._det_deb_loja.config(text="—", fg=COR["detalhe_txt"])

        # Créditos
        for w in self._det_cred_frame.winfo_children():
            w.destroy()

        if not grp['cred']:
            tk.Label(self._det_cred_frame,
                     text=grp.get('nota', 'Sem créditos vinculados nesta conta.'),
                     bg=COR["detalhe_bg"], fg=_t.FG_MUTED,
                     font=("Helvetica", 8, "italic"),
                     wraplength=180, justify='left').pack(anchor='w', pady=4)

        for idx, cid in enumerate(grp['cred'], 1):
            crow = self.cred[self.cred['_ID'] == cid]
            if crow.empty:
                continue
            crow = crow.iloc[0]

            card = tk.Frame(self._det_cred_frame, bg=_t.SURFACE2,
                            highlightbackground=_t.GRID_LINE,
                            highlightthickness=1)
            card.pack(fill='x', pady=2)

            hdr = tk.Frame(card, bg=_t.SEL_BG)
            hdr.pack(fill='x')
            tk.Label(hdr, text=f"  Crédito {idx}/{len(grp['cred'])}",
                     bg=_t.SEL_BG, fg=_t.ACCENT_GREEN,
                     font=("Helvetica", 8, "bold")).pack(side='left', pady=2)
            tk.Label(hdr, text=f"R$ {crow['VALOR_ORIG']:,.2f}  ",
                     bg=_t.SEL_BG, fg=_t.FG_TEXT,
                     font=("Helvetica", 8, "bold")).pack(side='right')

            for label, value in [
                ("Data",  crow['DATA'].strftime('%d/%m/%Y')),
                ("Doc",   str(crow.get('DOC_RNO', '') or '')),
                ("Hist.", str(crow.get('HISTORICO', '') or '')[:65]),
                ("Loja",  f"LJ {crow['LJ_NUM']}" if crow.get('LJ_NUM') else None),
            ]:
                if value is None:
                    continue
                rf = tk.Frame(card, bg=_t.SURFACE2)
                rf.pack(fill='x', padx=6, pady=1)
                tk.Label(rf, text=label + ":", bg=_t.SURFACE2,
                         fg=_t.FG_MUTED, font=("Helvetica", 8),
                         width=6, anchor='w').pack(side='left')
                tk.Label(rf, text=value, bg=_t.SURFACE2,
                         fg=COR["detalhe_txt"],
                         font=("Helvetica", 8), anchor='w',
                         wraplength=148, justify='left').pack(
                             side='left', fill='x')

    def _desfazer_grupo(self, gid: str):
        grp = next((g for g in self.grupos if g['gid'] == gid), None)
        if not grp:
            return

        metodo = grp['metodo']
        n_deb  = len(grp['deb'])
        n_cred = len(grp['cred'])

        confirmado = messagebox.askyesno(
            "Desfazer Conciliação",
            f"Desfazer o grupo {gid}?\n\n"
            f"Método: {metodo.replace('_', ' ')}\n"
            f"{n_deb} débito(s)  ×  {n_cred} crédito(s)\n\n"
            "Todos os registros voltarão para PENDENTE.")
        if not confirmado:
            return

        if self.conta_atual == 'CONFRONTO':
            self._reverter_confronto_grupo(grp)

        # Reverter débitos
        for iid in grp['deb']:
            mask = self.deb['_ID'] == iid
            if mask.any():
                self.deb.loc[mask, '_STATUS'] = 'PENDENTE'
                self.deb.loc[mask, '_GRUPO']  = None

        # Reverter créditos
        for iid in grp['cred']:
            mask = self.cred['_ID'] == iid
            if mask.any():
                self.cred.loc[mask, '_STATUS'] = 'PENDENTE'
                self.cred.loc[mask, '_GRUPO']  = None

        # Remover grupo da lista
        self.grupos = [g for g in self.grupos if g['gid'] != gid]

        # Atualizar tela
        self._esconder_detalhe()
        self._atualizar_painel_deb()
        self._atualizar_painel_cred()
        self._atualizar_hud()
        self._set_status(
            f"↩  Grupo {gid} desfeito — "
            f"{n_deb} débito(s) e {n_cred} crédito(s) voltaram para PENDENTE.")

    def _esconder_detalhe(self):
        self.det_content.pack_forget()
        self.det_hint.pack(pady=20)

    # ──────────────────────────────────────────────────────────
    # CALLBACKS DE FILTRO
    # ──────────────────────────────────────────────────────────

    def _on_filtro_deb(self, d):
        self.data_deb = d
        self._atualizar_painel_deb()

    def _on_filtro_cred(self, d):
        self.data_cred = d
        self._atualizar_painel_cred()

    # ──────────────────────────────────────────────────────────
    # FILTROS EXTRA — STATUS E VALOR
    # ──────────────────────────────────────────────────────────

    @staticmethod
    def _parse_valor(txt: str):
        """Converte texto (aceita '1234,56' ou '1234.56') em float, ou None."""
        txt = (txt or '').strip().replace('R$', '').strip()
        if not txt:
            return None
        if ',' in txt:
            txt = txt.replace('.', '').replace(',', '.')
        try:
            return float(txt)
        except ValueError:
            return None

    def _aplicar_filtros_extra(self, df: pd.DataFrame, painel: dict) -> pd.DataFrame:
        """Aplica os filtros de status e faixa de valor de um painel sobre o df."""
        status = painel['cb_status'].get()
        if status == 'Pendente':
            df = df[df['_STATUS'] == 'PENDENTE']
        elif status == 'Conciliado':
            df = df[df['_STATUS'] == 'CONCILIADO']

        vmin = self._parse_valor(painel['ent_valor_min'].get())
        vmax = self._parse_valor(painel['ent_valor_max'].get())
        if vmin is not None:
            df = df[df['VALOR_ORIG'] >= vmin]
        if vmax is not None:
            df = df[df['VALOR_ORIG'] <= vmax]
        return df

    # ──────────────────────────────────────────────────────────
    # AÇÕES
    # ──────────────────────────────────────────────────────────

    def _abrir_arquivo(self):
        if self.conta_atual == 'CONFRONTO':
            messagebox.showinfo(
                "Confronto ativo",
                "Selecione a conta 989 ou 1677 no topo antes de importar um arquivo.")
            return
        path = filedialog.askopenfilename(
            title="Selecione o arquivo de razão",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not path:
            return
        self._set_status(f"Carregando {os.path.basename(path)}…")
        self.update()
        try:
            self.deb, self.cred = load_and_prepare(path)
            self.grupos = []
            self.grupo_id_counter = 0
            self.sel_deb  = []
            self.sel_cred = []

            # ── Nome da empresa / perfil ────────────────────────────
            sugestao = os.path.splitext(os.path.basename(path))[0]
            nome = simpledialog.askstring(
                "Perfil da empresa",
                "Nome da empresa para este perfil\n"
                "(usado para salvar e continuar o progresso depois):",
                initialvalue=sugestao, parent=self)
            if nome and nome.strip():
                self.perfil_atual = nome.strip()
                self.perfil_slug  = slug_perfil(self.perfil_atual)
                self.perfil_data_importacao = datetime.now().strftime('%d/%m/%Y %H:%M')
            else:
                self.perfil_atual = None
                self.perfil_slug  = None
                self.perfil_data_importacao = None
            self.arquivo_origem = os.path.basename(path)
            self._atualizar_lbl_perfil()

            datas_d = sorted(self.deb['DATA'].dt.date.unique())
            datas_c = sorted(self.cred['DATA'].dt.date.unique())
            todas   = sorted(set(datas_d) | set(datas_c))

            self.painel_deb['filtro'].set_datas(datas_d or todas)
            self.painel_cred['filtro'].set_datas(datas_c or todas)
            self._atualizar_hud()
            self._esconder_detalhe()
            self._set_status(
                f"✓  Conta {self.conta_atual} — {len(self.deb):,} débitos | "
                f"{len(self.cred):,} créditos  |  Período: {todas[0]} → {todas[-1]}")
        except Exception as e:
            messagebox.showerror("Erro ao carregar", str(e))

    # ──────────────────────────────────────────────────────────
    # PERFIS — SALVAR / ABRIR PROGRESSO
    # ──────────────────────────────────────────────────────────

    def _atualizar_lbl_perfil(self):
        if self.perfil_atual:
            self.lbl_perfil.config(text=f"Perfil: {self.perfil_atual}", fg=_t.FG_TEXT)
        else:
            self.lbl_perfil.config(text="Perfil: — (não salvo)", fg=_t.FG_MUTED)

    def _salvar_progresso(self):
        if self.conta_atual == 'CONFRONTO':
            self._salvar_confronto()
            return
        if self.deb is None or self.cred is None:
            messagebox.showwarning("Aviso", "Nenhum dado carregado.")
            return
        if not self.perfil_atual:
            nome = simpledialog.askstring(
                "Salvar progresso",
                "Informe um nome de empresa para criar o perfil:", parent=self)
            if not nome or not nome.strip():
                return
            self.perfil_atual = nome.strip()
            self.perfil_slug  = slug_perfil(self.perfil_atual)
            self.perfil_data_importacao = datetime.now().strftime('%d/%m/%Y %H:%M')
            self._atualizar_lbl_perfil()

        agora = datetime.now().strftime('%d/%m/%Y %H:%M')
        dados = {
            'versao': 1,
            'nome_empresa': self.perfil_atual,
            'conta': self.conta_atual,
            'arquivo_origem': self.arquivo_origem,
            'data_importacao': self.perfil_data_importacao or agora,
            'data_ultima_gravacao': agora,
            'deb': self.deb,
            'cred': self.cred,
            'grupos': self.grupos,
            'grupo_id_counter': self.grupo_id_counter,
        }
        try:
            with open(caminho_perfil(self.perfil_slug, self.conta_atual), 'wb') as f:
                pickle.dump(dados, f)
            upsert_index({
                'slug': self.perfil_slug,
                'conta': self.conta_atual,
                'nome_empresa': self.perfil_atual,
                'arquivo_origem': self.arquivo_origem,
                'data_importacao': dados['data_importacao'],
                'data_ultima_gravacao': agora,
                'n_deb': len(self.deb),
                'n_cred': len(self.cred),
                'n_grupos': len(self.grupos),
            })
            self._set_status(
                f"💾  Progresso salvo — conta {self.conta_atual} — "
                f"perfil '{self.perfil_atual}' às {agora}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar perfil", str(e))

    def _abrir_perfil(self):
        if self.conta_atual == 'CONFRONTO':
            messagebox.showinfo(
                "Confronto ativo",
                "Selecione a conta 989 ou 1677 no topo antes de abrir um perfil individual.")
            return
        lista = [e for e in carregar_index() if e.get('conta') == self.conta_atual]
        if not lista:
            messagebox.showinfo(
                "Perfis",
                f"Nenhum perfil salvo ainda para a conta {self.conta_atual}.")
            return

        dlg = tk.Toplevel(self)
        dlg.title(f"Abrir Perfil — Conta {self.conta_atual}")
        dlg.configure(bg=COR["bg"])
        dlg.geometry("560x360")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        tk.Label(dlg, text=f"Selecione um perfil da conta {self.conta_atual} para carregar:",
                 bg=COR["bg"], fg=COR["text"], font=FONTE_BOLD
                 ).pack(anchor='w', padx=12, pady=(12, 6))

        frame_lb = tk.Frame(dlg, bg=COR["bg"])
        frame_lb.pack(fill='both', expand=True, padx=12)
        sb = ttk.Scrollbar(frame_lb, orient='vertical')
        lb = tk.Listbox(frame_lb, font=FONTE, yscrollcommand=sb.set,
                         activestyle='none', selectmode='browse',
                         relief='solid', bd=1)
        sb.config(command=lb.yview)
        lb.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        lista_ordenada = sorted(lista, key=lambda e: e.get('nome_empresa', '').lower())
        for e in lista_ordenada:
            lb.insert('end',
                f"{e.get('nome_empresa', '?')}   |   origem: {e.get('arquivo_origem', '—')}"
                f"   |   salvo em {e.get('data_ultima_gravacao', '—')}"
                f"   |   {e.get('n_deb', 0)} déb. / {e.get('n_cred', 0)} créd.")

        def _carregar_selecionado():
            sel = lb.curselection()
            if not sel:
                return
            entry = lista_ordenada[sel[0]]
            dlg.destroy()
            self._carregar_perfil(entry)

        def _excluir_selecionado():
            sel = lb.curselection()
            if not sel:
                return
            entry = lista_ordenada[sel[0]]
            if messagebox.askyesno(
                    "Excluir perfil",
                    f"Excluir permanentemente o perfil '{entry.get('nome_empresa')}' "
                    f"(conta {self.conta_atual})?\nEssa ação não pode ser desfeita.",
                    parent=dlg):
                try:
                    caminho = caminho_perfil(entry['slug'], entry.get('conta', self.conta_atual))
                    if os.path.exists(caminho):
                        os.remove(caminho)
                except Exception as e:
                    messagebox.showerror("Erro", str(e), parent=dlg)
                remover_index(entry['slug'], entry.get('conta', self.conta_atual))
                dlg.destroy()
                self._abrir_perfil()

        botoes = tk.Frame(dlg, bg=COR["bg"])
        botoes.pack(fill='x', padx=12, pady=10)
        tk.Button(botoes, text="Abrir", command=_carregar_selecionado,
                  bg=COR["azul"], fg="white", font=FONTE_BOLD, relief='flat',
                  cursor='hand2', padx=12, pady=6, bd=0).pack(side='left')
        tk.Button(botoes, text="🗑  Excluir", command=_excluir_selecionado,
                  bg=COR["vermelho"], fg="white", font=FONTE, relief='flat',
                  cursor='hand2', padx=10, pady=6, bd=0).pack(side='left', padx=6)
        tk.Button(botoes, text="Cancelar", command=dlg.destroy,
                  bg=COR["bg"], fg=COR["muted"], font=FONTE, relief='flat',
                  cursor='hand2', padx=10, pady=6, bd=0).pack(side='right')

        lb.bind('<Double-Button-1>', lambda e: _carregar_selecionado())

    def _carregar_perfil(self, entry: dict):
        conta = entry.get('conta', self.conta_atual)
        caminho = caminho_perfil(entry['slug'], conta)
        if not os.path.exists(caminho):
            messagebox.showerror("Erro", "Arquivo do perfil não foi encontrado no disco.")
            return
        self._set_status(f"Carregando perfil '{entry.get('nome_empresa')}' (conta {conta})…")
        self.update()
        try:
            with open(caminho, 'rb') as f:
                dados = pickle.load(f)

            self.deb    = dados['deb']
            self.cred   = dados['cred']
            self.grupos = dados.get('grupos', [])
            self.grupo_id_counter = dados.get(
                'grupo_id_counter', len(self.grupos))
            self.sel_deb  = []
            self.sel_cred = []

            self.perfil_atual = dados.get('nome_empresa', entry.get('nome_empresa'))
            self.perfil_slug  = entry['slug']
            self.arquivo_origem = dados.get('arquivo_origem')
            self.perfil_data_importacao = dados.get('data_importacao')
            self._atualizar_lbl_perfil()

            datas_d = sorted(self.deb['DATA'].dt.date.unique())
            datas_c = sorted(self.cred['DATA'].dt.date.unique())
            todas   = sorted(set(datas_d) | set(datas_c))
            self.painel_deb['filtro'].set_datas(datas_d or todas)
            self.painel_cred['filtro'].set_datas(datas_c or todas)
            self._atualizar_hud()
            self._esconder_detalhe()
            self._set_status(
                f"✓  Perfil '{self.perfil_atual}' (conta {conta}) carregado — "
                f"{len(self.deb):,} débitos | {len(self.cred):,} créditos "
                f"| {len(self.grupos)} grupos conciliados")
        except Exception as e:
            messagebox.showerror("Erro ao carregar perfil", str(e))

    def _conciliar_tudo(self):
        if self.deb is None:
            messagebox.showwarning("Aviso", "Abra um arquivo primeiro.")
            return
        self._set_status("Executando conciliação automática…")
        self.update()
        novos = auto_conciliar(self.deb, self.cred, ignorar_loja=True)
        for novo in novos:
            self.grupos.append(novo)
            if self.conta_atual == 'CONFRONTO':
                self._propagar_confronto_grupo(novo)
        self.sel_deb = []
        self.sel_cred = []
        self._atualizar_painel_deb()
        self._atualizar_painel_cred()
        self._atualizar_hud()
        self._set_status(f"✓  {len(novos)} grupos criados.")

    def _conciliar_selecao(self):
        if not self.sel_deb or not self.sel_cred:
            messagebox.showwarning("Seleção incompleta",
                                   "Selecione ao menos um DÉBITO e um CRÉDITO.")
            return
        val_d = round(
            self.deb[self.deb['_ID'].isin(self.sel_deb)]['VALOR_ORIG'].sum(), 2)
        val_c = round(
            self.cred[self.cred['_ID'].isin(self.sel_cred)]['VALOR_ORIG'].sum(), 2)
        if abs(val_d - val_c) > MARGEM_ERRO:
            messagebox.showerror("Valores não batem",
                f"Débitos:  R$ {val_d:,.2f}\n"
                f"Créditos: R$ {val_c:,.2f}\n\n"
                f"Diferença de R$ {abs(val_d-val_c):,.2f} excede a margem de ± R$ {MARGEM_ERRO:,.2f}.")
            return

        self.grupo_id_counter += 1
        gid = f"M{self.grupo_id_counter:05d}"
        for iid in self.sel_deb:
            i = self.deb.index[self.deb['_ID'] == iid][0]
            self.deb.at[i, '_STATUS'] = 'CONCILIADO'
            self.deb.at[i, '_GRUPO']  = gid
        for iid in self.sel_cred:
            i = self.cred.index[self.cred['_ID'] == iid][0]
            self.cred.at[i, '_STATUS'] = 'CONCILIADO'
            self.cred.at[i, '_GRUPO']  = gid
        novo = {'gid': gid, 'deb': list(self.sel_deb),
                'cred': list(self.sel_cred), 'metodo': 'MANUAL'}
        self.grupos.append(novo)
        if self.conta_atual == 'CONFRONTO':
            self._propagar_confronto_grupo(novo)
        self.sel_deb  = []
        self.sel_cred = []
        self._atualizar_painel_deb()
        self._atualizar_painel_cred()
        self._atualizar_hud()
        self._set_status(f"✓  Grupo manual {gid} — R$ {val_d:,.2f}")

    def _limpar_selecao(self):
        self.sel_deb  = []
        self.sel_cred = []
        self._atualizar_painel_deb()
        self._atualizar_painel_cred()
        self._atualizar_sel_labels()

    def _exportar(self):
        if self.deb is None:
            messagebox.showwarning("Aviso", "Nenhum dado carregado.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar relatório",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self._set_status("Gerando Excel…")
        self.update()
        try:
            exportar_xlsx(self.deb, self.cred, self.grupos, path)
            self._set_status(f"✓  Exportado: {path}")
            messagebox.showinfo("Exportado!", f"Arquivo salvo em:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    # ──────────────────────────────────────────────────────────
    # RENDERIZAÇÃO
    # ──────────────────────────────────────────────────────────

    def _atualizar_painel_deb(self):
        if self.deb is None or self.data_deb is None:
            return
        tv = self.painel_deb['tv']
        tv.delete(*tv.get_children())
        if self.painel_deb['var_todos'].get():
            sub = self.deb.copy()
        else:
            sub = self.deb[self.deb['DATA'].dt.date == self.data_deb]
        sub = self._aplicar_filtros_extra(sub, self.painel_deb)
        sub = sub.sort_values('VALOR_ORIG', ascending=False)
        for _, row in sub.iterrows():
            iid, st = row['_ID'], row['_STATUS']
            tag = ('CONCILIADO' if st == 'CONCILIADO'
                   else 'SEL_DEB' if iid in self.sel_deb else 'PENDENTE')
            tv.insert('', 'end', iid=iid, values=(
                row['DATA'].strftime('%d/%m/%Y'),
                str(row.get('DOC_RNO', '') or ''),
                str(row.get('HISTORICO', '') or '')[:50],
                str(row.get('LJ_DEB', '') or ''),
                f"R$ {row['VALOR_ORIG']:,.2f}",
                '✓ CONC.' if st == 'CONCILIADO' else 'PENDENTE'
            ), tags=(tag,))
        pend = (sub['_STATUS'] == 'PENDENTE').sum()
        self.painel_deb['lbl_count'].config(
            text=f"{len(sub)} reg. | {pend} pend.")
        tv.bind('<ButtonRelease-1>', self._on_click_deb)
        self._atualizar_sel_labels()

    def _atualizar_painel_cred(self):
        if self.cred is None or self.data_cred is None:
            return
        tv = self.painel_cred['tv']
        tv.delete(*tv.get_children())
        if self.painel_cred['var_todos'].get():
            sub = self.cred.copy()
        else:
            sub = self.cred[self.cred['DATA'].dt.date == self.data_cred]
        sub = self._aplicar_filtros_extra(sub, self.painel_cred)
        sub = sub.sort_values('VALOR_ORIG', ascending=False)
        for _, row in sub.iterrows():
            iid, st = row['_ID'], row['_STATUS']
            tag = ('CONCILIADO' if st == 'CONCILIADO'
                   else 'SEL_CRED' if iid in self.sel_cred else 'PENDENTE')
            tv.insert('', 'end', iid=iid, values=(
                row['DATA'].strftime('%d/%m/%Y'),
                str(row.get('DOC_RNO', '') or ''),
                str(row.get('HISTORICO', '') or '')[:55],
                str(row.get('LJ_NUM', '') or ''),
                f"R$ {row['VALOR_ORIG']:,.2f}",
                '✓ CONC.' if st == 'CONCILIADO' else 'PENDENTE'
            ), tags=(tag,))
        pend = (sub['_STATUS'] == 'PENDENTE').sum()
        self.painel_cred['lbl_count'].config(
            text=f"{len(sub)} reg. | {pend} pend.")
        tv.bind('<ButtonRelease-1>', self._on_click_cred)
        self._atualizar_sel_labels()

    def _on_click_deb(self, event):
        tv   = self.painel_deb['tv']
        item = tv.identify_row(event.y)
        if not item:
            return
        rec = self.deb[self.deb['_ID'] == item]
        if rec.empty:
            return
        self._mostrar_detalhe(item, 'deb')
        if rec.iloc[0]['_STATUS'] == 'CONCILIADO':
            self._set_status(
                f"Grupo {rec.iloc[0]['_GRUPO']} — veja o painel de detalhes →")
            return
        if item in self.sel_deb:
            self.sel_deb.remove(item)
            tv.item(item, tags=('PENDENTE',))
        else:
            self.sel_deb.append(item)
            tv.item(item, tags=('SEL_DEB',))
        self._atualizar_sel_labels()

    def _on_click_cred(self, event):
        tv   = self.painel_cred['tv']
        item = tv.identify_row(event.y)
        if not item:
            return
        rec = self.cred[self.cred['_ID'] == item]
        if rec.empty:
            return
        self._mostrar_detalhe(item, 'cred')
        if rec.iloc[0]['_STATUS'] == 'CONCILIADO':
            self._set_status(
                f"Grupo {rec.iloc[0]['_GRUPO']} — veja o painel de detalhes →")
            return
        if item in self.sel_cred:
            self.sel_cred.remove(item)
            tv.item(item, tags=('PENDENTE',))
        else:
            self.sel_cred.append(item)
            tv.item(item, tags=('SEL_CRED',))
        self._atualizar_sel_labels()

    def _atualizar_sel_labels(self):
        nd = len(self.sel_deb)
        nc = len(self.sel_cred)
        val_d = round(self.deb[self.deb['_ID'].isin(
            self.sel_deb)]['VALOR_ORIG'].sum(), 2) if nd and self.deb is not None else 0.0
        val_c = round(self.cred[self.cred['_ID'].isin(
            self.sel_cred)]['VALOR_ORIG'].sum(), 2) if nc and self.cred is not None else 0.0
        ok  = nd > 0 and nc > 0 and abs(val_d - val_c) <= MARGEM_ERRO
        self.painel_deb['lbl_sel'].config(
            text=f"{nd} sel.  R$ {val_d:,.2f}" if nd else "",
            fg=COR["azul"])
        suf = "  ✓ pronto para conciliar!" if ok else ""
        self.painel_cred['lbl_sel'].config(
            text=(f"{nc} sel.  R$ {val_c:,.2f}{suf}" if nc else ""),
            fg=COR["verde"] if ok else COR["amber"])

    def _atualizar_hud(self):
        if self.deb is None:
            return
        dc = self.deb[self.deb['_STATUS'] == 'CONCILIADO']
        cc = self.cred[self.cred['_STATUS'] == 'CONCILIADO']
        taxa_d = len(dc) / len(self.deb)  * 100 if len(self.deb)  else 0
        taxa_c = len(cc) / len(self.cred) * 100 if len(self.cred) else 0
        self.hud_cards['total_deb'].config(text=f"{len(self.deb):,}")
        self.hud_cards['total_cred'].config(text=f"{len(self.cred):,}")
        self.hud_cards['conc_deb'].config(text=f"{len(dc):,}")
        self.hud_cards['conc_cred'].config(text=f"{len(cc):,}")
        self.hud_cards['pend_deb'].config(text=f"{len(self.deb)-len(dc):,}")
        self.hud_cards['pend_cred'].config(text=f"{len(self.cred)-len(cc):,}")
        self.hud_cards['valor_conc'].config(
            text=f"R${dc['VALOR_ORIG'].sum():,.0f}")
        self.hud_cards['taxa'].config(
            text=f"D {taxa_d:.1f}%  C {taxa_c:.1f}%")

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=_t.BG_APP, height=26)
        bar.pack(fill='x', side='bottom')
        bar.pack_propagate(False)
        self.lbl_status = tk.Label(
            bar, text="Aguardando arquivo…",
            bg=_t.BG_APP, fg=_t.FG_MUTED, font=FONTE)
        self.lbl_status.pack(side='left', padx=12)

    def _set_status(self, msg: str):
        self.lbl_status.config(text=msg)
        self.update_idletasks()


# ══════════════════════════════════════════════════════════════
# PONTO DE ENTRADA — integração com menu_principal e standalone
# ══════════════════════════════════════════════════════════════

def _build_content(win: tk.Misc) -> None:
    """Monta o Conciliador de Contas dentro de qualquer janela (Tk ou Toplevel)."""
    win.title(f"Conciliador Financeiro — Conta {CONTAS[0]}")
    win.configure(bg=COR["bg"])
    win.geometry("1440x860")
    win.minsize(1100, 680)
    ConciliadorApp(win).pack(fill="both", expand=True)


if __name__ == '__main__':
    root = tk.Tk()
    _build_content(root)
    root.mainloop()