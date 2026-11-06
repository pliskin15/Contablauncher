import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import re
from decimal import Decimal

try:
    import pandas as pd
    PANDAS_OK = True
except Exception:
    PANDAS_OK = False

try:
    from pypdf import PdfReader
    PDF_OK = True
except Exception:
    PDF_OK = False

APP_TITLE = "Leitor de DI – BAP/Padrão"

ESTADOS = ["Amazonas", "Roraima"]

ESTADO_TO_UF = {
    "Amazonas": "AM",
    "Roraima": "RR",
}

# =========================
# Helpers
# =========================
def _to_float_ptbr(num_str, default=None):
    """
    Converte string numerica para float, suportando:
      - Formato pt-BR com ponto de milhar e virgula decimal: "1.713,00" -> 1713.0
      - Formato pt-BR sem ponto de milhar:                  "17,13"    -> 17.13
      - Formato americano (ponto decimal, sem virgula):      "17.13"    -> 17.13
      - Inteiros simples:                                    "11"       -> 11.0
    """
    if num_str is None:
        return default
    s = str(num_str).strip()
    if not s:
        return default
    try:
        has_comma = ',' in s
        has_dot   = '.' in s

        if has_comma and has_dot:
            # Ex: "1.713,00" -- ponto e milhar, virgula e decimal
            s = s.replace('.', '').replace(',', '.')
        elif has_comma and not has_dot:
            # Ex: "17,13" -- virgula e decimal (pt-BR sem milhar)
            s = s.replace(',', '.')
        elif has_dot and not has_comma:
            # Ex: "17.13" (americano) ou "1.713" (milhar pt-BR sem decimais)
            # Heuristica: exatamente 3 digitos apos o ponto => milhar
            parts = s.split('.')
            if len(parts) == 2 and len(parts[1]) == 3 and parts[0].lstrip('-').isdigit() and parts[1].isdigit():
                s = s.replace('.', '')   # remove separador de milhar
            # else: mantem como float americano (ex: "17.13", "332134.0")
        # else: sem virgula nem ponto => inteiro, mantem como esta

        return float(Decimal(s))
    except Exception:
        return default


def _extract_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    return "\n".join([(p.extract_text() or "") for p in reader.pages])


def _deaccent(s: str) -> str:
    import unicodedata
    if s is None:
        return ""
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')


def contar_produtos_por_mmnp(texto: str):
    if not texto:
        return None
    try:
        padrao = re.compile(r"\(\s*MM/N\s*/\s*POL\.?\s*\)", flags=re.I)
        return len(padrao.findall(texto))
    except Exception:
        return None


# =========================
# Scroll support
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self._canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self._vbar = ttk.Scrollbar(self, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vbar.set)
        self._canvas.pack(side='left', fill='both', expand=True)
        self._vbar.pack(side='right', fill='y')

        self.content = ttk.Frame(self._canvas)
        self._win = self._canvas.create_window((0, 0), window=self.content, anchor='nw')

        self.content.bind('<Configure>', lambda e: self._canvas.configure(scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>', lambda e: self._canvas.itemconfigure(self._win, width=e.width))
        self._canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        self._canvas.bind_all('<Button-4>', self._on_mousewheel)
        self._canvas.bind_all('<Button-5>', self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.delta:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        else:
            if getattr(event, 'num', None) == 4:
                self._canvas.yview_scroll(-3, 'units')
            elif getattr(event, 'num', None) == 5:
                self._canvas.yview_scroll(3, 'units')


# =========================
# Parser geral do PDF
# =========================
def parse_pdf_fields(pdf_path: str, modelo: str = "Amazonas") -> dict:
    text = _extract_text(pdf_path) or ""
    text_noacc = _deaccent(text)
    out: dict = {}
    out["modelo_estado"] = modelo
    m = re.search(r"Quantidade de Ad[ií]c[aã]es:\s*(\d+)", text, flags=re.I)
    if not m:
        m = re.search(r"Quantidade de Adicoes:\s*(\d+)", text_noacc, flags=re.I)
    out["quantidade_adicoes"] = int(m.group(1)) if m else None

    # 2) CARGA
    m = re.search(r"Embalagem:.*?Quantidade:\s*(\d+)", text, flags=re.S | re.I)
    out["carga_quantidade"] = int(m.group(1)) if m else None

    m = re.search(r"Peso\s+L[ií]quido:\s*([\d\.,]+)\s*K[gG]", text, flags=re.I)
    if not m:
        m = re.search(r"Peso\s+Liquido:\s*([\d\.,]+)\s*K[gG]", text_noacc, flags=re.I)
    out["carga_peso_liquido_kg"] = _to_float_ptbr(m.group(1)) if m else None

    # 3) DEMONSTRATIVO DE CÁLCULO (NA MOEDA)
    block_m = re.search(
        r"DEMONSTRATIVO DE CALCULO[\s\S]*?NA MOEDA[\s\S]*?(FOB[\s\S]*?)\n\*{5,}",
        text, flags=re.I,
    )
    moeda = {}
    if block_m:
        block = block_m.group(1)
        for key in ["FOB", "FRETE", "SEGURO", "CIF"]:
            lm = re.search(rf"^{key}[^\n]*$", block, flags=re.M | re.I)
            if lm:
                nums = re.findall(r"[\d\.]+,\d{2,}", lm.group(0))
                if nums:
                    moeda[key.lower()] = _to_float_ptbr(nums[-1])

        acrescimos_rs = 0.0
        for lm in re.finditer(r"^ACRESCIMO[^\n]*$", block, flags=re.M | re.I):
            nums = re.findall(r"[\d\.]+,\d{2,}", lm.group(0))
            if nums:
                acrescimos_rs += _to_float_ptbr(nums[-1]) or 0.0
        if acrescimos_rs:
            moeda["acrescimo"] = acrescimos_rs
        out["moeda_real"] = moeda

    # 4) Taxa US$ e SISCOMEX
    m = re.search(r"TAXA DO US\$\s*:\s*([\d,]+)\s*-\s*(\d{2}/\d{2}/\d{4})", text)
    if not m:
        m = re.search(r"TAXA DO US\$\s*:\s*([\d,]+)\s*-\s*(\d{2}/\d{2}/\d{4})", text_noacc)
    if m:
        out["taxa_usd"] = _to_float_ptbr(m.group(1))
        out["taxa_usd_data"] = m.group(2)

    m = re.search(r"TAXA DE UTILIZACAO[^\d]*R\$\s*([\d\.]+,\d{2,})", text, flags=re.I)
    if not m:
        m = re.search(r"TAXA DE UTILIZACAO[^\d]*R\$\s*([\d\.]+,\d{2,})", text_noacc, flags=re.I)
    out["taxa_utilizacao_siscomex_rs"] = _to_float_ptbr(m.group(1)) if m else None

    # 5) II e IPI (geral – primeira ocorrência)
    m = re.search(
        r"Imposto de Importa[cç][aã]o\s*\nRegime de Tributa[cç][aã]o:\s*([^\n]+)\nAl[ií]quota[^:]*:\s*([\d\.,]+)%",
        text, flags=re.I,
    )
    if not m:
        m = re.search(
            r"Imposto de Importacao\s*\nRegime de Tributacao:\s*([^\n]+)\nAliquota[^:]*:\s*([\d\.,]+)%",
            text_noacc, flags=re.I,
        )
    if m:
        out["ii_regime"] = m.group(1).strip()
        out["ii_aliquota_percent"] = _to_float_ptbr(m.group(2))

    m_reg = re.search(
        r"Imposto sobre Produtos Industrializados\s*\nRegime de Tributa[cç][aã]o:\s*([^\n]+)",
        text, flags=re.I,
    )
    if not m_reg:
        m_reg = re.search(
            r"Imposto sobre Produtos Industrializados\s*\nRegime de Tributacao:\s*([^\n]+)",
            text_noacc, flags=re.I,
        )
    if m_reg:
        out["ipi_regime"] = m_reg.group(1).strip()

    m_aliq = re.search(r"Al[ií]quota\s*Advalorem\s*\(TIPI\)\s*([\d\.,]+)%", text, flags=re.I)
    if not m_aliq:
        m_aliq = re.search(r"Aliquota\s*Advalorem\s*\(TIPI\)\s*([\d\.,]+)%", text_noacc, flags=re.I)
    if m_aliq:
        out["ipi_aliquota_percent"] = _to_float_ptbr(m_aliq.group(1))

    # 6) PIS/COFINS geral (primeira adição como referência)
    m = re.search(
        r"Dados Gerais Pis e Cofins\s*\nBase de C[áa]lculo:\s*R\$\s*([\d\.\,]+).*?\nRegime de Tributa[cç][aã]o:\s*([^\n]+)",
        text, flags=re.S | re.I,
    )
    if not m:
        m = re.search(
            r"Dados Gerais Pis e Cofins\s*\nBase de Calculo:\s*R\$\s*([\d\.\,]+).*?\nRegime de Tributacao:\s*([^\n]+)",
            text_noacc, flags=re.S | re.I,
        )
    if m:
        out["piscofins_base_calculo_rs"] = _to_float_ptbr(m.group(1))
        out["piscofins_regime"] = m.group(2).strip()

    m = re.search(r"Pis/Pasep\s*\nAl[ií]quota AdValorem:\s*([\d\.,]+)%", text, flags=re.I)
    if not m:
        m = re.search(r"Pis/Pasep\s*\nAliquota AdValorem:\s*([\d\.,]+)%", text_noacc, flags=re.I)
    if m:
        out["pis_aliquota_percent"] = _to_float_ptbr(m.group(1))

    m = re.search(r"Cofins\s*\nAl[ií]quota AdValorem:\s*([\d\.,]+)%", text, flags=re.I)
    if not m:
        m = re.search(r"Cofins\s*\nAliquota AdValorem:\s*([\d\.,]+)%", text_noacc, flags=re.I)
    if m:
        out["cofins_aliquota_percent"] = _to_float_ptbr(m.group(1))

    # 7) ADIÇÕES por bloco (peso e base PIS/COFINS de cada adição)
    adicoes = _parse_adicoes_blocos(text, text_noacc, out.get("modelo_estado", "Amazonas"))
    if adicoes:
        out["adicoes"] = adicoes

    # 8) Importador
    cnpj_m = re.search(r"CNPJ\s*[:\-]?\s*([\d\./\-]+)", text, flags=re.I)
    if not cnpj_m:
        cnpj_m = re.search(r"CNPJ\s*[:\-]?\s*([\d\./\-]+)", text_noacc, flags=re.I)
    if cnpj_m:
        out["cnpj_importador"] = re.sub(r"\D", "", cnpj_m.group(1))

    # 9) Fabricante
    fab_m = re.search(r"Fabricante/Produtor\s*\nNome:\s*([^\n]+)", text, flags=re.I)
    if not fab_m:
        fab_m = re.search(r"Fabricante/Produtor\s*\nNome:\s*([^\n]+)", text_noacc, flags=re.I)
    if fab_m:
        out["fabricante"] = fab_m.group(1).strip()
    else:
        # fallback genérico
        fab_m2 = re.search(r"(?:Exportador|Fabricante|Produtor)\s*:?\s*([^\n]+)", text, flags=re.I)
        if fab_m2:
            out["fabricante"] = fab_m2.group(1).strip()

    # 10) Contagem por MM/N/POL.
    try:
        out["produtos_qtd_por_mmnp"] = contar_produtos_por_mmnp(text)
    except Exception:
        out["produtos_qtd_por_mmnp"] = None

    return out


def _parse_adicoes_blocos(text: str, text_noacc: str, modelo: str = "Amazonas") -> list:
    """
    Extrai bloco por adição (cabeçalho 'Adição: DI/NNN').
    Para cada bloco captura:
      - peso_liquido_kg
      - piscofins_base_calculo_rs_item
      - vcmv_usd (valor total da adição em USD)
      - lista de itens de descrição (qtde + VUCV)
    """
    PAT_ADICAO = re.compile(
        r"(Adi[cç][aã]o:\s*[\d/\-]+\s*/\s*(\d{3})[\s\S]*?)(?=\nAdi[cç][aã]o:|\Z)",
        flags=re.I,
    )
    matches = list(PAT_ADICAO.finditer(text))
    if not matches:
        matches = list(PAT_ADICAO.finditer(text_noacc))

    adicoes = []
    for mblk in matches:
        blk = mblk.group(1)
        adicao_str = mblk.group(2)

        # Peso líquido da adição
        mp = re.search(
            r"Peso\s+L[ií]quido\s+da\s+Adi[cç][aã]o\s*:\s*([\d\.\,]+)\s*K[gG]",
            blk, flags=re.I,
        )
        if not mp:
            mp = re.search(
                r"Peso\s+Liquido\s+da\s+Adicao\s*:\s*([\d\.\,]+)\s*K[gG]",
                _deaccent(blk), flags=re.I,
            )
        peso_liq = _to_float_ptbr(mp.group(1)) if mp else None

        # VCMV (valor total da adição em USD)
        mv = re.search(r"V[C]?MV:\s*([\d\.,]+)\s*(?:DOLAR|D[ÓO]LAR)\b", blk, flags=re.I)
        vcmv_usd = _to_float_ptbr(mv.group(1)) if mv else None

        # Base PIS/COFINS do item
        mbc = re.search(r"Base de C[áa]lculo:\s*R\$\s*([\d\.,]+)", blk, flags=re.I)
        if not mbc:
            mbc = re.search(r"Base de Calculo:\s*R\$\s*([\d\.,]+)", _deaccent(blk), flags=re.I)
        bc_piscofins = _to_float_ptbr(mbc.group(1)) if mbc else None

        # Itens de descrição (múltiplos por adição no BAP)
        itens_desc = _parse_desc_items(blk, modelo)

        adicoes.append({
            "adicao": adicao_str,
            "peso_liquido_kg": peso_liq,
            "vcmv_usd": vcmv_usd,
            "piscofins_base_calculo_rs_item": bc_piscofins,
            "itens_desc": itens_desc,
            # Para compatibilidade com código anterior (usa 1º item quando há só 1)
            "descricao_qtde": itens_desc[0]["qtde"] if itens_desc else None,
            "descricao_vucv_usd": itens_desc[0]["vucv_usd"] if itens_desc else None,
        })

    return sorted(adicoes, key=lambda x: x.get("adicao") or "")


def _parse_desc_items(blk: str, modelo: str = "Amazonas") -> list:
    """
    Dentro de um bloco de adição, extrai cada linha de produto:
      Amazonas: Qtde: X  VUCV: Y  DESCRICAO, CODIGO_MODELO, CODIGO_PRODUTO
      Roraima:  Qtde: X  VUCV: Y  DESCRICAO, MARCA: BAP - CODIGO
    Retorna lista de dicts com qtde, vucv_usd e produto_codigo.
    """
    blk_noacc = _deaccent(blk)

    # Encontra a seção de Descrição Detalhada
    sec_m = re.search(
        r"Descri[cç][aã]o\s+Detalhada\s+da\s+Mercadoria([\s\S]*?)(?=Imposto de Importa|Dados Gerais|\Z)",
        blk, flags=re.I,
    )
    if not sec_m:
        sec_m = re.search(
            r"Descricao\s+Detalhada\s+da\s+Mercadoria([\s\S]*?)(?=Imposto de Importa|Dados Gerais|\Z)",
            blk_noacc, flags=re.I,
        )
    if not sec_m:
        return []

    sec = sec_m.group(1)

    # Divide por ocorrências de "Qtde:"
    # Cada chunk: "Qtde: X  VUCV: Y \n DESCRICAO, MODELO, CODIGO"
    chunks = re.split(r"(?m)^(?=\s*Qtde\s*:)", sec)

    itens = []
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue

        # Qtde
        m_q = re.search(r"Qtde\s*:\s*([\d\.,]+)", chunk, flags=re.I)
        if not m_q:
            continue
        qtde = _to_float_ptbr(m_q.group(1))

        # VUCV (valor unitário)
        m_v = re.search(r"VUCV\s*:\s*([\d\.,]+)", chunk, flags=re.I)
        vucv_usd = _to_float_ptbr(m_v.group(1)) if m_v else None

        # Última linha não-vazia e não-heading com vírgula → código = penúltimo token
        produto_codigo = ""
        desc_linha = ""
        linhas = [ln for ln in chunk.splitlines() if ln.strip()]
        _headings = re.compile(
            r"^(imposto|dados gerais|descri|demonstrativo|taxa|pis|cofins|ipi|qtde|vucv|vcmv)\b",
            flags=re.I,
        )
        for ln in reversed(linhas):
            s = ln.strip()
            if not s or _headings.match(s):
                continue
            if "," in s:
                desc_linha = s
                break

        if desc_linha:
            if modelo == "Roraima":
                # Tenta primeiro o formato "- CODIGO. MARCA:" (ex: "- FLU6002. MARCA: BAP.")
                m_cod = re.search(
                    r'-\s+([A-Z0-9]{2,15})\.\s+MARCA:',
                    chunk,
                    flags=re.I
                )
                if m_cod:
                    produto_codigo = m_cod.group(1).upper()
                else:
                    # Tenta formato "MARCA: BAP - CODIGO" (ex: "MARCA: BAP - PIV5087")
                    m_cod = re.search(
                        r'MARCA\s*:\s*BAP\s*-\s*([A-Z0-9\-]{2,20})',
                        chunk,
                        flags=re.I
                    )
                    if m_cod:
                        produto_codigo = m_cod.group(1).strip().upper()
                    else:
                        # fallback: penúltimo token por vírgula
                        partes = [p.strip().rstrip(".;") for p in desc_linha.split(",")]
                        if len(partes) >= 2:
                            produto_codigo = partes[-2]
                        elif len(partes) == 1:
                            produto_codigo = partes[-1]
            else:
                # Amazonas: formato "MARCA: BAP - PN: BAN14578D"
                m_cod = re.search(
                    r'(?:MARCA\s*:\s*BAP\s*-\s*)?PN\s*:\s*([A-Z0-9\-]{2,20})',
                    chunk,
                    flags=re.I
                )
                if m_cod:
                    produto_codigo = m_cod.group(1).strip().upper()
                else:
                    # Tenta formato legado "- CODIGO. MARCA:"
                    m_cod = re.search(
                        r'-\s+([A-Z0-9]{2,10})\.\s+MARCA:',
                        chunk,
                        flags=re.I
                    )
                    if m_cod:
                        produto_codigo = m_cod.group(1).upper()
                    else:
                        # fallback: penúltimo token por vírgula
                        partes = [p.strip().rstrip(".;") for p in desc_linha.split(",")]
                        if len(partes) >= 2:
                            produto_codigo = partes[-2]
                        elif len(partes) == 1:
                            produto_codigo = partes[-1]

        itens.append({
            "qtde": qtde,
            "vucv_usd": vucv_usd,
            "produto_codigo": produto_codigo,
            "desc_linha": desc_linha,
        })
    print(repr(desc_linha))
    print("---chunk---")
    print(repr(chunk))
    return itens


# =========================
# Parser BAP (produtos expandidos)
# =========================
def parse_bap_produtos(texto_pdf: str, modelo: str = "Amazonas") -> list:
    """
    Para BAP: lê cada adição e expande seus itens de descrição em uma lista plana.
    Cada item terá: adicao, qtde, vucv_usd, produto_codigo, peso_liq (rateado por qtde dentro da adição).
    """
    text_noacc = _deaccent(texto_pdf)

    PAT_ADICAO = re.compile(
        r"(Adi[cç][aã]o:\s*[\d/\-]+\s*/\s*(\d{3})[\s\S]*?)(?=\nAdi[cç][aã]o:|\Z)",
        flags=re.I,
    )
    matches = list(PAT_ADICAO.finditer(texto_pdf))
    if not matches:
        matches = list(PAT_ADICAO.finditer(text_noacc))

    itens_planos = []

    for mblk in matches:
        blk = mblk.group(1)
        adicao_str = mblk.group(2)

        # Peso total da adição
        mp = re.search(
            r"Peso\s+L[ií]quido\s+da\s+Adi[cç][aã]o\s*:\s*([\d\.\,]+)\s*K[gG]",
            blk, flags=re.I,
        )
        if not mp:
            mp = re.search(
                r"Peso\s+Liquido\s+da\s+Adicao\s*:\s*([\d\.\,]+)\s*K[gG]",
                _deaccent(blk), flags=re.I,
            )
        peso_liq_adicao = _to_float_ptbr(mp.group(1)) if mp else 0.0

        # Base PIS/COFINS da adição
        mbc = re.search(r"Base de C[áa]lculo:\s*R\$\s*([\d\.,]+)", blk, flags=re.I)
        if not mbc:
            mbc = re.search(r"Base de Calculo:\s*R\$\s*([\d\.,]+)", _deaccent(blk), flags=re.I)
        bc_piscofins = _to_float_ptbr(mbc.group(1)) if mbc else 0.0

        itens_desc = _parse_desc_items(blk, modelo)
        if not itens_desc:
            # adição sem itens identificados → skip silencioso
            continue

        # Rateia o peso da adição proporcional à qtde de cada item
        qtde_total_adicao = sum(i["qtde"] or 0.0 for i in itens_desc) or 1.0
        piscofins_total_adicao = bc_piscofins or 0.0

        for it in itens_desc:
            qtde_i = it["qtde"] or 0.0
            prop = (qtde_i / qtde_total_adicao) if qtde_total_adicao > 0 else 0.0
            peso_i = (peso_liq_adicao or 0.0) * prop
            bc_i = piscofins_total_adicao * prop

            itens_planos.append({
                "adicao": adicao_str,
                "descricao_qtde": qtde_i,
                "descricao_vucv_usd": it["vucv_usd"],
                "produto_codigo": it["produto_codigo"],
                "peso_liquido_kg": peso_i,
                "piscofins_base_calculo_rs_item": bc_i,
            })

    return itens_planos


# =========================
# App (Tkinter)
# =========================
class App(tk.Toplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(APP_TITLE)
        self.geometry("900x680")
        self.minsize(860, 600)
        if parent is not None:
            self.transient(parent)
            self.grab_set()
            self.protocol("WM_DELETE_WINDOW", self.destroy)

        self.pdf_path = tk.StringVar(value="")
        self.invoice_path = tk.StringVar(value="")
        self.data_str = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.di_str = tk.StringVar(value="")
        self.tipo = tk.StringVar(value="Padrão")
        self.modelo = tk.StringVar(value=ESTADOS[0])
        self.fornecedor_cod_str = tk.StringVar(value="")
        self.exportador_cod_str = tk.StringVar(value="")
        self.frete_var = tk.StringVar(value="0,00")
        self.arm_var = tk.StringVar(value="0,00")
        self.desp_var = tk.StringVar(value="0,00")
        self.mva_comum_var = tk.StringVar(value="0")

        self.produtos = []  # lista de (codigo_var, mva_var, mva_entry_widget)

        self._build_ui()

    def _build_ui(self):
        page = ScrollableFrame(self)
        page.pack(fill="both", expand=True)
        root = ttk.Frame(page.content, padding=16)
        root.pack(fill="both", expand=True)

        # Data
        linha_data = ttk.Frame(root)
        linha_data.pack(fill="x", pady=(0, 8))
        ttk.Label(linha_data, text="Data:", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Entry(linha_data, textvariable=self.data_str, width=14).pack(side="left", padx=(8, 0))
        ttk.Label(linha_data, text="(dd/mm/aaaa)").pack(side="left", padx=(6, 0))

        # DI
        linha_di = ttk.Frame(root)
        linha_di.pack(fill="x", pady=(0, 4))
        ttk.Label(linha_di, text="DI:", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Entry(linha_di, textvariable=self.di_str, width=22).pack(side="left", padx=(8, 0))

        # Fornecedor / Exportador
        linha_fx = ttk.Frame(root)
        linha_fx.pack(fill="x", pady=(0, 8))
        ttk.Label(linha_fx, text="Fornecedor (cód.):", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(linha_fx, textvariable=self.fornecedor_cod_str, width=16).pack(side="left", padx=(6, 20))
        ttk.Label(linha_fx, text="Exportador (cód.):", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(linha_fx, textvariable=self.exportador_cod_str, width=16).pack(side="left", padx=(6, 0))

        # Tipo
        linha_tipo = ttk.Frame(root)
        linha_tipo.pack(fill="x", pady=8)
        ttk.Label(linha_tipo, text="Tipo:", font=("Segoe UI", 11, "bold")).pack(side="left")
        self.cmb_tipo = ttk.Combobox(
            linha_tipo, textvariable=self.tipo, values=["Padrão", "BAP"], state="readonly", width=12
        )
        self.cmb_tipo.pack(side="left", padx=(8, 0))
        self.cmb_tipo.bind("<<ComboboxSelected>>", self._on_tipo_change)

        # Modelo
        linha_modelo = ttk.Frame(root)
        linha_modelo.pack(fill="x", pady=8)
        ttk.Label(linha_modelo, text="Modelo (Estado):", font=("Segoe UI", 11, "bold")).pack(side="left")
        cmb_modelo = ttk.Combobox(
            linha_modelo, textvariable=self.modelo, values=ESTADOS, state="readonly", width=18
        )
        cmb_modelo.pack(side="left", padx=(8, 0))
        cmb_modelo.bind("<<ComboboxSelected>>", self._on_modelo_change)

        # PDF
        self.linha_pdf = ttk.Frame(root)
        self.linha_pdf.pack(fill="x", pady=8)
        ttk.Label(self.linha_pdf, text="Importar PDF:", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Entry(self.linha_pdf, textvariable=self.pdf_path, width=50, state="readonly").pack(side="left", padx=8)
        ttk.Button(self.linha_pdf, text="Selecionar Arquivo…", command=self.selecionar_pdf).pack(side="left")

        # Invoice (Excel)
        self.linha_invoice = ttk.Frame(root)
        self.linha_invoice.pack(fill="x", pady=(0, 8))
        ttk.Label(self.linha_invoice, text="Invoice (Excel):", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Entry(self.linha_invoice, textvariable=self.invoice_path, width=50, state="readonly").pack(side="left", padx=8)
        ttk.Button(self.linha_invoice, text="Selecionar Arquivo…", command=self.selecionar_invoice).pack(side="left")
        ttk.Label(self.linha_invoice, text="(opcional — Cod. Produto / Peso)", foreground="gray").pack(side="left", padx=(8, 0))

        # Produtos
        self.frame_produtos = ttk.Frame(root)
        self.frame_produtos.pack(fill="x", pady=8)
        ttk.Label(self.frame_produtos, text="Produtos:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        linha_mva_comum = ttk.Frame(self.frame_produtos)
        linha_mva_comum.pack(fill="x", pady=(2, 6))
        ttk.Label(linha_mva_comum, text="MVA (comum) %:", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(linha_mva_comum, textvariable=self.mva_comum_var, width=10).pack(side="left", padx=(6, 0))
        ttk.Label(linha_mva_comum, text="(usado para todos os produtos quando Tipo = BAP e Modelo ≠ Roraima)").pack(side="left", padx=(8, 0))

        self.produtos_scroll = ScrollableFrame(self.frame_produtos)
        try:
            self.produtos_scroll._canvas.configure(height=150)
        except Exception:
            pass
        self.produtos_scroll.pack(fill="x", pady=(2, 6))
        self.produtos_rows_frame = self.produtos_scroll.content

        self._add_produto_row()

        # Despesas
        self.frame_despesas = ttk.Frame(root)
        self.frame_despesas.pack(fill="x", pady=(2, 10))
        ttk.Label(self.frame_despesas, text="Despesas:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        for label, var in [
            ("Frete Internacional (R$):", self.frete_var),
            ("Armazenagem (R$):", self.arm_var),
            ("Despachante (R$):", self.desp_var),
        ]:
            ln = ttk.Frame(self.frame_despesas)
            ln.pack(fill="x", pady=2)
            ttk.Label(ln, text=label).pack(side="left")
            ttk.Entry(ln, textvariable=var, width=16).pack(side="left", padx=(8, 0))

        # Ações
        linha_acoes = ttk.Frame(root)
        linha_acoes.pack(fill="x", pady=(4, 12))
        ttk.Button(linha_acoes, text="Executar", command=self.executar).pack(side="left")
        ttk.Label(linha_acoes, text=" ").pack(side="left")
        ttk.Button(linha_acoes, text="Exportar XLSX", command=self.exportar_xlsx).pack(side="left")

        # Saída
        ttk.Label(root, text="Informações‑Chave", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8, 4))
        frame_texto = ttk.Frame(root)
        frame_texto.pack(fill="both", expand=True)
        self.txt = tk.Text(frame_texto, wrap="word", height=16)
        self.txt.pack(side="left", fill="both", expand=True)
        sc = ttk.Scrollbar(frame_texto, orient="vertical", command=self.txt.yview)
        sc.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=sc.set)

        self._insert_hint()
        self._on_tipo_change()

    def _add_produto_row(self):
        parent = getattr(self, 'produtos_rows_frame', self.frame_produtos)
        linha = ttk.Frame(parent)
        linha.pack(fill="x", pady=2)

        codigo_var = tk.StringVar()
        mva_var = tk.StringVar()

        ttk.Label(linha, text="Código:").pack(side="left")
        ttk.Entry(linha, textvariable=codigo_var, width=20).pack(side="left", padx=(4, 12))
        ttk.Label(linha, text="MVA (%):").pack(side="left")
        mva_entry = ttk.Entry(linha, textvariable=mva_var, width=8)
        mva_entry.pack(side="left", padx=(4, 12))
        ttk.Button(linha, text="+", command=self._add_produto_row).pack(side="left")

        self.produtos.append((codigo_var, mva_var, mva_entry))
        try:
            self._toggle_row_mva_state()
        except Exception:
            pass

    def _ensure_prod_rows(self, n: int):
        if not isinstance(n, int) or n <= 0:
            return
        while len(self.produtos) < n:
            self._add_produto_row()

    def _toggle_row_mva_state(self):
        # No BAP com Roraima, MVA individual por produto → campos habilitados
        # No BAP com Amazonas, usa MVA comum → campos desabilitados
        is_bap = (self.tipo.get() == "BAP")
        is_roraima = (self.modelo.get() == "Roraima")
        disable = is_bap and not is_roraima
        state = "disabled" if disable else "normal"
        for _c_var, _m_var, m_entry in self.produtos:
            try:
                m_entry.configure(state=state)
            except Exception:
                pass

    def _on_modelo_change(self, event=None):
        self._toggle_row_mva_state()

    def _on_tipo_change(self, event=None):
        self.linha_pdf.pack(fill="x", pady=8)
        self._toggle_row_mva_state()

    def selecionar_pdf(self):
        caminho = filedialog.askopenfilename(title="Selecione um PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if caminho:
            self.pdf_path.set(caminho)

    def selecionar_invoice(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a Invoice (Excel)",
            filetypes=[("Planilha Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.invoice_path.set(caminho)

    def _ler_pesos_invoice(self) -> dict:
        """
        Lê o Excel da invoice e retorna um dict {cod_produto: peso_kg}.
        Procura colunas 'Cod. Produto' e 'Peso' (case-insensitive, ignora espaços).
        Aceita vírgula ou ponto como separador decimal.
        """
        caminho = self.invoice_path.get()
        if not caminho or not PANDAS_OK:
            return {}
        try:
            # Tenta ler sem cabeçalho fixo para encontrar a linha correta
            df_raw = pd.read_excel(caminho, header=None, dtype=str)

            # Encontra a linha que contém os cabeçalhos
            header_row = None
            col_cod = None
            col_peso = None
            for i, row in df_raw.iterrows():
                for j, cell in enumerate(row):
                    if isinstance(cell, str) and "cod" in cell.lower() and "produto" in cell.lower():
                        header_row = i
                        col_cod = j
                    if isinstance(cell, str) and "peso" in cell.lower():
                        if header_row == i:  # mesma linha
                            col_peso = j
                if header_row is not None and col_cod is not None and col_peso is not None:
                    break

            if header_row is None or col_cod is None or col_peso is None:
                messagebox.showwarning(
                    "Invoice",
                    "Não foi possível encontrar as colunas 'Cod. Produto' e 'Peso' na planilha.\n"
                    "Verifique se os cabeçalhos estão corretos."
                )
                return {}

            # Lê os dados abaixo do cabeçalho
            pesos = {}
            for i in range(header_row + 1, len(df_raw)):
                cod = df_raw.iat[i, col_cod]
                peso = df_raw.iat[i, col_peso]
                if pd.isna(cod) or cod is None or str(cod).strip() == "":
                    continue
                cod_str = str(cod).strip().upper()
                peso_f = _to_float_ptbr(str(peso).strip()) if peso and not pd.isna(peso) else None
                if cod_str and peso_f is not None:
                    pesos[cod_str] = pesos.get(cod_str, 0.0) + peso_f

            return pesos

        except Exception as e:
            messagebox.showwarning("Invoice", f"Erro ao ler a planilha da invoice:\n{e}")
            return {}

    # ===== BAP: popula UI com itens expandidos =====
    def _apply_bap_auto_products(self, dados: dict, texto_pdf: str):
        """
        Expande todos os itens BAP (múltiplos por adição) em uma lista plana
        e preenche as linhas de produto da UI com os códigos extraídos.
        """
        itens_bap = parse_bap_produtos(texto_pdf, modelo=dados.get("modelo_estado", "Amazonas")) or []

        if not itens_bap:
            return

        # Substitui 'adicoes' por lista plana de itens
        dados["adicoes"] = itens_bap
        dados["bap_itens"] = itens_bap
        # Preserva quantidade_adicoes do PDF (já extraído); só define se ainda não veio do PDF
        if not dados.get("quantidade_adicoes"):
            dados["quantidade_adicoes"] = len(itens_bap)

        codigos_ui = [it.get("produto_codigo") or "" for it in itens_bap]
        self._ensure_prod_rows(len(codigos_ui))
        for i, (codigo_var, _mva_var, *_w) in enumerate(self.produtos):
            if i < len(codigos_ui):
                codigo_var.set(codigos_ui[i])

    # ===== Lógica principal =====
    def executar(self):
        dados = {
            "data": self.data_str.get(),
            "di": self.di_str.get(),
            "tipo": self.tipo.get(),
            "modelo_estado": self.modelo.get(),
            "arquivo_pdf": os.path.basename(self.pdf_path.get()) if self.pdf_path.get() else "-",
            "produtos": [{"codigo": c.get(), "mva": m.get()} for c, m, *_ in self.produtos],
            "frete_internacional": self.frete_var.get(),
            "armazenagem": self.arm_var.get(),
            "despachante": self.desp_var.get(),
            "fornecedor_codigo": self.fornecedor_cod_str.get(),
            "exportador_codigo": self.exportador_cod_str.get(),
        }

        def _apportion_to_cents(total, weights):
            total = float(total or 0.0)
            n = len(weights)
            if n == 0:
                return []
            sum_w = sum(weights)
            if sum_w <= 0:
                base = int(round(total * 100.0))
                q, r = divmod(base, n)
                cents = [q] * n
                for i in range(r):
                    cents[i] += 1
                return [c / 100.0 for c in cents]
            raw = [(total * (w / sum_w)) for w in weights]
            floor_cents = [int(x * 100 // 1) for x in raw]
            remainder = int(round(total * 100)) - sum(floor_cents)
            order = sorted(range(n), key=lambda i: (raw[i] * 100 - floor_cents[i]), reverse=True)
            for i in range(max(0, remainder)):
                floor_cents[order[i % n]] += 1
            return [c / 100.0 for c in floor_cents]

        if self.pdf_path.get() and PDF_OK:
            try:
                dados_pdf = parse_pdf_fields(self.pdf_path.get(), modelo=self.modelo.get())
                dados.update(dados_pdf)
                texto_pdf = _extract_text(self.pdf_path.get())

                if self.tipo.get() == "Padrão":
                    qtd_prod_mmnp = dados.get("produtos_qtd_por_mmnp")
                    if isinstance(qtd_prod_mmnp, int) and qtd_prod_mmnp > 0:
                        self._ensure_prod_rows(qtd_prod_mmnp)

                if self.tipo.get() == "BAP":
                    self._apply_bap_auto_products(dados, texto_pdf)

                # Heurísticas
                qtd_seal = len(re.findall(r"\bSEAL\b", texto_pdf, flags=re.I))
                dados["transporte"] = qtd_seal * 1200

                qtd_adicoes_final = dados.get("quantidade_adicoes") or len(dados.get("adicoes") or [])
                dados["inmetro"] = qtd_adicoes_final * 53.53

                cif_val = (dados.get("moeda_real") or {}).get("cif") or 0.0

                try:
                    arm = float(Decimal(dados["armazenagem"].replace(".", "").replace(",", ".")))
                    desp = float(Decimal(dados["despachante"].replace(".", "").replace(",", ".")))
                except Exception:
                    arm, desp = 0.0, 0.0

                # PIS/COFINS total suspenso
                base_piscofins = dados.get("piscofins_base_calculo_rs") or 0.0
                pis_aliq = dados.get("pis_aliquota_percent") or 0.0
                cofins_aliq = dados.get("cofins_aliquota_percent") or 0.0
                if dados.get("piscofins_regime", "").strip().upper() in ["SUSPENSO", "SUSPENSAO"]:
                    pis_val = 0.0
                    cofins_val = 0.0
                else:
                    pis_val = base_piscofins * (pis_aliq / 100.0)
                    cofins_val = base_piscofins * (cofins_aliq / 100.0)

                siscomex = dados.get("taxa_utilizacao_siscomex_rs") or 0.0

                # ==== Rateio ====
                adicoes = dados.get("adicoes") or []
                n_itens = len(adicoes)

                pesos_raw, qtds_raw = [], []
                for it in adicoes:
                    w = it.get("peso_liquido_kg") or 0.0
                    q = it.get("descricao_qtde") or 0.0
                    try:
                        q = float(q)
                    except Exception:
                        q = 0.0
                    pesos_raw.append(float(w) if w else 0.0)
                    qtds_raw.append(q)

                total_peso_pdf = dados.get("carga_peso_liquido_kg") or 0.0
                criterio_rateio = "igual"

                # === Pesos da Invoice (cruzamento por código de produto) ===
                pesos_invoice = self._ler_pesos_invoice()
                usando_invoice = False

                if pesos_invoice and adicoes:
                    # Conta quantas vezes cada código aparece nas adições
                    from collections import Counter
                    contagem_cod = Counter(
                        (it.get("produto_codigo") or "").strip().upper()
                        for it in adicoes
                    )

                    pesos_invoice_list = []
                    todos_encontrados = True
                    for it in adicoes:
                        cod = (it.get("produto_codigo") or "").strip().upper()
                        peso_inv = pesos_invoice.get(cod)
                        if peso_inv is not None:
                            # Divide o peso pelo número de itens com esse código
                            n_ocorrencias = contagem_cod.get(cod, 1) or 1
                            pesos_invoice_list.append(peso_inv / n_ocorrencias)
                        else:
                            pesos_invoice_list.append(None)
                            todos_encontrados = False

                    # Aplica pesos da invoice onde encontrou; usa proporcional nos demais
                    if any(p is not None for p in pesos_invoice_list):
                        for idx_it, it in enumerate(adicoes):
                            if pesos_invoice_list[idx_it] is not None:
                                it["peso_liquido_kg"] = pesos_invoice_list[idx_it]
                        usando_invoice = True
                        criterio_rateio = "peso(invoice)" if todos_encontrados else "peso(invoice+proporcional)"
                        # Atualiza pesos_raw com os valores já corrigidos
                        pesos_raw = [float(it.get("peso_liquido_kg") or 0.0) for it in adicoes]

                        # Registra no dados para exibição
                        dados["invoice_usada"] = True
                        dados["invoice_arquivo"] = os.path.basename(self.invoice_path.get())
                        nao_encontrados = [
                            (it.get("produto_codigo") or "?")
                            for it, p in zip(adicoes, pesos_invoice_list) if p is None
                        ]
                        dados["invoice_nao_encontrados"] = nao_encontrados

                if sum(pesos_raw) > 0:
                    pesos_final = pesos_raw[:]
                    if not usando_invoice:
                        criterio_rateio = "peso"
                    # Preenche faltantes proporcionalmente (só quando NÃO veio da invoice)
                    if not usando_invoice and total_peso_pdf and total_peso_pdf > 0:
                        soma_presentes = sum(w for w in pesos_raw if w > 0)
                        faltantes_idx = [i for i, w in enumerate(pesos_raw) if w <= 0]
                        resto = max(0.0, total_peso_pdf - soma_presentes)
                        if faltantes_idx and resto > 0:
                            soma_qtd_falt = sum(qtds_raw[i] for i in faltantes_idx)
                            for i in faltantes_idx:
                                if soma_qtd_falt > 0:
                                    pesos_final[i] = (qtds_raw[i] / soma_qtd_falt) * resto
                                else:
                                    pesos_final[i] = resto / len(faltantes_idx)
                            criterio_rateio = "peso(pdf)+qtd(faltantes)"
                elif sum(qtds_raw) > 0:
                    pesos_final = qtds_raw[:]
                    criterio_rateio = "quantidade"
                else:
                    pesos_final = [1.0] * n_itens

                sum_calc = sum(pesos_final) or (n_itens or 1.0)
                p_calc = [(w / sum_calc) for w in pesos_final]
                pesos_int = [int(round(w)) for w in pesos_final]
                soma_int = sum(pesos_int) or (n_itens or 1)
                p_view = [(wi / soma_int) for wi in pesos_int]

                dados["criterio_rateio"] = criterio_rateio

                # ==== Totais MOEDA ====
                try:
                    frete_int_total = float(Decimal(dados["frete_internacional"].replace(".", "").replace(",", ".")))
                except Exception:
                    frete_int_total = 0.0

                # Frete, acrescimo e seguro rateados SEPARADAMENTE para evitar
                # erro de centavos por arredondamento em cima de total combinado.
                frete_somente   = (dados.get("moeda_real") or {}).get("frete") or 0.0
                acrescimo_total = (dados.get("moeda_real") or {}).get("acrescimo") or 0.0
                frete_pdf_total = frete_somente + acrescimo_total  # para exibicao/diferenca
                dados.setdefault("moeda_real", {})["frete_com_acrescimo"] = frete_pdf_total

                seguro_total = (dados.get("moeda_real") or {}).get("seguro") or 0.0
                despesas_di_total = frete_pdf_total + seguro_total
                dados["despesas_di_total"] = despesas_di_total

                diferenca_frete_total = frete_int_total - frete_pdf_total
                dados["diferenca_frete"] = diferenca_frete_total

                siscomex_total = siscomex
                inmetro_total = dados.get("inmetro") or 0.0
                arm_total = arm
                desp_total = desp
                transporte_total = dados.get("transporte") or 0.0

                # Rateios financeiros - frete e acrescimo separados para precisao de centavos
                frete_somente_parts = _apportion_to_cents(frete_somente,   pesos_final)
                acrescimo_parts     = _apportion_to_cents(acrescimo_total,  pesos_final)
                frete_pdf_parts     = [frete_somente_parts[i] + acrescimo_parts[i] for i in range(len(pesos_final))]
                frete_int_parts     = _apportion_to_cents(frete_int_total,  pesos_final)
                seguro_parts        = _apportion_to_cents(seguro_total,     pesos_final)
                siscomex_parts      = _apportion_to_cents(siscomex_total,   pesos_final)

                dados["debug_pesos"] = {
                    "sum_pesos_final": round(sum(pesos_final), 3),
                    "n_itens": n_itens,
                    "primeiros_5_pesos": [round(p, 4) for p in pesos_final[:5]],
                    "frete_somente": round(frete_somente, 2),
                    "acrescimo_total": round(acrescimo_total, 2),
                    "frete_pdf_total": round(frete_pdf_total, 2),
                    "seguro_total": round(seguro_total, 2),
                    "primeiros_5_frete_parts": [round(x, 4) for x in frete_pdf_parts[:5]],
                }

                taxa_usd = dados.get("taxa_usd") or 0.0
                itens_detalhe = []
                icms_total = 0.0
                fti_total = 0.0
                pis_total = pis_val
                cofins_total = cofins_val

                prod_list = dados.get("produtos") or []

                for idx, it in enumerate(adicoes):
                    p_i_calc = p_calc[idx] if idx < len(p_calc) else 0.0
                    p_i_view = p_view[idx] if idx < len(p_view) else 0.0

                    # Produto / MVA
                    prod_codigo = "-"
                    mva_val = 0.0

                    if self.tipo.get() == "BAP":
                        # Código vem do item BAP ou da UI
                        prod_codigo = (it.get("produto_codigo") or "")
                        if idx < len(prod_list):
                            ui_cod = (prod_list[idx].get("codigo") or "").strip()
                            if ui_cod:
                                prod_codigo = ui_cod

                        if dados.get("modelo_estado") == "Roraima":
                            # Roraima: MVA individual por produto (campo mva da UI)
                            if idx < len(prod_list):
                                mva_str = (prod_list[idx].get("mva") or "0").strip()
                                try:
                                    mva_val = float(Decimal(mva_str.replace(".", "").replace(",", ".")))
                                except Exception:
                                    mva_val = 0.0
                            else:
                                mva_val = 0.0
                        else:
                            # Amazonas e demais: MVA comum para todos
                            try:
                                mva_val = float(Decimal((self.mva_comum_var.get() or "0").strip().replace(".", "").replace(",", ".")))
                            except Exception:
                                mva_val = 0.0
                    else:
                        if idx < len(prod_list):
                            prod_codigo = (prod_list[idx].get("codigo") or "-")
                            mva_str = (prod_list[idx].get("mva") or "0").strip()
                            try:
                                mva_val = float(Decimal(mva_str.replace(".", "").replace(",", ".")))
                            except Exception:
                                mva_val = 0.0

                    # FOB do item = VUCV (USD) × Qtde × Taxa
                    vucv_usd = it.get("descricao_vucv_usd") or 0.0
                    qtde_i = it.get("descricao_qtde") or 0.0
                    fob_i = taxa_usd * vucv_usd * qtde_i

                    frete_pdf_i = frete_pdf_parts[idx] if idx < len(frete_pdf_parts) else 0.0
                    frete_int_i = frete_int_parts[idx] if idx < len(frete_int_parts) else 0.0
                    seguro_i = seguro_parts[idx] if idx < len(seguro_parts) else 0.0
                    despesas_di_i = frete_pdf_i + seguro_i
                    valor_mercadoria_i = fob_i + despesas_di_i
                    siscomex_i = siscomex_parts[idx] if idx < len(siscomex_parts) else 0.0
                    base_icms_item = fob_i + frete_pdf_i + seguro_i + siscomex_i
                    fti_i = 0.01 * valor_mercadoria_i
                    fti_total += fti_i

                    icms_item = base_icms_item * (mva_val / 100.0)
                    icms_total += icms_item

                    inmetro_i = inmetro_total * p_i_calc
                    arm_i = arm_total * p_i_calc
                    desp_i = desp_total * p_i_calc
                    transporte_i = transporte_total * p_i_calc
                    diferenca_frete_i = diferenca_frete_total * p_i_calc
                    pis_val_item = pis_total * p_i_calc
                    cofins_val_item = cofins_total * p_i_calc

                    cif_i = cif_val * p_i_calc
                    outros_i = (
                        pis_val_item + cofins_val_item + siscomex_i
                        + inmetro_i + arm_i + desp_i + fti_i + transporte_i
                    )
                    vtn_item = (
                        base_icms_item + icms_item + diferenca_frete_i
                        + inmetro_i + arm_i + desp_i + fti_i + transporte_i
                    )

                    preco_unit_rs = (fob_i / qtde_i) if qtde_i else 0.0
                    custo_unit_rs = (vtn_item / qtde_i) if qtde_i else 0.0

                    itens_detalhe.append({
                        "adicao": it.get("adicao"),
                        "produto_codigo": prod_codigo,
                        "mva_percent": mva_val,
                        "peso_kg": it.get("peso_liquido_kg"),
                        "peso_perc": p_i_view * 100.0,
                        "fob_item": fob_i,
                        "frete_pdf_item": frete_pdf_i,
                        "frete_int_item": frete_int_i,
                        "seguro_item": seguro_i,
                        "despesas_di_item": despesas_di_i,
                        "valor_mercadoria_item": valor_mercadoria_i,
                        "siscomex_item": siscomex_i,
                        "base_icms_item": base_icms_item,
                        "icms_item": icms_item,
                        "fti": fti_i,
                        "inmetro": inmetro_i,
                        "armazenagem": arm_i,
                        "despachante": desp_i,
                        "transporte": transporte_i,
                        "diferenca_frete_item": diferenca_frete_i,
                        "pis_val": pis_val_item,
                        "cofins_val": cofins_val_item,
                        "cif_item": cif_i,
                        "outros_item": outros_i,
                        "vtn_item": vtn_item,
                        "preco_unit_rs": preco_unit_rs,
                        "custo_unit_rs": custo_unit_rs,
                        "qtde_item": qtde_i,
                    })

                dados["fti"] = fti_total
                dados["outros"] = (
                    pis_total + cofins_total + siscomex_total
                    + inmetro_total + arm_total + desp_total + fti_total + transporte_total
                )
                dados["icms_pago"] = icms_total
                dados["itens_detalhe"] = itens_detalhe
                dados["valor_total_nota"] = (
                    (cif_val or 0.0) + siscomex + icms_total
                    + diferenca_frete_total + (dados["outros"] - siscomex)
                )

            except Exception as e:
                import traceback
                messagebox.showerror("Erro ao ler PDF", f"Falha ao processar o PDF:\n{e}\n\n{traceback.format_exc()}")
                dados.setdefault("inmetro", 53.53)
                dados.setdefault("transporte", 1200.00)
                dados.setdefault("outros", 1.00)
                dados.setdefault("fti", 1.00)
        from collections import Counter
        contagem = Counter((it.get("produto_codigo") or "").strip().upper() for it in adicoes)
        print(f"HE26 aparece {contagem.get('HE26')} vezes nas adições")
        print(f"peso invoice HE26 = {pesos_invoice.get('HE26')}")



        self._dados_extracao = dados
        self._render_saida(dados)

    # ===== Exportação XLSX =====
    def exportar_xlsx(self):
        conteudo = getattr(self, "_dados_extracao", None)
        if not conteudo:
            messagebox.showinfo("Sem conteúdo", "Execute a leitura antes de exportar.")
            return

        itens = conteudo.get("itens_detalhe") or []
        produtos = conteudo.get("produtos") or []
        max_len = max(len(itens), len(produtos)) if (itens or produtos) else 0

        uf = ESTADO_TO_UF.get(conteudo.get("modelo_estado"), "")
        empresa_produto = "1"
        fornecedor_cod = (conteudo.get("fornecedor_codigo") or "").upper()
        exportador_cod = (conteudo.get("exportador_codigo") or "").upper()
        numero_di = (conteudo.get("di") or "").upper()
        data_desembaraco = (conteudo.get("data") or "").upper()
        local_desembaraco = "MANAUS"
        cnpj_cd = re.sub(r"\D", "", conteudo.get("cnpj_importador") or "")

        colunas = [
            "EMPRESA PRODUTO", "FORNECEDOR", "PRODUTO", "QUANTIDADE",
            "PREÇO UNITÁRIO", "BASE DE CALCULO ICMS", "MVA", "ICMS ST",
            "NÚMERO DA DI", "DATA_DESEMBARACO", "LOCA_DESEMBARACO", "UF",
            "VIA DE TRANSPORTE", "TIPO INTERMEDIARIO", "CNPJ CD", "UF",
            "CODIGO DO EXPORTADOR", "NUM", "SEQ", "FABRICANTE",
        ]

        def fmt2(x):
            try:
                v = float(x)
                return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return "0,00"

        def fmt0(x):
            try:
                return str(int(round(float(x))))
            except Exception:
                return "0"

        def clean_fab(s):
            if not s:
                return ""
            s = s.upper()
            s = re.sub(r"\bNOME\s*:\s*", "", s)
            return s.strip()

        fabricante = clean_fab(conteudo.get("fabricante") or "")

        data_rows = []
        for i in range(max_len):
            it = itens[i] if i < len(itens) else {}
            prod = produtos[i] if i < len(produtos) else {"mva": "0", "codigo": ""}

            quantidade = it.get("qtde_item") or 0
            preco_unit = it.get("preco_unit_rs") or 0.0
            base_icms = it.get("base_icms_item") or 0.0

            is_bap = (conteudo.get("tipo") or "").upper() == "BAP"
            is_roraima = (conteudo.get("modelo_estado") or "") == "Roraima"

            if is_bap and not is_roraima:
                # Amazonas BAP: MVA comum para todos
                try:
                    mva_original = float(str(self.mva_comum_var.get() or 0).replace('.', '').replace(',', '.'))
                except Exception:
                    mva_original = 0.0
            else:
                # Padrão ou Roraima BAP: MVA individual por produto
                try:
                    mva_original = float(str(prod.get("mva") or 0).replace('.', '').replace(',', '.'))
                except Exception:
                    mva_original = float(it.get("mva_percent") or 0.0)

            mva_decimal = mva_original / 100.0
            icms_st_val = it.get("icms_item") or (base_icms * mva_decimal)

            produto_codigo = (it.get("produto_codigo") or prod.get("codigo") or "").upper()

            row = [
                empresa_produto, fornecedor_cod, produto_codigo,
                fmt0(quantidade), fmt2(preco_unit), fmt2(base_icms),
                fmt2(mva_original), fmt2(icms_st_val),
                numero_di, data_desembaraco, local_desembaraco, uf,
                "1", "1", cnpj_cd, uf, exportador_cod, "1", exportador_cod,
                fabricante,
            ]
            data_rows.append(row)

        sufixo = datetime.now().strftime("%Y%m%d_%H%M%S")
        if PANDAS_OK:
            try:
                df = pd.DataFrame(data_rows, columns=colunas)
                df = df.astype(str)
                caminho = filedialog.asksaveasfilename(
                    title="Salvar como",
                    defaultextension=".xlsx",
                    initialfile=f"tabela_produtos_{sufixo}.xlsx",
                    filetypes=[("Planilha Excel", "*.xlsx")],
                )
                if not caminho:
                    return
                df.to_excel(caminho, index=False, engine="openpyxl")
                messagebox.showinfo("Exportação concluída", f"Arquivo salvo em:\n{caminho}")
            except Exception as e:
                messagebox.showerror("Erro ao exportar", f"Erro ao gerar XLSX:\n{e}")
        else:
            import csv
            caminho = filedialog.asksaveasfilename(
                title="Salvar como (CSV - fallback)",
                defaultextension=".csv",
                initialfile=f"tabela_produtos_{sufixo}.csv",
                filetypes=[("CSV", "*.csv")],
            )
            if not caminho:
                return
            try:
                with open(caminho, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(colunas)
                    for r in data_rows:
                        writer.writerow(r)
                messagebox.showinfo("Exportação concluída (CSV)", f"Arquivo salvo em:\n{caminho}")
            except Exception as e:
                messagebox.showerror("Erro ao exportar", f"Falha ao salvar CSV:\n{e}")

    # ===== Renderização =====
    def _render_saida(self, d: dict):
        def fmt_num(n):
            if n is None:
                return "-"
            try:
                return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return str(n)

        base_piscofins = d.get("piscofins_base_calculo_rs") or 0
        pis_aliq = d.get("pis_aliquota_percent") or 0
        cofins_aliq = d.get("cofins_aliquota_percent") or 0
        if d.get("piscofins_regime", "").strip().upper() in ["SUSPENSO", "SUSPENSAO"]:
            pis_val = cofins_val = 0
        else:
            pis_val = base_piscofins * (pis_aliq / 100)
            cofins_val = base_piscofins * (cofins_aliq / 100)

        linhas = [
            "===== Resultado =====",
            f"Data: {d.get('data')}",
            f"DI: {d.get('di')}",
            f"Tipo: {d.get('tipo')}",
            f"Modelo: {d.get('modelo_estado')}",
            f"Arquivo: {d.get('arquivo_pdf')}",
            f"Fornecedor: {d.get('fornecedor_codigo', '')}",
            f"Exportador: {d.get('exportador_codigo', '')}",
        ]

        if d.get("invoice_usada"):
            linhas.append(f"Invoice: {d.get('invoice_arquivo', '')}  ✔ pesos reais aplicados")
            nao_enc = d.get("invoice_nao_encontrados") or []
            if nao_enc:
                linhas.append(f"  ⚠ Códigos não encontrados na invoice (peso proporcional): {', '.join(nao_enc)}")
        elif self.invoice_path.get():
            linhas.append("Invoice: (nenhum produto cruzado — usando peso proporcional)")

        if d.get("tipo", "").upper() == "BAP":
            linhas += [f"MVA (comum): {self.mva_comum_var.get()}%"]

        linhas += [
            f"Produtos (MM/N/POL.): {d.get('produtos_qtd_por_mmnp', '-')}",
            f"Itens BAP expandidos: {len(d.get('adicoes') or [])}",
            "",
            "Despesas:",
            f"• Frete Internacional: R$ {d.get('frete_internacional')}",
            f"• Armazenagem: R$ {d.get('armazenagem')}",
            f"• Despachante: R$ {d.get('despachante')}",
            f"• Inmetro: R$ {fmt_num(d.get('inmetro'))}",
            f"• Transporte: R$ {fmt_num(d.get('transporte'))}",
            f"• FTI: R$ {fmt_num(d.get('fti'))}",
            f"• Siscomex: R$ {fmt_num(d.get('taxa_utilizacao_siscomex_rs'))}",
            f"• Acréscimo (NA MOEDA): R$ {fmt_num((d.get('moeda_real') or {}).get('acrescimo'))}",
            f"• Despesas DI total: R$ {fmt_num(d.get('despesas_di_total'))}",
            f"• Outros: R$ {fmt_num(d.get('outros'))}",
            f"• ICMS Pago: R$ {fmt_num(d.get('icms_pago'))}",
            f"• Diferença de Frete: R$ {fmt_num(d.get('diferenca_frete'))}",
            f"• Valor Total da Nota: R$ {fmt_num(d.get('valor_total_nota'))}",
            "",
        ]

        itens_det = d.get("itens_detalhe") or []
        if itens_det:
            linhas.append(f"Detalhamento ({d.get('criterio_rateio', 'peso')}) – {len(itens_det)} itens:")
            for it in itens_det:
                linhas.append(
                    f"  [{it.get('adicao')}] {it.get('produto_codigo')} "
                    f"Qtde:{fmt_num(it.get('qtde_item'))} "
                    f"FOB:R${fmt_num(it.get('fob_item'))} "
                    f"Merc:R${fmt_num(it.get('valor_mercadoria_item'))} "
                    f"BaseICMS:R${fmt_num(it.get('base_icms_item'))} "
                    f"ICMS:R${fmt_num(it.get('icms_item'))} "
                    f"VTN:R${fmt_num(it.get('vtn_item'))} "
                    f"PUnit:R${fmt_num(it.get('preco_unit_rs'))}"
                )
            linhas.append("")

        dbg = d.get("debug_pesos") or {}
        if dbg:
            linhas += [
                "",
                "DEBUG Pesos:",
                f"• n_itens: {dbg.get('n_itens')}",
                f"• soma pesos_final: {dbg.get('sum_pesos_final')}",
                f"• primeiros 5 pesos: {dbg.get('primeiros_5_pesos')}",
                f"• frete_pdf_total: {dbg.get('frete_pdf_total')}",
                f"• seguro_total: {dbg.get('seguro_total')}",
                f"• primeiros 5 frete_parts: {dbg.get('primeiros_5_frete_parts')}",
            ]

        linhas += [
            "PDF – Extração:",
            f"• Qtd. Adições: {d.get('quantidade_adicoes', '-')}",
            f"• Carga Qtde: {d.get('carga_quantidade', '-')}",
            f"• Peso Líquido Total (Kg): {fmt_num(d.get('carga_peso_liquido_kg'))}",
            "",
            "MOEDA (R$):",
            f"• FOB: {fmt_num((d.get('moeda_real') or {}).get('fob'))}",
            f"• Frete+Acréscimo: {fmt_num((d.get('moeda_real') or {}).get('frete_com_acrescimo'))}",
            f"• Seguro: {fmt_num((d.get('moeda_real') or {}).get('seguro'))}",
            f"• CIF: {fmt_num((d.get('moeda_real') or {}).get('cif'))}",
            "",
            f"Taxa US$: {fmt_num(d.get('taxa_usd'))} ({d.get('taxa_usd_data', '-')})",
            f"Siscomex: R$ {fmt_num(d.get('taxa_utilizacao_siscomex_rs'))}",
            "",
            f"II: {d.get('ii_regime', '-')} {fmt_num(d.get('ii_aliquota_percent'))}%",
            f"IPI: {d.get('ipi_regime', '-')} {fmt_num(d.get('ipi_aliquota_percent'))}%",
            f"PIS/COFINS base: R$ {fmt_num(base_piscofins)} regime: {d.get('piscofins_regime', '-')}",
            f"PIS: {fmt_num(pis_aliq)}% = R$ {fmt_num(pis_val)}",
            f"COFINS: {fmt_num(cofins_aliq)}% = R$ {fmt_num(cofins_val)}",
        ]

        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", "\n".join(linhas))

    def _insert_hint(self):
        self.txt.insert(
            "1.0",
            "Dica: selecione o PDF, informe DI, Fornecedor, Exportador e Modelo.\n"
            "No BAP, os produtos são extraídos automaticamente (um por item de descrição).\n"
            "Use MVA (comum) para aplicar alíquota ICMS-ST a todos os itens.\n"
            "Depois, Executar → Exportar XLSX.\n"
        )


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    app = App(parent=root)
    app.mainloop()