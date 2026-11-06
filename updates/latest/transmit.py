import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys

from core import theme as _t   # identidade visual do Prisma (core/theme.py)

CNPJ_PADRAO = "CNPJ Dest att.xlsx"


def _achar_tabela_cnpj():
    """Procura a tabela CNPJ->UF distribuída com o app (ao lado do .exe,
    dentro do bundle ou na pasta do projeto). Retorna None se não existir."""
    pastas = []
    if getattr(sys, "frozen", False):
        pastas.append(os.path.dirname(sys.executable))
        if hasattr(sys, "_MEIPASS"):
            pastas.append(sys._MEIPASS)
    pastas.append(os.path.dirname(os.path.abspath(__file__)))
    for pasta in pastas:
        caminho = os.path.join(pasta, CNPJ_PADRAO)
        if os.path.isfile(caminho):
            return caminho
    return None

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    import xlsxwriter
except ImportError as _e:
    # NAO tente instalar via pip aqui: em um .exe compilado (PyInstaller),
    # sys.executable e o proprio programa, entao chamar
    # "sys.executable -m pip install ..." acaba executando o proprio app
    # de novo (e por isso parece que ele "volta ao menu"). Se isso disparar,
    # e sinal de que a dependencia nao foi empacotada corretamente no build.
    raise ImportError(
        f"Dependencia ausente no executavel: {_e}. "
        "Reempacote o programa incluindo pandas, openpyxl e xlsxwriter "
        "(ex.: pyinstaller --hidden-import=pandas --hidden-import=openpyxl "
        "--hidden-import=xlsxwriter)."
    ) from _e

# ── mapeamento de grupos CST ──────────────────────────────────────────────────
CST_GROUPS = [
    {"cst": "00", "origem": "Origem ICMS 00", "cst_col": "CST ICMS 00",
     "bc": "Vlr BC ICMS 00", "aliq": "Alíquota ICMS 00", "vlr": "Vlr ICMS 00",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "02", "origem": "Origem ICMS 02", "cst_col": "CST ICMS 02",
     "bc": "Vlr BC ICMS 02", "aliq": "Alíquota ICMS 02", "vlr": "Vlr ICMS 02",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "10", "origem": "Origem ICMS 10", "cst_col": "CST ICMS 10",
     "bc": "Vlr BC ICMS 10", "aliq": "Alíquota ICMS 10", "vlr": "Vlr ICMS 10",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "15", "origem": "Origem ICMS 15", "cst_col": "CST ICMS 15",
     "bc": None, "aliq": None, "vlr": "Vlr ICMS Mono 15",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "20", "origem": "Origem ICMS 20", "cst_col": "CST ICMS 20",
     "bc": "Vlr BC ICMS 20", "aliq": "Alíquota ICMS 20", "vlr": "Vlr ICMS 20",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "30", "origem": "Origem ICMS 30", "cst_col": "CST ICMS 30",
     "bc": None, "aliq": None, "vlr": None,
     "bc_st": "Vlr BC ST ICMS 30", "aliq_st": "Alíquota ST ICMS 30", "vlr_st": "Vlr ST ICMS 30"},
    {"cst": "40", "origem": "Origem ICMS 40", "cst_col": "CST ICMS 40",
     "bc": None, "aliq": None, "vlr": None,
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "51", "origem": "Origem ICMS 51", "cst_col": "CST ICMS 51",
     "bc": "Vlr BC ICMS 51", "aliq": "Alíquota ICMS 51", "vlr": "Vlr ICMS 51",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "53", "origem": "Origem ICMS 53", "cst_col": "CST ICMS 53",
     "bc": None, "aliq": None, "vlr": "Vlr mono op ICMS 53",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "60", "origem": "Origem ICMS 60", "cst_col": "CST ICMS 60",
     "bc": "Vlr BC ST ret ICMS 60", "aliq": "Alíquota ST ICMS 60", "vlr": None,
     "bc_st": "Vlr BC ST ret ICMS 60", "aliq_st": "Alíquota ST ICMS 60", "vlr_st": "Vlr subst ICMS 60"},
    {"cst": "61", "origem": "Origem ICMS 61", "cst_col": "CST ICMS 61",
     "bc": None, "aliq": None, "vlr": "Vlr mono ret ICMS 61",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "70", "origem": "Origem ICMS 70", "cst_col": "CST ICMS 70",
     "bc": "Vlr BC ICMS 70", "aliq": "Alíquota ICMS 70", "vlr": "Vlr ICMS 70",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
    {"cst": "90", "origem": "Origem ICMS 90", "cst_col": "CST ICMS 90",
     "bc": "Vlr BC ICMS 90", "aliq": "Alíquota ICMS 90", "vlr": "Vlr ICMS 90",
     "bc_st": None, "aliq_st": None, "vlr_st": None},
]

MODELO_COLS = [
    "Chave NF-e", "Status da Nota", "Nat. Operação", "Série NF", "Num NF",
    "UF Emitente", "CNPJ Emitente","Nome Emitente","UF Destinatário", "CNPJ Dest", "Nome Dest",
    "Data Emissão", "Vlr Unit Trib", "Item Nº", "Código Prod", "Descrição Prod",
    "NCM", "CEST", "CNPJ Fabr. Prod.", "EX TIPI", "CFOP", "Unid Comercial",
    "Qtd Comercial", "Vlr Unit Com", "Vlr Total Item", "Unid Trib", "Qtd Trib",
    "Vlr Frete", "Vlr Seguro", "Vlr Desc", "Vlr Outros", "Vlr Tributos",
    "Origem ICMS", "CST ICMS", "Vlr BC ICMS", "Alíquota ICMS", "Vlr ICMS",
    "Vlr BC ST", "Alíquota ST ICMS", "Vlr total ICMS ST","Vlr IPI", "Vlr ICMS Deson", "Valor IPITrib",
]


def get_val(row, col):
    if col is None:
        return None
    v = row.get(col)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return v


def processar(df, log_fn, cnpj_map=None):
    cols = set(df.columns)
    active_groups = [g for g in CST_GROUPS if g["cst_col"] in cols]
    csts_found = [g["cst"] for g in active_groups]
    log_fn(f"Grupos CST detectados: {', '.join(csts_found)}")

    CAMPOS_ESPECIAIS = {
        "UF Destinatário",
        "Origem ICMS", "CST ICMS", "Vlr BC ICMS", "Alíquota ICMS", "Vlr ICMS",
        "Vlr BC ST ICMS", "Alíquota ST ICMS", "Vlr ST ICMS", "Vlr ICMS Deson",
    }

    # campos diretos — copia vetorialmente
    df_out = pd.DataFrame(index=df.index)
    for c in MODELO_COLS:
        if c not in CAMPOS_ESPECIAIS:
            df_out[c] = df[c] if c in cols else None

    # UF Destinatário — lookup vetorizado
    if cnpj_map and "CNPJ Dest" in cols:
        serie = df["CNPJ Dest"].astype(str).str.strip().str.replace(r'\D', '', regex=True)
        df_out["UF Destinatário"] = serie.map(cnpj_map)
        sem_uf = serie[df_out["UF Destinatário"].isna() & serie.notna() & (serie != "None")].unique()
        if len(sem_uf):
            amostra = ", ".join(sorted(sem_uf)[:5])
            sufixo = "..." if len(sem_uf) > 5 else ""
            log_fn(f"⚠ {len(sem_uf)} CNPJ(s) sem UF: {amostra}{sufixo}")
    else:
        df_out["UF Destinatário"] = None

    # campos ICMS por CST — combina em cascata vetorialmente
    def build_icms_col(field):
        result = pd.Series(None, index=df.index, dtype=object)
        for g in active_groups:
            src = g.get(field)
            if not src or src not in cols:
                continue
            mask = df[g["cst_col"]].notna() & (df[g["cst_col"]].astype(str).str.strip() != "")
            result = result.where(~mask, df[src])
        return result

    df_out["Origem ICMS"]      = build_icms_col("origem")
    df_out["CST ICMS"]         = build_icms_col("cst_col")
    df_out["Vlr BC ICMS"]      = build_icms_col("bc")
    df_out["Alíquota ICMS"]    = build_icms_col("aliq")
    df_out["Vlr ICMS"]         = build_icms_col("vlr")
    df_out["Vlr BC ST ICMS"]   = build_icms_col("bc_st")
    df_out["Alíquota ST ICMS"] = build_icms_col("aliq_st")
    df_out["Vlr ST ICMS"]      = build_icms_col("vlr_st")
    df_out["Vlr ICMS Deson"]   = build_icms_col("deson")

    df_out = df_out[MODELO_COLS]
    log_fn(f"Linhas processadas: {len(df_out):,}")
    return df_out, csts_found


def salvar_xlsx(df_out, caminho, log_fn):
    COR_HEADER   = "#1F5FA6"
    COR_ICMS_HDR = "#2E7D32"
    COR_LINHA_A  = "#EEF3FB"
    COR_LINHA_B  = "#FFFFFF"

    icms_start = MODELO_COLS.index("Origem ICMS")  # 0-based

    wb = xlsxwriter.Workbook(caminho, {"constant_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("Modelo")

    # formatos
    hdr_base = {"bold": True, "font_color": "#FFFFFF", "font_name": "Arial",
                "font_size": 9, "align": "center", "valign": "vcenter",
                "text_wrap": True, "border": 1, "border_color": "#BDBDBD"}
    fmt_hdr_azul  = wb.add_format({**hdr_base, "bg_color": COR_HEADER})
    fmt_hdr_verde = wb.add_format({**hdr_base, "bg_color": COR_ICMS_HDR})

    cell_base = {"font_name": "Arial", "font_size": 9,
                 "border": 1, "border_color": "#BDBDBD", "valign": "vcenter"}
    fmt_a = wb.add_format({**cell_base, "bg_color": COR_LINHA_A})
    fmt_b = wb.add_format({**cell_base, "bg_color": COR_LINHA_B})

    # cabeçalho
    for ci, col_name in enumerate(MODELO_COLS):
        fmt = fmt_hdr_verde if ci >= icms_start else fmt_hdr_azul
        ws.write(0, ci, col_name, fmt)
        ws.set_column(ci, ci, min(max(len(col_name), 10) + 2, 30))

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, 0, len(MODELO_COLS) - 1)

    # dados
    for ri, row in enumerate(df_out.itertuples(index=False), start=1):
        fmt = fmt_a if ri % 2 == 0 else fmt_b
        for ci, val in enumerate(row):
            if val is None or (isinstance(val, float) and (pd.isna(val) or val != val)):
                val = ""
            ws.write(ri, ci, val, fmt)

    wb.close()
    log_fn(f"Arquivo salvo: {os.path.basename(caminho)}")


# ── interface ─────────────────────────────────────────────────────────────────
# ── interface ─────────────────────────────────────────────────────────────────
class TransmitFrame(tk.Frame):
    """Frame embutível — usado dentro de um Toplevel do menu principal.
    Mantém sua própria paleta fixa de cores (mesmo padrão dos módulos
    Cofre/Brinks/Prossegur), sem depender do tema dinâmico T()."""

    def __init__(self, master):
        super().__init__(master, bg=_t.BG_APP)
        self._df_out = None
        self._build()

    def _secao(self, titulo: str) -> tk.LabelFrame:
        """Quadro de seção no estilo do Prisma: borda fina e título ciano."""
        return tk.LabelFrame(self, text=f" {titulo} ", font=("Segoe UI", 9, "bold"),
                             bg=_t.BG_APP, fg=_t.ACCENT_CYAN, bd=0, relief="flat",
                             highlightthickness=1, highlightbackground=_t.GRID_LINE,
                             highlightcolor=_t.GRID_LINE)

    def _botao(self, parent, texto: str, comando, bg: str, fg: str = "white", **kw) -> tk.Button:
        return tk.Button(parent, text=texto, font=("Segoe UI", 9, "bold"), bg=bg, fg=fg,
                         activebackground=_t.PRIMARY_HOVER, activeforeground="white",
                         relief="flat", bd=0, cursor="hand2", padx=14, pady=5,
                         command=comando, **kw)

    def _build(self):
        PAD = {"padx": 20, "pady": 10}

        # ── título ──
        frm_title = tk.Frame(self, bg=_t.CARD_BG)
        frm_title.pack(fill="x")
        tk.Label(frm_title, text="Conversor NF-e", font=("Segoe UI", 14, "bold"),
                 bg=_t.CARD_BG, fg=_t.FG_TEXT).pack(side="left", padx=20, pady=14)
        tk.Label(frm_title, text="Banco Geral → Modelo ICMS Unificado",
                 font=("Segoe UI", 9), bg=_t.CARD_BG, fg=_t.FG_MUTED).pack(side="left", pady=14)
        tk.Frame(self, bg=_t.GRID_LINE, height=1).pack(fill="x")

        # ── entrada ──
        frm_in = self._secao("1. Arquivo de entrada")
        frm_in.pack(fill="x", **PAD)

        self.var_entrada = tk.StringVar(value="Nenhum arquivo selecionado")
        tk.Label(frm_in, textvariable=self.var_entrada, font=("Segoe UI", 9),
                 bg=_t.BG_APP, fg=_t.FG_TEXT, anchor="w").pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self._botao(frm_in, "Selecionar .xlsx", self._selecionar_entrada,
                    _t.PRIMARY).pack(side="right", padx=10, pady=8)

        # ── tabela cnpj ──
        frm_cnpj = self._secao("2. Tabela de CNPJ (opcional)")
        frm_cnpj.pack(fill="x", **PAD)

        self.var_cnpj = tk.StringVar(value="Nenhum arquivo selecionado  —  UF Destinatário ficará vazio")
        tk.Label(frm_cnpj, textvariable=self.var_cnpj, font=("Segoe UI", 9),
                 bg=_t.BG_APP, fg=_t.FG_TEXT, anchor="w").pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self._botao(frm_cnpj, "Selecionar .xlsx", self._selecionar_cnpj,
                    _t.SURFACE2, fg=_t.FG_TEXT).pack(side="right", padx=10, pady=8)

        # ── saída ──
        frm_out = self._secao("3. Arquivo de saída")
        frm_out.pack(fill="x", **PAD)

        self.var_saida = tk.StringVar(value="Nenhum arquivo selecionado")
        tk.Label(frm_out, textvariable=self.var_saida, font=("Segoe UI", 9),
                 bg=_t.BG_APP, fg=_t.FG_TEXT, anchor="w").pack(side="left", fill="x", expand=True, padx=10, pady=8)
        self._botao(frm_out, "Escolher destino", self._selecionar_saida,
                    _t.SURFACE2, fg=_t.FG_TEXT).pack(side="right", padx=10, pady=8)

        # ── stats ──
        frm_stats = tk.Frame(self, bg=_t.BG_APP)
        frm_stats.pack(fill="x", padx=20, pady=(0, 8))
        self._stat_boxes = {}
        for label, key in [("Linhas", "linhas"), ("Colunas originais", "colunas"), ("CSTs detectados", "csts")]:
            f = tk.Frame(frm_stats, bg=_t.CARD_BG, bd=0, highlightthickness=1,
                         highlightbackground=_t.GRID_LINE)
            f.pack(side="left", expand=True, fill="x", padx=(0, 8))
            tk.Label(f, text=label, font=("Segoe UI", 8), bg=_t.CARD_BG, fg=_t.FG_MUTED).pack(pady=(8, 0))
            lbl = tk.Label(f, text="—", font=("Segoe UI", 16, "bold"), bg=_t.CARD_BG, fg=_t.ACCENT_CYAN)
            lbl.pack(pady=(0, 8))
            self._stat_boxes[key] = lbl

        # ── log ──
        frm_log = self._secao("Log")
        frm_log.pack(fill="both", expand=True, **PAD)

        self.txt_log = tk.Text(frm_log, height=8, font=("Consolas", 8),
                               bg=_t.CARD_BG, fg=_t.FG_TEXT, relief="flat", bd=0,
                               insertbackground=_t.FG_TEXT, state="disabled", wrap="word")
        sb = ttk.Scrollbar(frm_log, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt_log.pack(fill="both", expand=True, padx=4, pady=4)

        # ── barra de progresso ──
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=400)
        self.progress.pack(padx=20, pady=(0, 6), fill="x")

        # ── botão converter ──
        self.btn_converter = tk.Button(
            self, text="▶  Converter", font=("Segoe UI", 11, "bold"),
            bg=_t.PRIMARY, fg="white", activebackground=_t.PRIMARY_HOVER,
            activeforeground="white", relief="flat", bd=0, cursor="hand2",
            padx=24, pady=10, command=self._iniciar_conversao,
            state="disabled"
        )
        self.btn_converter.pack(pady=(0, 16))

        self._path_entrada = None
        self._path_saida = None
        self._path_cnpj = _achar_tabela_cnpj()
        if self._path_cnpj:
            self.var_cnpj.set(f"{os.path.basename(self._path_cnpj)}  (padrão do sistema)")

    def _log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"  {msg}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _set_stat(self, key, val):
        self._stat_boxes[key].config(text=str(val))

    def _selecionar_entrada(self):
        p = filedialog.askopenfilename(
            title="Selecionar banco geral",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if not p:
            return
        self._path_entrada = p
        self.var_entrada.set(os.path.basename(p))
        # sugerir saída automática
        base = os.path.splitext(p)[0]
        self._path_saida = base + "_modelo.xlsx"
        self.var_saida.set(os.path.basename(self._path_saida))
        self._verificar_pronto()

    def _selecionar_cnpj(self):
        p = filedialog.askopenfilename(
            title="Selecionar tabela de CNPJ",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if not p:
            return
        self._path_cnpj = p
        self.var_cnpj.set(os.path.basename(p))
        self._verificar_pronto()

    def _selecionar_saida(self):
        p = filedialog.asksaveasfilename(
            title="Salvar modelo",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")]
        )
        if not p:
            return
        self._path_saida = p
        self.var_saida.set(os.path.basename(p))
        self._verificar_pronto()

    def _verificar_pronto(self):
        if self._path_entrada and self._path_saida:
            self.btn_converter.configure(state="normal")

    def _iniciar_conversao(self):
        self.btn_converter.configure(state="disabled")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
        self.progress.start(10)
        t = threading.Thread(target=self._converter, daemon=True)
        t.start()

    def _converter(self):
        try:
            self._log(f"Lendo: {os.path.basename(self._path_entrada)}")

            # tenta calamine (muito mais rápido); cai para openpyxl se não disponível
            try:
                import python_calamine  # noqa
                engine = "calamine"
            except ImportError:
                engine = "openpyxl"

            df = pd.read_excel(self._path_entrada, dtype=str, engine=engine)
            df.columns = [str(c).strip() for c in df.columns]

            n_linhas = len(df)
            n_colunas = len(df.columns)
            self._log(f"Linhas: {n_linhas:,}   Colunas: {n_colunas}  (engine: {engine})")
            self.after(0, self._set_stat, "linhas",   f"{n_linhas:,}")
            self.after(0, self._set_stat, "colunas",  str(n_colunas))

            # carregar tabela de CNPJ → UF
            cnpj_map = None
            if self._path_cnpj:
                self._log(f"Carregando tabela CNPJ: {os.path.basename(self._path_cnpj)}")
                df_cnpj = pd.read_excel(self._path_cnpj, dtype=str, header=None)
                # localiza a linha que contém "CNPJ" e "UF" (em qualquer posição)
                header_row = None
                for i, row in df_cnpj.iterrows():
                    vals = [str(v).strip().upper() for v in row if pd.notna(v)]
                    if "CNPJ" in vals and "UF" in vals:
                        header_row = i
                        break
                if header_row is None:
                    self._log("⚠ Tabela CNPJ não contém colunas 'CNPJ' e 'UF' — UF Destinatário ficará vazio")
                else:
                    df_cnpj.columns = [str(v).strip() if pd.notna(v) else f"_col{i}" for i, v in enumerate(df_cnpj.iloc[header_row])]
                    df_cnpj = df_cnpj.iloc[header_row + 1:].reset_index(drop=True)
                    col_cnpj = next((c for c in df_cnpj.columns if c.upper() == "CNPJ"), None)
                    col_uf   = next((c for c in df_cnpj.columns if c.upper() == "UF"),   None)
                    df_cnpj = df_cnpj.dropna(subset=[col_cnpj, col_uf])
                    # normaliza: remove pontuação, mantém só dígitos
                    cnpj_limpo = df_cnpj[col_cnpj].str.strip().str.replace(r'\D', '', regex=True)
                    cnpj_map = dict(zip(cnpj_limpo, df_cnpj[col_uf].str.strip()))
                    self._log(f"Tabela CNPJ carregada: {len(cnpj_map):,} registros")

            df_out, csts = processar(df, self._log, cnpj_map)
            self.after(0, self._set_stat, "csts", str(len(csts)))

            self._log("Gravando xlsx...")
            salvar_xlsx(df_out, self._path_saida, self._log)

            self._df_out = df_out
            self.after(0, self._concluido)
        except Exception as e:
            self.after(0, self._erro, str(e))

    def _concluido(self):
        self.progress.stop()
        self.progress["value"] = 100
        self._log("✓ Conversão concluída com sucesso!")
        self.btn_converter.configure(state="normal")
        messagebox.showinfo("Concluído",
                            f"Arquivo salvo em:\n{self._path_saida}")

    def _erro(self, msg):
        self.progress.stop()
        self.btn_converter.configure(state="normal")
        self._log(f"ERRO: {msg}")
        messagebox.showerror("Erro", msg)


# --------------------------------------------------------------------------- #
# INTEGRAÇÃO COM O MENU PRINCIPAL
# --------------------------------------------------------------------------- #

def _build_content(win):
    """Monta a tela dentro de uma janela (Toplevel) já criada pelo
    menu_principal.py, seguindo o padrão dos outros módulos do app."""
    win.title("Conversor NF-e — ICMS Unificado")
    win.resizable(False, False)
    win.configure(bg=_t.BG_APP)
    _t.apply_theme(win)

    frame = TransmitFrame(win)
    frame.pack(fill="both", expand=True)

    win.update_idletasks()
    return frame


if __name__ == "__main__":
    # Execução independente (fora do app principal), útil para testes.
    root = tk.Tk()
    root.title("Conversor NF-e — ICMS Unificado")
    root.resizable(False, False)
    root.configure(bg=_t.BG_APP)
    _t.apply_theme(root)
    TransmitFrame(root).pack(fill="both", expand=True)

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    root.mainloop()