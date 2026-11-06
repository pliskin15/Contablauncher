import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import re
from collections import defaultdict

# ── cores e fontes ──────────────────────────────────────────────
from core import theme as _t     # identidade visual do Prisma (core/theme.py)

BG       = _t.BG_APP
CARD     = _t.CARD_BG
ACCENT   = _t.ACCENT_BLUE
GREEN    = _t.ACCENT_GREEN
RED      = _t.ACCENT_RED
YELLOW   = _t.ACCENT_ORANGE
TXT      = _t.FG_TEXT
TXT2     = _t.FG_MUTED
HOVER    = _t.GRID_LINE         # hover / linha de tabela
FIELD    = _t.SURFACE2          # campos de texto
FONT     = ("Segoe UI", 10)
FONT_B   = ("Segoe UI", 10, "bold")
FONT_H   = ("Segoe UI", 13, "bold")
FONT_SM  = ("Segoe UI", 9)

# ── parsing ─────────────────────────────────────────────────────
def parse_razao(path):
    """Retorna dois dicts: total_por_filial e z_por_filial (conta partida xx-1365).

    Suporta o formato Razão Modelo I (.xls/.xlsx) com uma sheet por filial.
    Cada sheet tem:
      - Linha com col0='FILIAL ', col1=número da filial
      - Linha de cabeçalho com col0='DATA'
      - Lançamentos abaixo do cabeçalho
    """
    engine = "xlrd" if str(path).lower().endswith(".xls") else None
    kwargs = {"engine": engine} if engine else {}

    xl = pd.ExcelFile(path, **kwargs)
    # ignora sheet de resumo
    sheet_names = [s for s in xl.sheet_names
                   if "RESUMO" not in s.upper()]

    total  = defaultdict(float)
    z_vals = defaultdict(float)

    for sheet in sheet_names:
        df_raw = xl.parse(sheet, header=None, dtype=str)

        current_filial = None
        in_data        = False
        idx_cp         = None
        idx_debito     = None

        for _, row in df_raw.iterrows():
            vals  = row.tolist()
            first = str(vals[0]).strip() if pd.notna(vals[0]) else ""

            # ── linha de identificação de filial ──
            if first.upper().startswith("FILIAL"):
                try:
                    current_filial = int(float(str(vals[1]).strip()))
                except (ValueError, TypeError, IndexError):
                    pass
                in_data = False
                continue

            # ── linha de cabeçalho da tabela ──
            if first.upper() == "DATA":
                col_names = [str(v).strip().lower() for v in vals]
                try:
                    idx_cp     = col_names.index("conta partida")
                    idx_debito = col_names.index("debito")
                    in_data    = True
                except ValueError:
                    in_data = False
                continue

            # ── linhas de total / rodapé — ignorar ──
            if "TOTAL" in first.upper():
                continue

            # ── linha de lançamento ──
            if in_data and current_filial is not None:
                try:
                    cp  = str(vals[idx_cp]).strip()
                    deb = float(str(vals[idx_debito]).replace(",", "."))
                except (ValueError, TypeError, IndexError):
                    continue

                total[current_filial] += deb
                if re.search(r"-\s*1365\s*$", cp):
                    z_vals[current_filial] += deb

    if not total:
        raise ValueError(
            "Nenhum lançamento encontrado no Razão. "
            "Verifique se o arquivo está no formato Razão Modelo I."
        )

    return dict(total), dict(z_vals)


def parse_pdf(path):
    """Extrai filial → total_filial do PDF, linha a linha.
    Filiais sem número explícito: 1ª ocorrência = 29, 2ª = 61."""
    filial_totals = {}
    UNNAMED_SEQUENCE = [29, 61]   # ordem de aparição no PDF
    unnamed_idx = 0

    with pdfplumber.open(path) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    HEADER_LINE   = re.compile(
        r"^(?:Emiss|Hora|Usu|Data Contabil|Filial|Data Documento|Pagina)",
        re.IGNORECASE)
    FILIAL_NUMBERED   = re.compile(r"^(\d+)\s+[A-ZÁÉÍÓÚ]")
    FILIAL_UNNUMBERED = re.compile(r"^[A-ZÁÉÍÓÚ].*:")   # endereço sem número na frente
    TOTAL_FILIAL      = re.compile(r"TOTAL FILIAL\s+([\d\.,]+)")

    current_filial = None

    for line in full_text.splitlines():
        line = line.strip()
        if not line or HEADER_LINE.match(line):
            continue

        # filial com número
        mf = FILIAL_NUMBERED.match(line)
        if mf:
            current_filial = int(mf.group(1))
            continue

        # filial sem número (endereço direto, ex.: "AV. XPTO : BAIRRO - CIDADE")
        if FILIAL_UNNUMBERED.match(line) and current_filial is None:
            if unnamed_idx < len(UNNAMED_SEQUENCE):
                current_filial = UNNAMED_SEQUENCE[unnamed_idx]
                unnamed_idx += 1
            continue

        # total filial
        mt = TOTAL_FILIAL.search(line)
        if mt and current_filial is not None:
            val_str = mt.group(1).replace(".", "").replace(",", ".")
            try:
                filial_totals[current_filial] = float(val_str)
            except ValueError:
                pass
            current_filial = None

    return filial_totals


# ── exportação Excel ────────────────────────────────────────────
def exportar_excel(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliação Energia"

    # cores
    C_HDR_BG  = "2F3655"
    C_HDR_FG  = "FFFFFF"
    C_OK      = "C6EFCE"
    C_DIV     = "FFC7CE"
    C_SEMPDF  = "FFEB9C"
    C_SEMRAZ  = "EDEDED"
    C_TOTAL   = "D9E1F2"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["Filial", "X — Razão", "Y — PDF", "Z — Cta 1365", "X - Y - Z", "Status"]
    col_widths = [10, 18, 18, 18, 18, 16]

    # cabeçalho
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = Font(bold=True, color=C_HDR_FG, name="Arial", size=10)
        cell.fill      = PatternFill("solid", fgColor=C_HDR_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22

    TAG_COLOR = {"ok": C_OK, "diverg": C_DIV, "sem_pdf": C_SEMPDF, "sem_raz": C_SEMRAZ}
    FMT_BRL   = '#,##0.00'

    for r_idx, r in enumerate(rows, 2):
        bg = TAG_COLOR.get(r["tag"], "FFFFFF")
        fill = PatternFill("solid", fgColor=bg)

        def cell(col, val, fmt=None, align="center"):
            c = ws.cell(row=r_idx, column=col, value=val)
            c.font      = Font(name="Arial", size=10)
            c.fill      = fill
            c.border    = border
            c.alignment = Alignment(horizontal=align)
            if fmt:
                c.number_format = fmt
            return c

        cell(1, r["filial"])
        cell(2, r["x_razao"],   FMT_BRL)
        cell(3, r["y_pdf"],     FMT_BRL)
        cell(4, r["z_1365"],    FMT_BRL)
        cell(5, r["resultado"], FMT_BRL)
        cell(6, r["status"])

    # linha de totais
    n = len(rows)
    tot_row = n + 2
    fill_tot = PatternFill("solid", fgColor=C_TOTAL)

    def tot_cell(col, val, fmt=None):
        c = ws.cell(row=tot_row, column=col, value=val)
        c.font      = Font(bold=True, name="Arial", size=10)
        c.fill      = fill_tot
        c.border    = border
        c.alignment = Alignment(horizontal="center")
        if fmt: c.number_format = fmt
        return c

    tot_cell(1, "TOTAL")
    for col_letter, col_idx in [("B", 2), ("C", 3), ("D", 4), ("E", 5)]:
        tot_cell(col_idx, f"=SUM({col_letter}2:{col_letter}{n+1})", FMT_BRL)
    tot_cell(6, "")

    # freeze e auto-filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{n+1}"

    wb.save(path)


# ── aplicação ────────────────────────────────────────────────────
class EnergiaFrame(tk.Frame):
    """Frame embutível — usado dentro de um Toplevel do menu principal.
    Mantém sua própria paleta fixa de cores (mesmo padrão dos módulos
    Cofre/Brinks/Prossegur), sem depender do tema dinâmico T()."""

    def __init__(self, master):
        super().__init__(master, bg=BG)
        self.razao_path = tk.StringVar(value="")
        self.pdf_path   = tk.StringVar(value="")
        self.razao_data = {}   # filial → total X
        self.z_data     = {}   # filial → Z
        self.pdf_data   = {}   # filial → Y

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── cabeçalho ──
        hdr = tk.Frame(self, bg=ACCENT, height=54)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚡  Conciliação de Energia Elétrica",
                 bg=ACCENT, fg="white", font=("Segoe UI", 14, "bold"),
                 padx=20).pack(side="left", pady=12)

        # ── painel de arquivos ──
        files_frame = tk.Frame(self, bg=BG, pady=10)
        files_frame.pack(fill="x", padx=20)

        self._file_row(files_frame, "Razão Energia (.xlsx):", self.razao_path,
                       lambda: self._pick("xlsx"), 0)
        self._file_row(files_frame, "Relatório de NF (.pdf):", self.pdf_path,
                       lambda: self._pick("pdf"), 1)

        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(pady=(6, 14))

        tk.Button(btn_row, text="▶  Processar Conciliação",
                  bg=ACCENT, fg="white", font=FONT_B,
                  relief="flat", padx=20, pady=8,
                  activebackground=_t.PRIMARY_HOVER, cursor="hand2",
                  command=self._processar).pack(side="left", padx=6)

        self.btn_export = tk.Button(btn_row, text="📥  Exportar Excel",
                  bg=GREEN, fg="white", font=FONT_B,
                  relief="flat", padx=20, pady=8,
                  activebackground="#00a862", cursor="hand2",
                  state="disabled", command=self._exportar)
        self.btn_export.pack(side="left", padx=6)

        # ── cards de resumo ──
        self.summary_frame = tk.Frame(self, bg=BG)
        self.summary_frame.pack(fill="x", padx=20, pady=(0, 10))

        # ── tabela ──
        table_card = tk.Frame(self, bg=CARD, bd=0)
        table_card.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        # filtro / pesquisa
        filter_row = tk.Frame(table_card, bg=CARD)
        filter_row.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(filter_row, text="Filtrar filial:", bg=CARD, fg=TXT2,
                 font=FONT_SM).pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self._apply_filter())
        ent = tk.Entry(filter_row, textvariable=self.filter_var,
                       bg=FIELD, fg=TXT, insertbackground=TXT,
                       relief="flat", font=FONT, width=10)
        ent.pack(side="left", padx=6)

        # status filter
        self.status_filter = tk.StringVar(value="Todos")
        for label in ("Todos", "OK", "Divergência", "Sem PDF", "Sem Razão"):
            rb = tk.Radiobutton(filter_row, text=label,
                                variable=self.status_filter, value=label,
                                bg=CARD, fg=TXT2, selectcolor=CARD,
                                activebackground=CARD, font=FONT_SM,
                                command=self._apply_filter)
            rb.pack(side="left", padx=4)

        # treeview
        cols = ("filial", "x_razao", "y_pdf", "z_1365", "resultado", "status")
        self.tree = ttk.Treeview(table_card, columns=cols, show="headings",
                                 selectmode="browse")

        headers = {
            "filial":    ("Filial", 70),
            "x_razao":   ("X — Razão", 130),
            "y_pdf":     ("Y — PDF", 130),
            "z_1365":    ("Z — Cta 1365", 130),
            "resultado": ("X - Y - Z", 130),
            "status":    ("Status", 110),
        }
        for col, (text, w) in headers.items():
            self.tree.heading(col, text=text,
                              command=lambda c=col: self._sort(c))
            self.tree.column(col, width=w, anchor="center", minwidth=60)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=CARD, foreground=TXT,
                        fieldbackground=CARD, rowheight=26,
                        font=FONT)
        style.configure("Treeview.Heading",
                        background=BG, foreground=_t.ACCENT_CYAN,
                        font=("Segoe UI", 8, "bold"), relief="flat", padding=(6, 6))
        style.map("Treeview", background=[("selected", _t.SEL_BG)],
                  foreground=[("selected", TXT)])

        self.tree.tag_configure("ok",       foreground=GREEN)
        self.tree.tag_configure("diverg",   foreground=RED)
        self.tree.tag_configure("sem_pdf",  foreground=YELLOW)
        self.tree.tag_configure("sem_raz",  foreground=TXT2)

        vsb = ttk.Scrollbar(table_card, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # ── rodapé ──
        self.status_bar = tk.Label(self, text="Selecione os arquivos e clique em Processar.",
                                   bg=BG, fg=TXT2, font=FONT_SM,
                                   anchor="w", padx=14)
        self.status_bar.pack(fill="x", side="bottom")

        self._all_rows = []   # cache para filtro

    def _file_row(self, parent, label, var, cmd, row):
        tk.Label(parent, text=label, bg=BG, fg=TXT2, font=FONT_SM,
                 width=26, anchor="e").grid(row=row, column=0, padx=(0, 6),
                                            pady=4, sticky="e")
        ent = tk.Entry(parent, textvariable=var, width=62,
                       bg=CARD, fg=TXT, insertbackground=TXT,
                       relief="flat", font=FONT, state="readonly")
        ent.grid(row=row, column=1, padx=(0, 6))
        tk.Button(parent, text="📂 Abrir", command=cmd,
                  bg=CARD, fg=ACCENT, font=FONT_SM, relief="flat",
                  activebackground=HOVER, cursor="hand2",
                  padx=8).grid(row=row, column=2)

    # ── pickers ──
    def _pick(self, tipo):
        if tipo == "xlsx":
            p = filedialog.askopenfilename(
                title="Selecionar Razão Energia",
                filetypes=[("Excel", "*.xlsx *.xls")])
            if p:
                self.razao_path.set(p)
        else:
            p = filedialog.askopenfilename(
                title="Selecionar Relatório de NF",
                filetypes=[("PDF", "*.pdf")])
            if p:
                self.pdf_path.set(p)

    # ── processamento ──
    def _processar(self):
        rx = self.razao_path.get()
        px = self.pdf_path.get()
        if not rx or not px:
            messagebox.showwarning("Atenção", "Selecione os dois arquivos.")
            return
        self.status_bar.config(text="⏳ Processando…")
        self.update_idletasks()
        try:
            self.razao_data, self.z_data = parse_razao(rx)
            self.pdf_data = parse_pdf(px)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.status_bar.config(text="Erro no processamento.")
            return

        self._build_rows()
        self._update_summary()
        self.btn_export.config(state="normal")
        self.status_bar.config(text=f"✅ Conciliação concluída — "
                               f"{len(self._all_rows)} filiais processadas.")

    def _exportar(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            title="Salvar relatório",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="conciliacao_energia.xlsx")
        if not path:
            return
        try:
            exportar_excel(self._all_rows, path)
            self.status_bar.config(text=f"✅ Relatório salvo em {path}")
            messagebox.showinfo("Exportado", f"Arquivo salvo com sucesso:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    def _build_rows(self):
        all_filiais = sorted(
            set(self.razao_data) | set(self.pdf_data))
        rows = []
        for f in all_filiais:
            x = self.razao_data.get(f, 0.0)
            y = self.pdf_data.get(f, 0.0)
            z = self.z_data.get(f, 0.0)
            res = x - y - z

            has_razao = f in self.razao_data
            has_pdf   = f in self.pdf_data

            if not has_pdf:
                status = "Sem PDF"
                tag    = "sem_pdf"
            elif not has_razao:
                status = "Sem Razão"
                tag    = "sem_raz"
            elif abs(res) < 0.02:
                status = "OK ✓"
                tag    = "ok"
            else:
                status = "Divergência"
                tag    = "diverg"

            rows.append({
                "filial":    f,
                "x_razao":   x,
                "y_pdf":     y,
                "z_1365":    z,
                "resultado": res,
                "status":    status,
                "tag":       tag,
            })
        self._all_rows = rows
        self._apply_filter()

    def _apply_filter(self):
        text_f  = self.filter_var.get().strip().lower()
        stat_f  = self.status_filter.get()

        self.tree.delete(*self.tree.get_children())
        for r in self._all_rows:
            if text_f and text_f not in str(r["filial"]).lower():
                continue
            if stat_f != "Todos":
                if stat_f == "OK"          and r["tag"] != "ok":      continue
                if stat_f == "Divergência" and r["tag"] != "diverg":  continue
                if stat_f == "Sem PDF"     and r["tag"] != "sem_pdf": continue
                if stat_f == "Sem Razão"   and r["tag"] != "sem_raz": continue

            self.tree.insert("", "end", tags=(r["tag"],), values=(
                r["filial"],
                f"R$ {r['x_razao']:>12,.2f}",
                f"R$ {r['y_pdf']:>12,.2f}",
                f"R$ {r['z_1365']:>12,.2f}",
                f"R$ {r['resultado']:>12,.2f}",
                r["status"],
            ))

    def _update_summary(self):
        for w in self.summary_frame.winfo_children():
            w.destroy()

        total_x   = sum(self.razao_data.values())
        total_y   = sum(self.pdf_data.values())
        total_z   = sum(self.z_data.values())
        total_res = total_x - total_y - total_z

        n_ok     = sum(1 for r in self._all_rows if r["tag"] == "ok")
        n_div    = sum(1 for r in self._all_rows if r["tag"] == "diverg")
        n_sem_p  = sum(1 for r in self._all_rows if r["tag"] == "sem_pdf")
        n_total  = len(self._all_rows)

        cards = [
            ("Total X (Razão)",    f"R$ {total_x:,.2f}",   ACCENT),
            ("Total Y (PDF)",      f"R$ {total_y:,.2f}",   ACCENT),
            ("Total Z (Cta 1365)", f"R$ {total_z:,.2f}",   YELLOW),
            ("Resultado X-Y-Z",    f"R$ {total_res:,.2f}",
             GREEN if abs(total_res) < 0.05 else RED),
            ("Filiais OK",         f"{n_ok} / {n_total}",  GREEN),
            ("Divergências",       str(n_div),              RED if n_div else GREEN),
            ("Sem PDF",            str(n_sem_p),            YELLOW if n_sem_p else GREEN),
        ]
        for i, (lbl, val, color) in enumerate(cards):
            card = tk.Frame(self.summary_frame, bg=CARD,
                            padx=14, pady=8)
            card.grid(row=0, column=i, padx=5, sticky="nsew")
            self.summary_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=lbl, bg=CARD, fg=TXT2, font=FONT_SM).pack()
            tk.Label(card, text=val, bg=CARD, fg=color, font=FONT_B).pack()

    def _sort(self, col):
        data = [(self.tree.set(child, col), child)
                for child in self.tree.get_children("")]
        # tenta ordenar numericamente (remove R$ e vírgulas)
        def key(x):
            s = x[0].replace("R$", "").replace(",", "").replace(" ", "")
            try:    return float(s)
            except: return x[0]
        data.sort(key=key)
        for idx, (_, child) in enumerate(data):
            self.tree.move(child, "", idx)


# --------------------------------------------------------------------------- #
# INTEGRAÇÃO COM O MENU PRINCIPAL
# --------------------------------------------------------------------------- #

def _build_content(win):
    """Monta a tela dentro de uma janela (Toplevel) já criada pelo
    menu_principal.py, seguindo o padrão dos outros módulos do app."""
    win.title("Conciliação de Energia Elétrica — PMZ")
    win.geometry("1200x780")
    win.minsize(960, 620)
    win.configure(bg=BG)

    frame = EnergiaFrame(win)
    frame.pack(fill="both", expand=True)
    return frame


if __name__ == "__main__":
    # Execução independente (fora do app principal), útil para testes.
    root = tk.Tk()
    root.title("Conciliação de Energia Elétrica — PMZ")
    root.geometry("1200x780")
    root.configure(bg=BG)
    EnergiaFrame(root).pack(fill="both", expand=True)
    root.mainloop()