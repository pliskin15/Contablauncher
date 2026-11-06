"""
comparar_tabelas_gui.py
------------------------
Interface gráfica para comparar a tabela base (CSV) com a tabela
de comparação (Excel), retornando notas fiscais presentes apenas
na segunda tabela.

Requisitos:
    pip install pandas openpyxl

Rodar:
    python comparar_tabelas_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
from pathlib import Path

from core import theme as _t   # identidade visual do Prisma (core/theme.py)


# =============================================================================
# LÓGICA DE COMPARAÇÃO
# =============================================================================

def comparar(path_base, path_comp, callback_log, callback_done):
    try:
        callback_log("Lendo tabela base...")
        ext_base = Path(path_base).suffix.lower()
        if ext_base == ".csv":
            df_base = pd.read_csv(path_base, dtype=str, sep=None, engine="python")
        else:
            df_base = pd.read_excel(path_base, dtype=str)

        callback_log(f"  → {len(df_base):,} registros carregados.")

        callback_log("Lendo tabela de comparação...")
        df_comp = pd.read_excel(path_comp, dtype=str)
        callback_log(f"  → {len(df_comp):,} registros carregados.")

        # Validações
        for col in ["NOTA", "SERIE", "EMISSAO"]:
            if col not in df_base.columns:
                raise ValueError(f"Coluna '{col}' não encontrada na tabela base.\nColunas: {list(df_base.columns)}")
        for col in ["Num NF", "Série NF", "Data Emissão"]:
            if col not in df_comp.columns:
                raise ValueError(f"Coluna '{col}' não encontrada na tabela comparar.\nColunas: {list(df_comp.columns)}")

        callback_log("Normalizando chaves...")

        # Base
        df_base["_NOTA"]    = df_base["NOTA"].str.strip().str.lstrip("0")
        df_base["_SERIE"]   = df_base["SERIE"].str.strip().str.lstrip("0")
        df_base["_EMISSAO"] = pd.to_datetime(
            df_base["EMISSAO"].str.strip(), dayfirst=True, errors="coerce"
        ).dt.strftime("%Y-%m-%d")

        # Comparar
        df_comp["_NOTA"]    = df_comp["Num NF"].str.strip().str.lstrip("0")
        df_comp["_SERIE"]   = df_comp["Série NF"].str.strip().str.lstrip("0")
        df_comp["_EMISSAO"] = pd.to_datetime(
            df_comp["Data Emissão"].str.strip(), errors="coerce"
        ).dt.strftime("%Y-%m-%d")

        df_base["__chave__"] = df_base["_NOTA"] + "|" + df_base["_SERIE"] + "|" + df_base["_EMISSAO"]
        df_comp["__chave__"] = df_comp["_NOTA"] + "|" + df_comp["_SERIE"] + "|" + df_comp["_EMISSAO"]

        callback_log("Comparando...")

        chaves_base = set(df_base["__chave__"])
        mask        = ~df_comp["__chave__"].isin(chaves_base)

        COLUNAS = ["Chave NF-e", "Série NF", "Num NF", "Data Emissão", "Vlr Total NF", "Nome Emitente"]
        df_resultado = df_comp.loc[mask, COLUNAS].copy()

        total_base  = len(df_base)
        total_comp  = len(df_comp)
        total_match = total_comp - len(df_resultado)
        total_novos = len(df_resultado)

        callback_log("")
        callback_log("=" * 42)
        callback_log(f"  Registros na base:          {total_base:>8,}")
        callback_log(f"  Registros na comparação:    {total_comp:>8,}")
        callback_log(f"  ✔ Match na base:            {total_match:>8,}")
        callback_log(f"  ✗ Apenas na 2ª tabela:      {total_novos:>8,}")
        callback_log("=" * 42)

        callback_done(df_resultado, total_novos)

    except Exception as e:
        callback_log(f"\n❌ ERRO: {e}")
        callback_done(None, -1)


# =============================================================================
# INTERFACE
# =============================================================================

class NotasPendentesFrame(tk.Frame):
    """Frame reutilizável — pode ser embutido em qualquer Toplevel."""
    def __init__(self, master):
        super().__init__(master, bg=_t.BG_APP)
        self._df_resultado = None
        self._build()

    def _entry(self, row: int) -> tk.Entry:
        e = tk.Entry(self, width=52, font=("Segoe UI", 9), bg=_t.SURFACE2, fg=_t.FG_TEXT,
                     insertbackground=_t.FG_TEXT, relief="flat", bd=0,
                     highlightthickness=1, highlightbackground=_t.GRID_LINE,
                     highlightcolor=_t.ACCENT)
        e.grid(row=row, column=1, padx=12, pady=6, ipady=5)
        return e

    def _build(self):
        PAD = dict(padx=12, pady=6)

        # ── Título ──
        tk.Label(
            self, text="Comparador de Tabelas NF",
            font=("Segoe UI", 14, "bold"), bg=_t.BG_APP, fg=_t.FG_TEXT
        ).grid(row=0, column=0, columnspan=3, pady=(16, 4))

        tk.Label(
            self, text="Chave: NOTA · SERIE · EMISSAO",
            font=("Segoe UI", 9), bg=_t.BG_APP, fg=_t.FG_MUTED
        ).grid(row=1, column=0, columnspan=3, pady=(0, 12))

        # ── Tabela Base ──
        tk.Label(self, text="Tabela Base (CSV/Excel):",
                 font=("Segoe UI", 10, "bold"), bg=_t.BG_APP, fg=_t.FG_TEXT
                 ).grid(row=2, column=0, sticky="w", **PAD)

        self._entry_base = self._entry(2)

        tk.Button(self, text="Procurar", command=self._browse_base,
                  font=("Segoe UI", 9), relief="flat", bd=0, bg=_t.SURFACE2, fg=_t.FG_TEXT,
                  activebackground=_t.GRID_LINE, activeforeground="white",
                  cursor="hand2", padx=12, pady=4
                  ).grid(row=2, column=2, padx=(0, 12))

        # ── Tabela Comparar ──
        tk.Label(self, text="Tabela Comparar (Excel):",
                 font=("Segoe UI", 10, "bold"), bg=_t.BG_APP, fg=_t.FG_TEXT
                 ).grid(row=3, column=0, sticky="w", **PAD)

        self._entry_comp = self._entry(3)

        tk.Button(self, text="Procurar", command=self._browse_comp,
                  font=("Segoe UI", 9), relief="flat", bd=0, bg=_t.SURFACE2, fg=_t.FG_TEXT,
                  activebackground=_t.GRID_LINE, activeforeground="white",
                  cursor="hand2", padx=12, pady=4
                  ).grid(row=3, column=2, padx=(0, 12))

        # ── Botão comparar ──
        self._btn_comparar = tk.Button(
            self, text="▶  Comparar", command=self._iniciar,
            font=("Segoe UI", 10, "bold"), relief="flat", bd=0,
            bg=_t.PRIMARY, fg="white", activebackground=_t.PRIMARY_HOVER,
            activeforeground="white", cursor="hand2", padx=20, pady=8
        )
        self._btn_comparar.grid(row=4, column=0, columnspan=3, pady=(8, 4))

        # ── Barra de progresso ──
        self._progress = ttk.Progressbar(self, mode="indeterminate", length=460)
        self._progress.grid(row=5, column=0, columnspan=3, padx=12, pady=(0, 4))

        # ── Log ──
        tk.Label(self, text="LOG", font=("Segoe UI", 8, "bold"),
                 bg=_t.BG_APP, fg=_t.ACCENT_CYAN).grid(row=6, column=0, sticky="w", padx=12)

        frame_log = tk.Frame(self, bg=_t.BG_APP)
        frame_log.grid(row=7, column=0, columnspan=3, padx=12, pady=(0, 8))

        self._txt_log = tk.Text(
            frame_log, width=68, height=12,
            font=("Consolas", 9), state="disabled",
            bg=_t.CARD_BG, fg=_t.FG_TEXT, relief="flat", bd=0,
            insertbackground=_t.FG_TEXT
        )
        scroll = ttk.Scrollbar(frame_log, command=self._txt_log.yview)
        self._txt_log.configure(yscrollcommand=scroll.set)
        self._txt_log.pack(side="left")
        scroll.pack(side="right", fill="y")

        # ── Botão exportar ──
        self._btn_exportar = tk.Button(
            self, text="💾  Exportar resultado (.xlsx)", command=self._exportar,
            font=("Segoe UI", 10), relief="flat", bd=0,
            bg=_t.ACCENT_GREEN, fg="#001a0d", activebackground="#00a862",
            cursor="hand2", padx=18, pady=8, state="disabled"
        )
        self._btn_exportar.grid(row=8, column=0, columnspan=3, pady=(0, 16))

    # ── helpers ──────────────────────────────────────────────────────────────

    def _browse_base(self):
        path = filedialog.askopenfilename(
            title="Selecione a Tabela Base",
            filetypes=[("CSV / Excel", "*.csv *.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            self._entry_base.delete(0, tk.END)
            self._entry_base.insert(0, path)

    def _browse_comp(self):
        path = filedialog.askopenfilename(
            title="Selecione a Tabela de Comparação",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            self._entry_comp.delete(0, tk.END)
            self._entry_comp.insert(0, path)

    def _log(self, msg):
        self._txt_log.configure(state="normal")
        self._txt_log.insert(tk.END, msg + "\n")
        self._txt_log.see(tk.END)
        self._txt_log.configure(state="disabled")

    def _clear_log(self):
        self._txt_log.configure(state="normal")
        self._txt_log.delete("1.0", tk.END)
        self._txt_log.configure(state="disabled")

    def _iniciar(self):
        path_base = self._entry_base.get().strip()
        path_comp = self._entry_comp.get().strip()

        if not path_base or not path_comp:
            messagebox.showwarning("Atenção", "Selecione os dois arquivos antes de comparar.")
            return

        self._clear_log()
        self._df_resultado = None
        self._btn_comparar.configure(state="disabled")
        self._btn_exportar.configure(state="disabled")
        self._progress.start(10)

        def run():
            comparar(path_base, path_comp,
                     callback_log=lambda m: self.after(0, self._log, m),
                     callback_done=lambda df, n: self.after(0, self._finalizar, df, n))

        threading.Thread(target=run, daemon=True).start()

    def _finalizar(self, df, total_novos):
        self._progress.stop()
        self._btn_comparar.configure(state="normal")

        if df is not None and total_novos > 0:
            self._df_resultado = df
            self._btn_exportar.configure(state="normal")
            self._log(f"\n✅ Pronto! {total_novos:,} registros para exportar.")
        elif df is not None and total_novos == 0:
            self._log("\n✅ Nenhuma divergência encontrada.")
        # erro já foi logado em comparar()

    def _exportar(self):
        if self._df_resultado is None:
            return

        path = filedialog.asksaveasfilename(
            title="Salvar resultado",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="apenas_na_segunda_tabela.xlsx"
        )
        if not path:
            return

        try:
            self._df_resultado.to_excel(path, index=False)
            self._log(f"\n💾 Arquivo salvo em:\n   {path}")
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso!\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))


# =============================================================================
# ENTRY POINTS
# =============================================================================

def _build_content(win: tk.Toplevel) -> None:
    """Chamado pelo menu_principal — injeta o conteúdo no Toplevel recebido."""
    win.title("Notas Pendentes")
    win.geometry("680x560")
    win.resizable(False, False)
    frame = NotasPendentesFrame(win)
    frame.pack(fill="both", expand=True)


# Permite rodar standalone para testes
if __name__ == "__main__":
    root = tk.Tk()
    _build_content(root)
    root.mainloop()