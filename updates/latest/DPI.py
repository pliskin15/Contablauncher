"""
Consulta de Documentos de Importação - SEFAZ-AM (GDDI)

Combina dois endpoints:
  1) GET  /dfe-api/gddi/v1/documento-importacao/relatorio/detalhado/{inscricaoEstadual}/{periodoReferencia}
     -> Lista as DAIs de um contribuinte em um período (ex: 202601 = jan/2026)
  2) POST /dfe-api/gddi/v1/documento-importacao/selagem?nrDaiFormatado=...
     -> Consulta dados de selagem de uma DAI específica

Requisitos:
    pip install requests-pkcs12

Uso:
    python DPI.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import threading

from requests_pkcs12 import get, post

try:
    from core import theme
except ImportError:
    # Permite rodar o arquivo isoladamente (fora da estrutura do app)
    import theme

BASE = "https://sistemas.sefaz.am.gov.br/dfe-api/gddi/v1/documento-importacao"
URL_RELATORIO = BASE + "/relatorio/detalhado/{inscricaoEstadual}/{periodoReferencia}"
URL_SELAGEM = BASE + "/selagem"

COLUNAS_RELATORIO = [
    "txNrDaiFormatado", "dtDocumento", "cdTributo", "codTributacao",
    "vlDai", "vlFob", "vlIcmsCalculado", "vlIcmsDeclarado",
    "pctAliquota", "pctReducao", "vlTaxaDolar", "csComplementar",
    "csRetificacao", "dtNotificacao", "txCanalVistoria", "cdSubGrupoTributo",
]


class ConsultaDAIApp:
    def __init__(self, root):
        self.root = root
        root.title("Consulta DAI - SEFAZ-AM")
        root.geometry("1100x720")
        root.minsize(900, 560)

        theme.apply_theme(root)
        self._build_ui()
        self.dados_relatorio_completo = []

    # ---------------------------------------------------------------
    # ESTRUTURA GERAL
    # ---------------------------------------------------------------
    def _build_ui(self):
        # --- Toolbar (igual ao padrão do restante do app) ---
        toolbar = ttk.Frame(self.root, style="Toolbar.TFrame", height=52)
        toolbar.pack(fill="x", side="top")
        toolbar.pack_propagate(False)

        ttk.Label(toolbar, text="🛃  Consulta DAI",
                  style="Toolbar.TLabel",
                  font=("Segoe UI Semibold", 14)).pack(side="left", padx=20, pady=10)
        ttk.Label(toolbar, text="SEFAZ-AM · GDDI", style="Toolbar.TLabel",
                  font=("Segoe UI", 10)).pack(side="right", padx=20)

        ttk.Separator(self.root, orient="horizontal").pack(fill="x")

        body = ttk.Frame(self.root, style="App.TFrame")
        body.pack(fill="both", expand=True, padx=16, pady=14)

        self._montar_cartao_certificado(body)

        notebook = ttk.Notebook(body)
        notebook.pack(fill="both", expand=True, pady=(14, 0))

        self.tab_relatorio = ttk.Frame(notebook, style="App.TFrame")
        self.tab_selagem = ttk.Frame(notebook, style="App.TFrame")
        notebook.add(self.tab_relatorio, text="1. Relatório Detalhado (listar DAIs)")
        notebook.add(self.tab_selagem, text="2. Consulta de Selagem")

        self._montar_tab_relatorio()
        self._montar_tab_selagem()

    def _montar_cartao_certificado(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame")
        card.pack(fill="x")

        inner = ttk.Frame(card, style="Card.TFrame")
        inner.pack(fill="x", padx=16, pady=14)

        ttk.Label(inner, text="Certificado Digital (A1)",
                  style="Section.TLabel", background=theme.CARD_BG).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        self.cert_path_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.cert_path_var, width=58).grid(
            row=1, column=0, sticky="we", padx=(0, 8)
        )
        ttk.Button(inner, text="Selecionar...", style="TButton",
                   command=self.selecionar_certificado).grid(row=1, column=1, padx=(0, 16))

        ttk.Label(inner, text="Senha:", background=theme.CARD_BG,
                  foreground=theme.FG_TEXT).grid(row=1, column=2, sticky="e", padx=(0, 8))
        self.cert_senha_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.cert_senha_var, show="*", width=20).grid(
            row=1, column=3, sticky="w"
        )

        inner.columnconfigure(0, weight=1)

    # ---------------------------------------------------------------
    # ABA 1: Relatório Detalhado
    # ---------------------------------------------------------------
    def _montar_tab_relatorio(self):
        wrapper = ttk.Frame(self.tab_relatorio, style="App.TFrame")
        wrapper.pack(fill="both", expand=True, pady=(12, 0))

        # --- Parâmetros de consulta ---
        card_params = ttk.Frame(wrapper, style="Card.TFrame")
        card_params.pack(fill="x")

        frame_params = ttk.Frame(card_params, style="Card.TFrame")
        frame_params.pack(fill="x", padx=16, pady=12)

        ttk.Label(frame_params, text="Inscrição Estadual:", background=theme.CARD_BG,
                  foreground=theme.FG_TEXT).grid(row=0, column=0, sticky="w")
        self.ie_var = tk.StringVar()
        ttk.Entry(frame_params, textvariable=self.ie_var, width=18).grid(
            row=0, column=1, padx=(8, 24))

        ttk.Label(frame_params, text="Período (AAAAMM, ex: 202601):", background=theme.CARD_BG,
                  foreground=theme.FG_TEXT).grid(row=0, column=2, sticky="w")
        self.periodo_var = tk.StringVar()
        ttk.Entry(frame_params, textvariable=self.periodo_var, width=14).grid(
            row=0, column=3, padx=(8, 16))

        self.btn_relatorio = ttk.Button(frame_params, text="Consultar Relatório",
                                        style="Primary.TButton", command=self.consultar_relatorio)
        self.btn_relatorio.grid(row=0, column=4, padx=(8, 0))

        # --- Filtro local ---
        card_filtro = ttk.Frame(wrapper, style="Card.TFrame")
        card_filtro.pack(fill="x", pady=(10, 0))

        frame_filtro = ttk.Frame(card_filtro, style="Card.TFrame")
        frame_filtro.pack(fill="x", padx=16, pady=10)

        ttk.Label(frame_filtro, text="Buscar DAI específica na lista carregada:",
                  background=theme.CARD_BG, foreground=theme.FG_TEXT).pack(side="left")
        self.filtro_dai_var = tk.StringVar()
        self.filtro_dai_var.trace_add("write", lambda *args: self._filtrar_tabela())
        ttk.Entry(frame_filtro, textvariable=self.filtro_dai_var, width=25).pack(
            side="left", padx=12)
        ttk.Button(frame_filtro, text="Ver detalhes completos", style="Ghost.TButton",
                   command=self._ver_detalhes_item).pack(side="left", padx=4)

        self.status_relatorio_var = tk.StringVar(value="Aguardando consulta...")
        ttk.Label(wrapper, textvariable=self.status_relatorio_var,
                  style="Muted.TLabel").pack(anchor="w", padx=4, pady=(10, 4))

        # --- Tabela de resultados ---
        frame_tree = ttk.Frame(wrapper, style="Card.TFrame")
        frame_tree.pack(fill="both", expand=True)

        tree_inner = ttk.Frame(frame_tree, style="Card.TFrame")
        tree_inner.pack(fill="both", expand=True, padx=10, pady=10)

        scroll_y = ttk.Scrollbar(tree_inner, orient="vertical")
        scroll_x = ttk.Scrollbar(tree_inner, orient="horizontal")

        self.tree = ttk.Treeview(
            tree_inner, columns=COLUNAS_RELATORIO, show="headings",
            yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set, height=15,
        )
        for col in COLUNAS_RELATORIO:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor="w")

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree_inner.rowconfigure(0, weight=1)
        tree_inner.columnconfigure(0, weight=1)

        ttk.Label(
            wrapper,
            text="Dê duplo-clique em uma linha para enviar o número da DAI direto para a aba de Consulta de Selagem.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=4, pady=(8, 0))

        self.tree.bind("<Double-1>", self._enviar_dai_para_selagem)

    def consultar_relatorio(self):
        cert_path = self.cert_path_var.get().strip()
        cert_senha = self.cert_senha_var.get()
        ie = self.ie_var.get().strip()
        periodo = self.periodo_var.get().strip()

        if not cert_path:
            messagebox.showwarning("Atenção", "Selecione o certificado (.pfx).")
            return
        if not ie or not periodo:
            messagebox.showwarning("Atenção", "Informe Inscrição Estadual e Período.")
            return

        self.btn_relatorio.config(state="disabled")
        self.status_relatorio_var.set("Consultando...")
        for item in self.tree.get_children():
            self.tree.delete(item)

        threading.Thread(
            target=self._executar_relatorio, args=(cert_path, cert_senha, ie, periodo), daemon=True
        ).start()

    def _executar_relatorio(self, cert_path, cert_senha, ie, periodo):
        try:
            url = URL_RELATORIO.format(inscricaoEstadual=ie, periodoReferencia=periodo)
            response = get(
                url,
                pkcs12_filename=cert_path,
                pkcs12_password=cert_senha,
                verify=True,
                timeout=30,
            )
            self.root.after(0, self._preencher_relatorio, response)
        except Exception as e:
            self.root.after(0, self._erro_relatorio, str(e))

    def _preencher_relatorio(self, response):
        self.status_relatorio_var.set(f"Status HTTP: {response.status_code}")
        self.btn_relatorio.config(state="normal")
        try:
            dados = response.json()
        except ValueError:
            messagebox.showerror("Erro", f"Resposta não é JSON:\n{response.text[:500]}")
            return

        if isinstance(dados, dict):
            # caso a API embrulhe a lista em uma chave (ex: {"itens": [...]})
            for chave in ("itens", "content", "resultado", "lista"):
                if chave in dados and isinstance(dados[chave], list):
                    dados = dados[chave]
                    break
            else:
                dados = [dados]

        if not dados:
            messagebox.showinfo("Resultado", "Nenhum documento encontrado para esse período.")
            return

        self.dados_relatorio_completo = dados  # guarda o JSON cru de cada item, indexado por posição
        for idx, item in enumerate(dados):
            valores = [item.get(col, "") for col in COLUNAS_RELATORIO]
            self.tree.insert("", "end", iid=str(idx), values=valores)

    def _erro_relatorio(self, mensagem):
        self.status_relatorio_var.set("Erro na requisição")
        self.btn_relatorio.config(state="normal")
        messagebox.showerror("Erro", mensagem)

    def _filtrar_tabela(self):
        """Filtra as linhas visíveis da tabela pelo número da DAI digitado, sem refazer a consulta."""
        termo = self.filtro_dai_var.get().strip().lower()

        for item in self.tree.get_children():
            self.tree.detach(item)

        if not termo:
            for idx in range(len(self.dados_relatorio_completo)):
                self.tree.reattach(str(idx), "", "end")
            return

        for idx, item in enumerate(self.dados_relatorio_completo):
            nr_dai = str(item.get("txNrDaiFormatado", "")).lower()
            if termo in nr_dai:
                self.tree.reattach(str(idx), "", "end")

    def _ver_detalhes_item(self):
        """Abre uma janela com TODOS os campos do item selecionado (ou do único resultado do filtro)."""
        selecionado = self.tree.focus()
        visiveis = self.tree.get_children()

        if not selecionado and len(visiveis) == 1:
            selecionado = visiveis[0]

        if not selecionado:
            messagebox.showwarning("Atenção", "Selecione uma linha na tabela (ou filtre até sobrar uma só).")
            return

        idx = int(selecionado)
        item = self.dados_relatorio_completo[idx]

        janela = tk.Toplevel(self.root)
        janela.title(f"Detalhes - DAI {item.get('txNrDaiFormatado', '')}")
        janela.geometry("520x540")
        janela.configure(bg=theme.BG_APP)
        theme.apply_theme(janela)

        ttk.Frame(janela, style="Toolbar.TFrame", height=4).pack(fill="x")

        texto = tk.Text(janela, wrap="word", bg=theme.CARD_BG, fg=theme.FG_TEXT,
                        insertbackground=theme.FG_TEXT, relief="flat",
                        font=("Consolas", 10), padx=12, pady=12,
                        highlightthickness=1, highlightbackground=theme.GRID_LINE,
                        highlightcolor=theme.PRIMARY)
        texto.pack(fill="both", expand=True, padx=12, pady=12)
        texto.insert("1.0", json.dumps(item, indent=2, ensure_ascii=False))

        ttk.Button(janela, text="Usar essa DAI na Consulta de Selagem",
                   style="Primary.TButton",
                   command=lambda: self.dai_var.set(item.get("txNrDaiFormatado", ""))).pack(pady=(0, 14))

    def _enviar_dai_para_selagem(self, event):
        selecionado = self.tree.focus()
        if not selecionado:
            return
        valores = self.tree.item(selecionado, "values")
        if valores:
            nr_dai = valores[0]  # primeira coluna = txNrDaiFormatado
            self.dai_var.set(nr_dai)
            messagebox.showinfo("OK", f"DAI '{nr_dai}' enviada para a aba de Consulta de Selagem.")

    # ---------------------------------------------------------------
    # ABA 2: Consulta de Selagem
    # ---------------------------------------------------------------
    def _montar_tab_selagem(self):
        wrapper = ttk.Frame(self.tab_selagem, style="App.TFrame")
        wrapper.pack(fill="both", expand=True, pady=(12, 0))

        card_query = ttk.Frame(wrapper, style="Card.TFrame")
        card_query.pack(fill="x")

        frame_query = ttk.Frame(card_query, style="Card.TFrame")
        frame_query.pack(fill="x", padx=16, pady=12)

        ttk.Label(frame_query, text="Número da DAI (ex: 00/0000000-0):", background=theme.CARD_BG,
                  foreground=theme.FG_TEXT).grid(row=0, column=0, sticky="w")
        self.dai_var = tk.StringVar()
        ttk.Entry(frame_query, textvariable=self.dai_var, width=30).grid(row=0, column=1, padx=12)

        self.btn_selagem = ttk.Button(frame_query, text="Consultar Selagem",
                                      style="Primary.TButton", command=self.consultar_selagem)
        self.btn_selagem.grid(row=0, column=2, padx=(8, 0))

        self.status_selagem_var = tk.StringVar(value="Aguardando consulta...")
        ttk.Label(wrapper, textvariable=self.status_selagem_var,
                  style="Muted.TLabel").pack(anchor="w", padx=4, pady=(10, 4))

        card_result = ttk.Frame(wrapper, style="Card.TFrame")
        card_result.pack(fill="both", expand=True)

        self.result_text = tk.Text(card_result, wrap="word", bg=theme.CARD_BG, fg=theme.FG_TEXT,
                                   insertbackground=theme.FG_TEXT, relief="flat",
                                   font=("Consolas", 10), padx=12, pady=12,
                                   highlightthickness=0)
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)

    def consultar_selagem(self):
        cert_path = self.cert_path_var.get().strip()
        cert_senha = self.cert_senha_var.get()
        nr_dai = self.dai_var.get().strip()

        if not cert_path:
            messagebox.showwarning("Atenção", "Selecione o certificado (.pfx).")
            return
        if not nr_dai:
            messagebox.showwarning("Atenção", "Informe o número da DAI.")
            return

        self.btn_selagem.config(state="disabled")
        self.status_selagem_var.set("Consultando...")
        self.result_text.delete("1.0", tk.END)

        threading.Thread(
            target=self._executar_selagem, args=(cert_path, cert_senha, nr_dai), daemon=True
        ).start()

    def _executar_selagem(self, cert_path, cert_senha, nr_dai):
        try:
            response = post(
                URL_SELAGEM,
                params={"nrDaiFormatado": nr_dai},
                pkcs12_filename=cert_path,
                pkcs12_password=cert_senha,
                verify=True,
                timeout=30,
            )
            self.root.after(0, self._mostrar_resultado_selagem, response)
        except Exception as e:
            self.root.after(0, self._erro_selagem, str(e))

    def _mostrar_resultado_selagem(self, response):
        self.status_selagem_var.set(f"Status HTTP: {response.status_code}")
        self.btn_selagem.config(state="normal")
        try:
            corpo = json.dumps(response.json(), indent=2, ensure_ascii=False)
        except ValueError:
            corpo = response.text
        self.result_text.insert("1.0", corpo)

    def _erro_selagem(self, mensagem):
        self.status_selagem_var.set("Erro na requisição")
        self.btn_selagem.config(state="normal")
        self.result_text.insert("1.0", f"Erro: {mensagem}")
        messagebox.showerror("Erro", mensagem)

    # ---------------------------------------------------------------
    def selecionar_certificado(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o certificado .pfx",
            filetypes=[("Certificado PKCS12", "*.pfx *.p12"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.cert_path_var.set(caminho)


if __name__ == "__main__":
    root = tk.Tk()
    app = ConsultaDAIApp(root)
    root.mainloop()