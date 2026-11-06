"""
Painel "Consulta CT-e": cola a chave de acesso e mostra conhecimento, data, valor, frete,
fornecedor, remetente, destinatário, tomador, CFOP e notas vinculadas.

Os CT-e vêm da fila de documentos da SEFAZ (Distribuição DF-e) de cada certificado
cadastrado e ficam num cache local (%APPDATA%\\ConsultaCTe). Outros certificados podem
ser adicionados a qualquer momento pela própria tela.
"""

import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from lxml import etree

from core import theme as _t
from cte import dados


def _build_content(win) -> None:
    win.title("Consulta CT-e")
    win.geometry("1020x820")
    win.minsize(880, 660)
    _t.apply_theme(win)
    CtePanel(win).pack(fill="both", expand=True)


class CtePanel(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=_t.BG_APP)
        self.cache = dados.Cache()
        self.senhas: dict[str, str] = {}      # senhas digitadas nesta sessão
        self.fila: queue.Queue = queue.Queue()
        self.sincronizando = False
        self.chave_var = tk.StringVar()
        self.nfe_var = tk.StringVar()
        self._build()
        self._atualizar_certificados()
        self._processar_fila()

    # ── construção ───────────────────────────────────────────────────────────

    def _card(self, titulo, **pack):
        outer = tk.Frame(self, bg=_t.GRID_LINE)
        tk.Label(outer, text=titulo.upper(), font=("Segoe UI", 8, "bold"), bg=_t.GRID_LINE,
                 fg=_t.ACCENT_CYAN, padx=12, pady=5).pack(anchor="w")
        inner = tk.Frame(outer, bg=_t.CARD_BG, padx=12, pady=10)
        inner.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        outer.pack(**pack)
        return inner

    def _build(self):
        topo = tk.Frame(self, bg=_t.BG_APP)
        topo.pack(fill="x", padx=20, pady=(16, 8))
        tk.Label(topo, text="Consulta CT-e", font=_t.FONT_TITLE, bg=_t.BG_APP,
                 fg=_t.FG_TEXT).pack(side="left")
        self.lbl_cache = tk.Label(topo, text="", font=("Segoe UI", 9), bg=_t.BG_APP, fg=_t.FG_MUTED)
        self.lbl_cache.pack(side="right")

        # certificados
        f = self._card("Certificados", fill="x", padx=20, pady=(0, 10))
        cols = ("empresa", "cnpj", "estab", "validade", "situacao", "sync")
        self.tv_cert = ttk.Treeview(f, columns=cols, show="headings", height=4, selectmode="browse")
        for col, txt, larg in [("empresa", "Empresa", 240), ("cnpj", "CNPJ", 140),
                               ("estab", "Estab.", 50), ("validade", "Validade", 90), ("situacao", "Situação", 280),
                               ("sync", "Última atualização", 130)]:
            self.tv_cert.heading(col, text=txt)
            self.tv_cert.column(col, width=larg, anchor="w")
        self.tv_cert.pack(fill="x")
        linha = tk.Frame(f, bg=_t.CARD_BG)
        linha.pack(fill="x", pady=(8, 0))
        self.linha_certs = linha
        ttk.Button(linha, text="+ Adicionar certificado", command=self._adicionar_cert).pack(side="left")
        ttk.Button(linha, text="Remover", command=self._remover_cert).pack(side="left", padx=(6, 0))
        ttk.Button(linha, text="Estabelecimentos…", command=self._editar_estabelecimentos).pack(side="left", padx=(6, 0))
        self.btn_atualizar = ttk.Button(linha, text="Atualizar CT-e da SEFAZ",
                                        style="Primary.TButton", command=self._atualizar_todos)
        self.btn_atualizar.pack(side="right")
        self.lbl_status = tk.Label(linha, text="", font=("Segoe UI", 9), bg=_t.CARD_BG, fg=_t.FG_MUTED)
        self.lbl_status.pack(side="right", padx=10)

        self.btn_recolher = tk.Button(f.master, text="▾ Recolher", command=self._alternar_certs,
                                      relief="flat", cursor="hand2", font=("Segoe UI", 8), bd=0,
                                      bg=_t.SURFACE2, fg=_t.FG_TEXT, activebackground=_t.GRID_LINE,
                                      activeforeground=_t.FG_TEXT, padx=8, pady=1)
        self.btn_recolher.place(relx=1.0, x=-8, y=3, anchor="ne")
        self._alternar_certs(restaurar=True)

        # busca
        f = self._card("Chave de acesso do CT-e", fill="x", padx=20, pady=(0, 10))
        linha = tk.Frame(f, bg=_t.CARD_BG)
        linha.pack(fill="x")
        ent = tk.Entry(linha, textvariable=self.chave_var, font=("Consolas", 11))
        ent.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        ent.bind("<Return>", lambda _e: self._buscar())
        ent.focus_set()
        ttk.Button(linha, text="Buscar", style="Primary.TButton", command=self._buscar).pack(side="left")
        ttk.Button(linha, text="Abrir XML…", command=self._abrir_xml).pack(side="left", padx=(6, 0))

        # NF-e (consulta independente do CT-e)
        f = self._card("Chave da NF-e  •  transportadora da nota", fill="x", padx=20, pady=(0, 10))
        linha = tk.Frame(f, bg=_t.CARD_BG)
        linha.pack(fill="x")
        ent_nfe = tk.Entry(linha, textvariable=self.nfe_var, font=("Consolas", 11))
        ent_nfe.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        ent_nfe.bind("<Return>", lambda _e: self._buscar_nfe())
        ttk.Button(linha, text="Consultar NF-e", style="Primary.TButton",
                   command=self._buscar_nfe).pack(side="left")

        self.area = tk.Frame(self, bg=_t.BG_APP)
        self.area.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        self._mensagem("Cole a chave de 44 dígitos e clique em Buscar.", _t.FG_MUTED)

    # ── certificados ─────────────────────────────────────────────────────────

    def _atualizar_certificados(self):
        self.tv_cert.delete(*self.tv_cert.get_children())
        for c in dados.carregar_config()["certificados"]:
            v = c["validade"]
            validade = f"{v[8:10]}/{v[5:7]}/{v[:4]}"
            situacao = "VENCIDO" if c.get("vencido") else c.get("situacao", "")
            self.tv_cert.insert("", "end", iid=c["cnpj"], values=(
                c["nome"], dados.formatar_cnpj(c["cnpj"]), len(dados.estabelecimentos_de(c)), validade, situacao,
                c.get("ultima_sync") or "—"))
        self.lbl_cache.config(text=f"{self.cache.total()} CT-e no cache local")

    def _adicionar_cert(self):
        pfx = filedialog.askopenfilename(parent=self, title="Selecione o certificado digital (.pfx / .p12)",
                                         filetypes=[("Certificado", "*.pfx *.p12"), ("Todos", "*.*")])
        if not pfx:
            return
        senha = simpledialog.askstring("Senha do certificado", "Senha do certificado:",
                                       show="●", parent=self)
        if not senha:
            return
        lembrar = bool(dados.keyring) and messagebox.askyesno(
            "Lembrar senha", "Guardar a senha no Gerenciador de Credenciais do Windows,\n"
                             "para não precisar digitar de novo?", parent=self)
        try:
            cert = dados.adicionar_certificado(pfx, senha, lembrar)
        except dados.CertificadoErro as e:
            messagebox.showerror("Certificado", str(e), parent=self)
            return
        self.senhas[cert["cnpj"]] = senha
        self._atualizar_certificados()
        self._sincronizar([cert])          # já testa o certificado e baixa os CT-e

    def _editar_estabelecimentos(self):
        sel = self.tv_cert.selection()
        if not sel:
            messagebox.showinfo("Estabelecimentos", "Selecione um certificado na lista.", parent=self)
            return
        cnpj_cert = sel[0]
        cert = next(c for c in dados.carregar_config()["certificados"] if c["cnpj"] == cnpj_cert)

        top = tk.Toplevel(self)
        top.title("Estabelecimentos do certificado")
        top.geometry("460x520")
        top.configure(bg=_t.BG_APP)
        top.transient(self)
        top.grab_set()
        tk.Label(top, text=(f"{cert['nome']}\n\nUm CNPJ por linha (pode colar a coluna do Excel).\n"
                            f"Só são aceitos CNPJs da mesma raiz ({cnpj_cert[:8]}). O CNPJ do próprio\n"
                            "certificado sempre é consultado."),
                 justify="left", anchor="w", bg=_t.BG_APP, fg=_t.FG_TEXT).pack(fill="x", padx=14, pady=(12, 6))
        txt = tk.Text(top, font=("Consolas", 10), height=16)
        txt.pack(fill="both", expand=True, padx=14)
        txt.insert("1.0", "\n".join(dados.formatar_cnpj(c) for c in dados.estabelecimentos_de(cert)))

        def salvar():
            linhas = [l for l in txt.get("1.0", "end").splitlines() if l.strip()]
            aceitos, recusados = dados.definir_estabelecimentos(cnpj_cert, linhas)
            top.destroy()
            self._atualizar_certificados()
            msg = f"{len(aceitos)} estabelecimento(s) cadastrado(s)."
            if recusados:
                msg += (f"\n\n{len(recusados)} linha(s) recusada(s) (CNPJ inválido ou de outra raiz):\n"
                        + "\n".join(recusados[:8]) + ("\n…" if len(recusados) > 8 else ""))
            msg += "\n\nClique em “Atualizar CT-e da SEFAZ” para baixar os documentos deles."
            messagebox.showinfo("Estabelecimentos", msg, parent=self)

        rod = tk.Frame(top, bg=_t.BG_APP)
        rod.pack(fill="x", padx=14, pady=10)
        ttk.Button(rod, text="Salvar", style="Primary.TButton", command=salvar).pack(side="right")
        ttk.Button(rod, text="Cancelar", command=top.destroy).pack(side="right", padx=(0, 6))

    def _alternar_certs(self, restaurar=False):
        """Recolhe/expande a lista de certificados (os botões continuam visíveis)."""
        cfg = dados.carregar_config()
        recolhido = cfg.get("ui", {}).get("certs_recolhido", False)
        if not restaurar:
            recolhido = not recolhido
            cfg.setdefault("ui", {})["certs_recolhido"] = recolhido
            dados.salvar_config(cfg)
        if recolhido:
            self.tv_cert.pack_forget()
            self.btn_recolher.config(text="▸ Mostrar lista")
        else:
            self.tv_cert.pack(fill="x", before=self.linha_certs)
            self.btn_recolher.config(text="▾ Recolher")

    def _remover_cert(self):
        sel = self.tv_cert.selection()
        if not sel:
            messagebox.showinfo("Remover", "Selecione um certificado na lista.", parent=self)
            return
        if messagebox.askyesno("Remover", "Remover este certificado do app?\n"
                                          "(Os CT-e já baixados continuam no cache.)", parent=self):
            dados.remover_certificado(sel[0])
            self.senhas.pop(sel[0], None)
            self._atualizar_certificados()

    def _senha_de(self, cert) -> str | None:
        senha = self.senhas.get(cert["cnpj"]) or dados.senha_salva(cert["cnpj"])
        while True:
            if not senha:
                senha = simpledialog.askstring("Senha do certificado",
                                               f"Senha de {cert['nome']}:", show="●", parent=self)
                if not senha:
                    return None
            try:
                dados.ler_certificado(cert["arquivo"], senha)
            except dados.CertificadoErro as e:
                messagebox.showerror("Certificado", str(e), parent=self)
                senha = None
                continue
            self.senhas[cert["cnpj"]] = senha
            return senha

    def _atualizar_todos(self):
        certs = dados.carregar_config()["certificados"]
        if not certs:
            messagebox.showinfo("Atualizar", "Adicione um certificado primeiro.", parent=self)
            return
        self._sincronizar(certs)

    def _sincronizar(self, certs):
        if self.sincronizando:
            return
        trabalho = []
        for c in certs:
            if c.get("vencido"):
                continue
            senha = self._senha_de(c)
            if senha:
                trabalho.append((c, senha))
        if not trabalho:
            return
        self.sincronizando = True
        self.btn_atualizar.config(state="disabled")
        threading.Thread(target=self._thread_sync, args=(trabalho,), daemon=True).start()

    def _thread_sync(self, trabalho):
        total = 0
        for cert, senha in trabalho:
            r = dados.sincronizar(cert, senha, self.cache,
                                  progresso=lambda m: self.fila.put(("status", m)))
            total += r["novos"]
            self.fila.put(("refresh", None))
        self.fila.put(("fim", total))

    def _processar_fila(self):
        if not self.winfo_exists():
            return
        try:
            while True:
                tipo, valor = self.fila.get_nowait()
                if tipo == "status":
                    self.lbl_status.config(text=valor)
                elif tipo == "refresh":
                    self._atualizar_certificados()
                elif tipo == "nota":
                    token, idx, texto, diverge = valor
                    if token == getattr(self, "_token", None):
                        try:
                            self._tv_nota_set(idx, texto, diverge)
                        except tk.TclError:
                            pass        # a tabela já foi substituída
                elif tipo == "nfe_ok":
                    self._mostrar_nfe(valor)
                elif tipo == "nfe_erro":
                    self._mensagem(valor, _t.ACCENT_ORANGE)
                elif tipo == "fim":
                    self.sincronizando = False
                    self.btn_atualizar.config(state="normal")
                    self.lbl_status.config(text=f"Concluído: {valor} CT-e novo(s).")
                    self._atualizar_certificados()
        except queue.Empty:
            pass
        self.after(150, self._processar_fila)

    # ── busca ────────────────────────────────────────────────────────────────

    def _buscar(self):
        chave = dados.normalizar_chave(self.chave_var.get())
        if len(chave) != 44:
            self._mensagem(f"A chave deve ter 44 dígitos (você digitou {len(chave)}).", _t.ACCENT_ORANGE)
            return
        if chave[20:22] in ("55", "65"):          # chave de NF-e: usa a consulta própria da NF-e
            self.nfe_var.set(chave)
            self._buscar_nfe(chave)
            return
        if not dados.dv_chave_ok(chave):
            self._mensagem("Dígito verificador da chave inválido — confira se ela foi copiada inteira.", _t.ACCENT_ORANGE)
            return
        self._mostrar_chave(chave)

    def _mostrar_chave(self, chave, nota_foco=None, outros=None):
        xml = self.cache.obter(chave)
        if xml is None:
            info = dados.decodificar_chave(chave)
            self._mensagem(
                "CT-e não encontrado no cache.\n\n"
                f"Pela chave: CT-e nº {info['numero']}, série {info['serie']}, UF {info['uf']}, "
                f"emissão {info['emissao_aamm']}, emitente {dados.formatar_cnpj(info['cnpj_emitente'])}.\n\n"
                "Clique em “Atualizar CT-e da SEFAZ” e busque de novo. Se continuar sem aparecer, "
                "nenhum certificado cadastrado é parte deste CT-e (ou ele tem mais de ~3 meses).",
                _t.ACCENT_ORANGE)
            return
        self._mostrar(dados.extrair_dados(etree.fromstring(xml)), nota_foco, outros)

    # ── NF-e: transportadora da nota ─────────────────────────────────────────

    def _buscar_nfe(self, chave=None):
        chave = dados.normalizar_chave(chave or self.nfe_var.get())
        if len(chave) != 44:
            self._mensagem(f"A chave da NF-e deve ter 44 dígitos (você digitou {len(chave)}).", _t.ACCENT_ORANGE)
            return
        if not dados.dv_chave_ok(chave):
            self._mensagem("Dígito verificador da chave inválido — confira se ela foi copiada inteira.", _t.ACCENT_ORANGE)
            return
        if chave[20:22] not in ("55", "65"):
            self._mensagem("Esta chave não é de NF-e (modelo 55). Para CT-e use o campo de cima.", _t.ACCENT_ORANGE)
            return
        certs = []
        for c in dados.carregar_config()["certificados"]:
            if c.get("vencido"):
                continue
            senha = self._senha_de(c)
            if senha:
                certs.append((c, senha))
        if not certs:
            self._mensagem("Adicione (e libere a senha de) ao menos um certificado para consultar a NF-e.", _t.ACCENT_ORANGE)
            return
        self._mensagem("Consultando a SEFAZ…", _t.FG_MUTED)
        threading.Thread(target=self._thread_nfe, args=(chave, certs), daemon=True).start()

    def _thread_nfe(self, chave, certs):
        try:
            r = dados.consultar_nfe(chave, certs, self.cache,
                                    progresso=lambda m: self.fila.put(("status", m)))
            self.fila.put(("nfe_ok", r))
        except dados.NfeNaoEncontrada as e:
            self.fila.put(("nfe_erro", str(e)))
        except Exception as e:
            self.fila.put(("nfe_erro", f"Erro ao consultar a NF-e: {e}"))

    def _mostrar_nfe(self, d):
        self._limpar_area()
        pf = dados.formatar_cnpj
        outer = tk.Frame(self.area, bg=_t.GRID_LINE)
        outer.pack(fill="x", pady=(0, 10))
        tk.Label(outer, text="NF-E  •  TRANSPORTADORA DA NOTA", font=("Segoe UI", 8, "bold"),
                 bg=_t.GRID_LINE, fg=_t.ACCENT_CYAN, padx=12, pady=5).pack(anchor="w")
        f = tk.Frame(outer, bg=_t.CARD_BG, padx=12, pady=10)
        f.pack(fill="x", padx=1, pady=(0, 1))

        t = d["transportadora"]
        if t["nome"]:
            nome = f"{t['nome']}\n{pf(t['doc'])}" if t["doc"] else t["nome"]
        else:
            nome = "Não informada na nota"
        self._campo(f, 0, 0, "Transportadora", nome, _t.ACCENT_CYAN if t["nome"] else _t.ACCENT_ORANGE, span=3)
        self._campo(f, 0, 3, "Frete (modalidade)", d["mod_frete_txt"] or "—",
                    _t.ACCENT_ORANGE if d["mod_frete"] == "1" else _t.ACCENT_GREEN, span=2)
        self._campo(f, 1, 0, "Município / UF", t["municipio"], span=1)
        self._campo(f, 1, 1, "IE da transportadora", t["ie"], span=1)
        self._campo(f, 1, 2, "Veículo", d["veiculo"], span=1)
        self._campo(f, 1, 3, "Volumes", str(d["volumes"]) if d["volumes"] else "—", span=1)
        self._campo(f, 1, 4, "Peso bruto (kg)", f"{d['peso_bruto']:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    if d["peso_bruto"] else "—", span=1)

        self._campo(f, 2, 0, "Nota", d["numero"])
        self._campo(f, 2, 1, "Série", d["serie"])
        self._campo(f, 2, 2, "Data de emissão", dados.formatar_data(d["data"]))
        self._campo(f, 2, 3, "Valor da nota", "R$ " + dados.formatar_valor(d["valor"]), span=2)

        self._campo(f, 3, 0, "Emitente", f"{d['emitente']['nome']}\n{pf(d['emitente']['doc'])}", span=2)
        self._campo(f, 3, 2, "Destinatário", f"{d['destinatario']['nome']}\n{pf(d['destinatario']['doc'])}", span=2)
        cancelada = d["status"] not in ("100", "")
        self._campo(f, 3, 4, "Situação", "Autorizada" if not cancelada else f"{d['status']} {d['motivo']}",
                    _t.ACCENT_RED if cancelada else _t.ACCENT_GREEN)
        self._campo(f, 4, 0, "Chave de acesso", d["chave"], span=5)

    # ── transportadora de cada nota do CT-e (consulta automática da NF-e) ────

    def _consultar_notas(self, d, tv):
        """Consulta, em segundo plano, a transportadora declarada em cada NF-e do CT-e."""
        self._token = getattr(self, "_token", 0) + 1
        token = self._token
        if not d["notas"]:
            return
        todos = [c for c in dados.carregar_config()["certificados"] if not c.get("vencido")]
        # a empresa que recebeu este CT-e é parte das NF-e dele: começa pelos certificados
        # da mesma raiz de CNPJ das partes do CT-e (evita pedir senha de todos)
        raizes = {d[p]["doc"][:8] for p in ("remetente", "destinatario", "tomador", "emitente")
                  if d[p]["doc"]}
        certs = [c for c in todos if c["cnpj"][:8] in raizes] or todos
        liberados = []
        for c in certs:
            senha = self._senha_de(c)
            if senha:
                liberados.append((c, senha))
        if not liberados:
            for i in range(len(d["notas"])):
                tv.set(str(i), "transp", "— (libere a senha de um certificado)")
            return
        threading.Thread(target=self._thread_notas,
                         args=(token, d["notas"], liberados, d["emitente"]["doc"]),
                         daemon=True).start()

    def _thread_notas(self, token, notas, certs, emitente_cte):
        limite = False
        for i, n in enumerate(notas):
            diverge = False
            if limite:
                texto = "— (limite da SEFAZ, tente em 1h)"
            else:
                try:
                    r = dados.consultar_nfe(n["chave"], certs, self.cache)
                    t = r["transportadora"]
                    if t["nome"]:
                        texto = t["nome"] + (f" — {dados.formatar_cnpj(t['doc'])}" if t["doc"] else "")
                        diverge = not dados.mesma_empresa(t["doc"], emitente_cte)
                    else:
                        texto = "Não informada na nota"
                except dados.NfeNaoEncontrada as e:
                    msg = str(e)
                    if "resumo" in msg:
                        texto = "— só resumo (exige Ciência da Operação)"
                    elif "limite" in msg:
                        texto, limite = "— (limite da SEFAZ, tente em 1h)", True
                    else:
                        texto = "— NF-e não localizada"
                except Exception as e:
                    texto = f"— erro: {e}"[:80]
            self.fila.put(("nota", (token, i, texto, diverge)))

    def _tv_nota_set(self, idx, texto, diverge=False):
        self._tv_notas.set(str(idx), "transp", texto)
        self._tv_notas.item(str(idx), tags=("diverge",) if diverge else ())

    def _ativar_copia(self, tv, d):
        """Copiar a chave da NF-e da tabela: botão direito, Ctrl+C ou duplo clique."""
        chaves = {str(i): n["chave"] for i, n in enumerate(d["notas"])}
        aviso = tk.Label(tv.master, bg=_t.CARD_BG, fg=_t.FG_MUTED, font=("Segoe UI", 8), anchor="w",
                         text="Botão direito, duplo clique ou Ctrl+C na linha para copiar a chave da NF-e")
        aviso.pack(anchor="w", pady=(6, 0))
        padrao = aviso.cget("text")

        def copiar(texto, rotulo):
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()
            aviso.config(text=f"✔ {rotulo}", fg=_t.ACCENT_GREEN)
            aviso.after(2500, lambda: aviso.winfo_exists() and aviso.config(text=padrao, fg=_t.FG_MUTED))

        def selecionadas():
            return [chaves[i] for i in tv.selection() if i in chaves]

        def copiar_selecao(_e=None):
            sel = selecionadas()
            if sel:
                copiar("\n".join(sel), "Chave copiada" if len(sel) == 1 else f"{len(sel)} chaves copiadas")
            return "break"

        menu = tk.Menu(tv, tearoff=0)
        menu.add_command(label="Copiar chave da NF-e", command=copiar_selecao)
        menu.add_command(label="Copiar todas as chaves",
                         command=lambda: copiar("\n".join(chaves.values()), f"{len(chaves)} chaves copiadas"))
        menu.add_command(label="Copiar transportadora da nota",
                         command=lambda: tv.selection() and copiar(tv.set(tv.selection()[0], "transp"),
                                                                    "Transportadora copiada"))

        def abrir_menu(e):
            linha = tv.identify_row(e.y)
            if linha:
                tv.selection_set(linha)
                menu.tk_popup(e.x_root, e.y_root)

        def duplo_clique(e):
            if tv.identify_row(e.y):
                copiar_selecao()

        tv.bind("<Button-3>", abrir_menu)
        tv.bind("<Control-c>", copiar_selecao)
        tv.bind("<Control-C>", copiar_selecao)
        tv.bind("<Double-1>", duplo_clique)

    def _abrir_xml(self):
        caminho = filedialog.askopenfilename(parent=self, title="Selecione o XML do CT-e",
                                             filetypes=[("XML", "*.xml"), ("Todos", "*.*")])
        if not caminho:
            return
        try:
            self._mostrar(dados.extrair_dados(etree.parse(caminho).getroot()))
        except Exception as e:
            self._mensagem(f"Não consegui ler esse XML como CT-e: {e}", _t.ACCENT_RED)

    # ── exibição ─────────────────────────────────────────────────────────────

    def _limpar_area(self):
        for w in self.area.winfo_children():
            w.destroy()

    def _mensagem(self, texto, cor):
        self._limpar_area()
        tk.Label(self.area, text=texto, bg=_t.BG_APP, fg=cor, font=("Segoe UI", 10),
                 justify="left", wraplength=900).pack(anchor="w", pady=10)

    def _campo(self, parent, linha, coluna, titulo, valor, cor=None, span=1):
        f = tk.Frame(parent, bg=_t.CARD_BG)
        f.grid(row=linha, column=coluna, columnspan=span, sticky="w", padx=(0, 28), pady=6)
        tk.Label(f, text=titulo.upper(), font=("Segoe UI", 8, "bold"), bg=_t.CARD_BG,
                 fg=_t.FG_MUTED).pack(anchor="w")
        tk.Label(f, text=valor or "—", font=("Segoe UI", 11, "bold"), bg=_t.CARD_BG,
                 fg=cor or _t.FG_TEXT, justify="left").pack(anchor="w")

    def _mostrar(self, d, nota_foco=None, outros=None):
        self._limpar_area()
        pf = dados.formatar_cnpj

        topo = tk.Frame(self.area, bg=_t.GRID_LINE)
        topo.pack(fill="x", pady=(0, 10))
        f = tk.Frame(topo, bg=_t.CARD_BG, padx=12, pady=10)
        f.pack(fill="x", padx=1, pady=1)

        cancelado = d["status"] not in ("100", "")
        self._campo(f, 0, 0, "Conhecimento", d["numero"])
        self._campo(f, 0, 1, "Série", d["serie"])
        self._campo(f, 0, 2, "Data de emissão", dados.formatar_data(d["data"]))
        self._campo(f, 0, 3, "Valor do serviço", "R$ " + dados.formatar_valor(d["valor"]))
        self._campo(f, 0, 4, "Frete", d["frete"] or "—",
                    _t.ACCENT_ORANGE if d["frete"] == "FOB" else _t.ACCENT_GREEN)

        self._campo(f, 1, 0, "Fornecedor (transportadora)",
                    f"{d['emitente']['nome']}\n{pf(d['emitente']['doc'])}", span=2)
        self._campo(f, 1, 2, "Remetente", d["remetente"]["nome"])
        self._campo(f, 1, 3, "Destinatário", d["destinatario"]["nome"], span=2)

        rotulo = {"rem": "Remetente", "dest": "Destinatário", "exped": "Expedidor",
                  "receb": "Recebedor"}.get(d["tomador_tipo"], "Outro")
        self._campo(f, 2, 0, "Tomador do serviço",
                    f"{rotulo} — {d['tomador']['nome']}\n{pf(d['tomador']['doc'])}", span=2)
        self._campo(f, 2, 2, "CFOP / Natureza", f"{d['cfop']} — {d['natureza']}", span=2)
        self._campo(f, 2, 4, "Situação",
                    "Autorizado" if not cancelado else f"{d['status']} {d['motivo']}",
                    _t.ACCENT_RED if cancelado else _t.ACCENT_GREEN)

        self._campo(f, 3, 0, "Origem → Destino", f"{d['origem']}  →  {d['destino']}", span=2)
        self._campo(f, 3, 2, "Chave de acesso", d["chave"], span=3)

        notas = tk.Frame(self.area, bg=_t.GRID_LINE)
        notas.pack(fill="both", expand=True)
        tk.Label(notas, text=f"NOTAS VINCULADAS ({len(d['notas'])})  •  EM VERMELHO: TRANSPORTADORA DA NF-E DIFERENTE DO FORNECEDOR DO CT-E", font=("Segoe UI", 8, "bold"),
                 bg=_t.GRID_LINE, fg=_t.ACCENT_CYAN, padx=12, pady=5).pack(anchor="w")
        corpo = tk.Frame(notas, bg=_t.CARD_BG, padx=12, pady=10)
        corpo.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        tv = ttk.Treeview(corpo, columns=("nota", "serie", "transp", "chave"), show="headings",
                          height=max(2, min(len(d["notas"]), 8)))
        for col, txt, larg in [("nota", "Nota", 90), ("serie", "Série", 60),
                               ("transp", "Transportadora da nota (NF-e)", 360), ("chave", "Chave de acesso", 360)]:
            tv.heading(col, text=txt)
            tv.column(col, width=larg, anchor="w")
        tv.tag_configure("diverge", foreground="#FF6B6B", background="#3A1D24")
        for i, n in enumerate(d["notas"]):
            tv.insert("", "end", iid=str(i), values=(n["nota"], n["serie"], "consultando…", n["chave"]))
        tv.pack(fill="both", expand=True)
        self._tv_notas = tv
        self._consultar_notas(d, tv)
        self._ativar_copia(tv, d)
