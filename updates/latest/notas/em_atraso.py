# -*- coding: utf-8 -*-
"""
Ferramenta para reduzir a base de Notas Fiscais e enviar por email.

Fluxo:
1. Usuário seleciona o arquivo .xlsx de origem (base completa).
2. O programa filtra as linhas com STATUS = "AUTORIZADA" e cujo
   número de dias em atraso (hoje - EMISSAO) esteja entre 83 e 90
   (intervalo inclusive, configurável no topo do arquivo).
3. Gera uma nova planilha .xlsx apenas com as colunas solicitadas.
4. Envia essa planilha por email (Outlook/Office365 - SMTP) para o
   destinatário fixo, anexando o arquivo gerado.

Requisitos:
    pip install openpyxl

Autor: gerado com apoio do Claude
"""

import os
import smtplib
import ssl
import threading
from datetime import datetime, date
from email.message import EmailMessage

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Tema do app principal — usa as cores do core/theme.py quando disponível.
# Se o arquivo for rodado sozinho (fora do app), cai num tema padrão simples.
try:
    from core.theme import T, registrar_callback
except Exception:
    def T(chave: str) -> str:
        _FALLBACK = {
            "BG": "#0f1521", "PAINEL": "#161d2c", "BORDA": "#2a3344",
            "ACENTO": "#5b9bf5", "ACENTO2": "#3d7fe0", "TEXTO": "#e3e8f0",
            "TEXTO_SEC": "#7c8aa0", "VERDE": "#4ec98c", "AMARELO": "#e0a23c",
            "VERMELHO": "#e0586c",
        }
        return _FALLBACK.get(chave, "#e3e8f0")

    def registrar_callback(fn) -> None:
        pass

# --------------------------------------------------------------------------- #
# CONFIGURAÇÕES FIXAS - ajuste aqui se algo mudar no futuro
# --------------------------------------------------------------------------- #

# Intervalo de dias em atraso que deve entrar na planilha reduzida
DIAS_ATRASO_MIN = 60
DIAS_ATRASO_MAX = 90

# Status que deve ser considerado
STATUS_VALIDO = "AUTORIZADA"

# Colunas da planilha de ORIGEM que serão usadas (nomes devem bater com o
# cabeçalho real do arquivo .xlsx exportado)
COL_FILIAL = "FILIAL"
COL_DESC_FORNECEDOR = "DESC FORNECEDOR"
COL_NOTA = "NOTA"
COL_SERIE = "SERIE"
COL_VALOR_LIQUIDO = "VALOR LIQUIDO"
COL_EMISSAO = "EMISSAO"
COL_CHAVE = "CHAVE"
COL_CFOP = "CFOP"
COL_DESC_CFOP = "DESC CFOP"
COL_NATUREZA_OPERACAO = "NATUREZA OPERAÇÃO"
COL_STATUS = "STATUS"

# Ordem e nomes finais das colunas na planilha REDUZIDA
COLUNAS_SAIDA = [
    "FILIAL",
    "DESC FORNECEDOR",
    "NOTA",
    "SERIE",
    "VALOR LIQUIDO",
    "EMISSAO",
    "DIAS EM ATRASO",
    "CHAVE",
    "CFOP",
    "DESC CFOP",
    "NATUREZA OPERAÇÃO",
    "STATUS",
]

# Configurações de email (Office365 / Outlook)
SMTP_SERVIDOR = "smtp.office365.com"
SMTP_PORTA = 587
EMAIL_REMETENTE = "joao.teles@grupopmz.com.br"
EMAIL_DESTINATARIO = "fiscal@grupopmz.com.br"
EMAIL_ASSUNTO = "Notas Fiscais em Atraso ({}-{} dias) - {}".format(
    DIAS_ATRASO_MIN, DIAS_ATRASO_MAX, date.today().strftime("%d/%m/%Y")
)


# --------------------------------------------------------------------------- #
# LÓGICA DE PROCESSAMENTO
# --------------------------------------------------------------------------- #

def normalizar_texto(valor):
    """Remove espaços extras e deixa em maiúsculo para comparação segura."""
    if valor is None:
        return ""
    return str(valor).strip().upper()


def obter_data_emissao(valor):
    """Converte o valor da célula EMISSAO em um objeto date, se possível."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    # Tenta interpretar strings em formatos comuns
    texto = str(valor).strip()
    formatos = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def processar_planilha(caminho_origem, log_callback=None):
    """
    Lê o arquivo de origem, filtra e retorna:
        (lista_de_linhas_filtradas, caminho_arquivo_gerado)
    lista_de_linhas_filtradas é usada apenas para exibir contagem/preview.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("Abrindo arquivo de origem...")
    wb_origem = openpyxl.load_workbook(caminho_origem, data_only=True)
    ws_origem = wb_origem.active

    linhas = list(ws_origem.iter_rows(values_only=True))
    if not linhas:
        raise ValueError("A planilha de origem está vazia.")

    cabecalho = [normalizar_texto(c) if c else "" for c in linhas[0]]

    # Mapear nome da coluna -> índice, comparando de forma normalizada
    mapa_colunas = {}
    colunas_necessarias = {
        COL_FILIAL, COL_DESC_FORNECEDOR, COL_NOTA, COL_SERIE,
        COL_VALOR_LIQUIDO, COL_EMISSAO, COL_CHAVE, COL_CFOP,
        COL_DESC_CFOP, COL_NATUREZA_OPERACAO, COL_STATUS,
    }
    for idx, nome in enumerate(cabecalho):
        mapa_colunas[nome] = idx

    faltando = [c for c in colunas_necessarias if normalizar_texto(c) not in mapa_colunas]
    if faltando:
        raise ValueError(
            "As seguintes colunas não foram encontradas na planilha de origem: "
            + ", ".join(faltando)
        )

    def idx(nome_coluna):
        return mapa_colunas[normalizar_texto(nome_coluna)]

    hoje = date.today()
    resultado = []

    log("Filtrando linhas (status AUTORIZADA e atraso entre {} e {} dias)...".format(
        DIAS_ATRASO_MIN, DIAS_ATRASO_MAX
    ))

    for linha in linhas[1:]:
        if linha is None or all(v is None for v in linha):
            continue

        status = normalizar_texto(linha[idx(COL_STATUS)])
        if status != STATUS_VALIDO:
            continue

        data_emissao = obter_data_emissao(linha[idx(COL_EMISSAO)])
        if data_emissao is None:
            continue

        dias_atraso = (hoje - data_emissao).days
        if not (DIAS_ATRASO_MIN <= dias_atraso <= DIAS_ATRASO_MAX):
            continue

        linha_saida = {
            "FILIAL": linha[idx(COL_FILIAL)],
            "DESC FORNECEDOR": linha[idx(COL_DESC_FORNECEDOR)],
            "NOTA": linha[idx(COL_NOTA)],
            "SERIE": linha[idx(COL_SERIE)],
            "VALOR LIQUIDO": linha[idx(COL_VALOR_LIQUIDO)],
            "EMISSAO": data_emissao,
            "DIAS EM ATRASO": dias_atraso,
            "CHAVE": linha[idx(COL_CHAVE)],
            "CFOP": linha[idx(COL_CFOP)],
            "DESC CFOP": linha[idx(COL_DESC_CFOP)],
            "NATUREZA OPERAÇÃO": linha[idx(COL_NATUREZA_OPERACAO)],
            "STATUS": linha[idx(COL_STATUS)],
        }
        resultado.append(linha_saida)

    log("{} linha(s) encontrada(s) dentro dos critérios.".format(len(resultado)))

    caminho_saida = gerar_arquivo_saida(resultado, caminho_origem)
    return resultado, caminho_saida


def gerar_arquivo_saida(linhas, caminho_origem):
    """Gera o arquivo .xlsx reduzido formatado e retorna o caminho salvo."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Notas em Atraso"

    fonte_padrao = "Arial"

    # Cabeçalho
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fonte_cabecalho = Font(name=fonte_padrao, bold=True, color="FFFFFF", size=11)

    for col_num, nome_coluna in enumerate(COLUNAS_SAIDA, start=1):
        celula = ws.cell(row=1, column=col_num, value=nome_coluna)
        celula.font = fonte_cabecalho
        celula.fill = fill_cabecalho
        celula.alignment = Alignment(horizontal="center", vertical="center")

    # Dados
    fonte_dados = Font(name=fonte_padrao, size=10)
    for row_num, linha in enumerate(linhas, start=2):
        for col_num, nome_coluna in enumerate(COLUNAS_SAIDA, start=1):
            valor = linha.get(nome_coluna)
            celula = ws.cell(row=row_num, column=col_num, value=valor)
            celula.font = fonte_dados
            if nome_coluna == "EMISSAO":
                celula.number_format = "DD/MM/YYYY"
            if nome_coluna == "VALOR LIQUIDO":
                celula.number_format = "#,##0.00"

    # Ajuste automático (aproximado) da largura das colunas
    for col_num, nome_coluna in enumerate(COLUNAS_SAIDA, start=1):
        letra = get_column_letter(col_num)
        maior = len(nome_coluna)
        for linha in linhas:
            valor = linha.get(nome_coluna)
            if valor is not None:
                maior = max(maior, len(str(valor)))
        ws.column_dimensions[letra].width = min(max(maior + 2, 10), 45)

    ws.freeze_panes = "A2"

    pasta_origem = os.path.dirname(os.path.abspath(caminho_origem))
    nome_arquivo = "Notas_Em_Atraso_{}.xlsx".format(date.today().strftime("%Y%m%d"))
    caminho_saida = os.path.join(pasta_origem, nome_arquivo)

    wb.save(caminho_saida)
    return caminho_saida


# --------------------------------------------------------------------------- #
# ENVIO DE EMAIL
# --------------------------------------------------------------------------- #

def enviar_email(caminho_anexo, senha, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("Conectando ao servidor {}...".format(SMTP_SERVIDOR))

    msg = EmailMessage()
    msg["Subject"] = EMAIL_ASSUNTO
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = EMAIL_DESTINATARIO
    msg.set_content(
        "Prezados,\n\n"
        "Segue em anexo a planilha com as notas fiscais autorizadas que estão "
        "com atraso entre {} e {} dias, referente a {}.\n\n"
        "Att,\nJoão Teles".format(
            DIAS_ATRASO_MIN, DIAS_ATRASO_MAX, date.today().strftime("%d/%m/%Y")
        )
    )

    with open(caminho_anexo, "rb") as f:
        conteudo = f.read()
    nome_anexo = os.path.basename(caminho_anexo)
    msg.add_attachment(
        conteudo,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=nome_anexo,
    )

    contexto = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as servidor:
        servidor.starttls(context=contexto)
        log("Autenticando...")
        servidor.login(EMAIL_REMETENTE, senha)
        log("Enviando email...")
        servidor.send_message(msg)

    log("Email enviado com sucesso para {}.".format(EMAIL_DESTINATARIO))


# --------------------------------------------------------------------------- #
# INTERFACE GRÁFICA (TKINTER)
# --------------------------------------------------------------------------- #

class NotasEmAtrasoFrame(tk.Frame):
    """Frame embutível — usado dentro de um Toplevel do app principal,
    seguindo o tema (core/theme.py). Também pode rodar sozinho (ver o
    bloco no final do arquivo)."""

    def __init__(self, master):
        super().__init__(master, bg=T("BG"))

        self.caminho_origem = tk.StringVar()
        self.caminho_saida = None
        self.linhas_filtradas = []

        self._montar_layout()
        registrar_callback(self._reaplicar_tema)

    # ------------------------------------------------------------------ #
    def _montar_layout(self):
        pad = {"padx": 12, "pady": 6}

        self.lbl_titulo = tk.Label(
            self,
            text="Redutor de Notas Fiscais em Atraso",
            font=("Segoe UI Semibold", 14),
            bg=T("BG"), fg=T("ACENTO"),
        )
        self.lbl_titulo.pack(pady=(16, 4))

        self.lbl_subtitulo = tk.Label(
            self,
            text="Status = AUTORIZADA  |  Atraso entre {} e {} dias".format(
                DIAS_ATRASO_MIN, DIAS_ATRASO_MAX
            ),
            font=("Segoe UI", 10),
            bg=T("BG"), fg=T("TEXTO_SEC"),
        )
        self.lbl_subtitulo.pack(pady=(0, 12))

        # --- Seleção de arquivo ---------------------------------------- #
        self.frame_arquivo = tk.LabelFrame(
            self, text="1. Arquivo de origem (base completa)", padx=10, pady=10,
            bg=T("PAINEL"), fg=T("TEXTO"),
        )
        self.frame_arquivo.pack(fill="x", **pad)

        linha_arquivo = tk.Frame(self.frame_arquivo, bg=T("PAINEL"))
        linha_arquivo.pack(fill="x")
        self.linha_arquivo = linha_arquivo

        self.entry_arquivo = tk.Entry(
            linha_arquivo, textvariable=self.caminho_origem, state="readonly",
            readonlybackground=T("BG"), fg=T("TEXTO"), relief="flat",
        )
        self.entry_arquivo.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=3)

        self.btn_selecionar = tk.Button(
            linha_arquivo, text="Selecionar arquivo...", command=self.selecionar_arquivo,
            bg=T("BORDA"), fg=T("TEXTO"), relief="flat", cursor="hand2",
            activebackground=T("ACENTO2"), activeforeground="#ffffff",
        )
        self.btn_selecionar.pack(side="right")

        # --- Processamento ------------------------------------------------ #
        self.frame_processo = tk.LabelFrame(
            self, text="2. Gerar planilha reduzida", padx=10, pady=10,
            bg=T("PAINEL"), fg=T("TEXTO"),
        )
        self.frame_processo.pack(fill="x", **pad)

        self.btn_processar = tk.Button(
            self.frame_processo, text="Gerar planilha reduzida", command=self.processar,
            state="disabled", bg=T("BORDA"), fg=T("TEXTO"), relief="flat", cursor="hand2",
            activebackground=T("ACENTO2"), activeforeground="#ffffff",
        )
        self.btn_processar.pack(anchor="w")

        self.label_resultado = tk.Label(
            self.frame_processo, text="Nenhum arquivo gerado ainda.",
            bg=T("PAINEL"), fg=T("TEXTO_SEC"),
        )
        self.label_resultado.pack(anchor="w", pady=(6, 0))

        # --- Envio de email ------------------------------------------------ #
        self.frame_email = tk.LabelFrame(
            self, text="3. Enviar por email", padx=10, pady=10,
            bg=T("PAINEL"), fg=T("TEXTO"),
        )
        self.frame_email.pack(fill="x", **pad)

        self.lbl_de = tk.Label(self.frame_email, text="De: {}".format(EMAIL_REMETENTE),
                                bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        self.lbl_de.grid(row=0, column=0, sticky="w")
        self.lbl_para = tk.Label(self.frame_email, text="Para: {}".format(EMAIL_DESTINATARIO),
                                  bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        self.lbl_para.grid(row=1, column=0, sticky="w")

        self.lbl_senha = tk.Label(self.frame_email, text="Senha do email:",
                                   bg=T("PAINEL"), fg=T("TEXTO"))
        self.lbl_senha.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.entry_senha = tk.Entry(
            self.frame_email, show="*", width=30,
            bg=T("BG"), fg=T("TEXTO"), insertbackground=T("TEXTO"), relief="flat",
        )
        self.entry_senha.grid(row=2, column=1, sticky="w", pady=(8, 0), ipady=3)

        self.btn_enviar = tk.Button(
            self.frame_email, text="Enviar email", command=self.enviar, state="disabled",
            bg=T("BORDA"), fg=T("TEXTO"), relief="flat", cursor="hand2",
            activebackground=T("ACENTO2"), activeforeground="#ffffff",
        )
        self.btn_enviar.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # --- Log ------------------------------------------------------------ #
        self.frame_log = tk.LabelFrame(
            self, text="Status", padx=10, pady=10, bg=T("PAINEL"), fg=T("TEXTO"),
        )
        self.frame_log.pack(fill="both", expand=True, **pad)

        self.text_log = tk.Text(
            self.frame_log, height=8, state="disabled", wrap="word",
            bg=T("BG"), fg=T("TEXTO"), insertbackground=T("TEXTO"),
            relief="flat", borderwidth=0,
        )
        self.text_log.pack(fill="both", expand=True)

    # ------------------------------------------------------------------ #
    def _reaplicar_tema(self):
        """Reaplica as cores quando o usuário troca de tema no app."""
        self.config(bg=T("BG"))
        self.lbl_titulo.config(bg=T("BG"), fg=T("ACENTO"))
        self.lbl_subtitulo.config(bg=T("BG"), fg=T("TEXTO_SEC"))

        for frame in (self.frame_arquivo, self.frame_processo, self.frame_email, self.frame_log):
            frame.config(bg=T("PAINEL"), fg=T("TEXTO"))

        self.linha_arquivo.config(bg=T("PAINEL"))
        self.entry_arquivo.config(readonlybackground=T("BG"), fg=T("TEXTO"))
        self.btn_selecionar.config(bg=T("BORDA"), fg=T("TEXTO"), activebackground=T("ACENTO2"))
        self.btn_processar.config(bg=T("BORDA"), fg=T("TEXTO"), activebackground=T("ACENTO2"))
        self.label_resultado.config(bg=T("PAINEL"))
        self.lbl_de.config(bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        self.lbl_para.config(bg=T("PAINEL"), fg=T("TEXTO_SEC"))
        self.lbl_senha.config(bg=T("PAINEL"), fg=T("TEXTO"))
        self.entry_senha.config(bg=T("BG"), fg=T("TEXTO"), insertbackground=T("TEXTO"))
        self.btn_enviar.config(bg=T("BORDA"), fg=T("TEXTO"), activebackground=T("ACENTO2"))
        self.text_log.config(bg=T("BG"), fg=T("TEXTO"), insertbackground=T("TEXTO"))

    # ------------------------------------------------------------------ #
    def log(self, mensagem):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", mensagem + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")
        self.update_idletasks()

    # ------------------------------------------------------------------ #
    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de origem",
            filetypes=[("Arquivos Excel", "*.xlsx")],
        )
        if caminho:
            self.caminho_origem.set(caminho)
            self.btn_processar.config(state="normal")
            self.label_resultado.config(text="Arquivo selecionado. Clique em 'Gerar planilha reduzida'.")
            self.btn_enviar.config(state="disabled")
            self.caminho_saida = None

    # ------------------------------------------------------------------ #
    def processar(self):
        caminho = self.caminho_origem.get()
        if not caminho:
            messagebox.showwarning("Atenção", "Selecione um arquivo primeiro.")
            return

        self.btn_processar.config(state="disabled")
        self.log("Iniciando processamento de: {}".format(caminho))

        def tarefa():
            try:
                linhas, caminho_saida = processar_planilha(caminho, log_callback=self.log)
                self.linhas_filtradas = linhas
                self.caminho_saida = caminho_saida

                if len(linhas) == 0:
                    self.label_resultado.config(
                        text="Nenhuma nota encontrada dentro do critério de atraso.",
                        fg=T("AMARELO"),
                    )
                    self.btn_enviar.config(state="disabled")
                else:
                    self.label_resultado.config(
                        text="{} nota(s) encontrada(s). Arquivo salvo em:\n{}".format(
                            len(linhas), caminho_saida
                        ),
                        fg=T("VERDE"),
                    )
                    self.btn_enviar.config(state="normal")
            except Exception as exc:
                self.log("ERRO: {}".format(exc))
                messagebox.showerror("Erro ao processar", str(exc))
            finally:
                self.btn_processar.config(state="normal")

        threading.Thread(target=tarefa, daemon=True).start()

    # ------------------------------------------------------------------ #
    def enviar(self):
        if not self.caminho_saida or not os.path.exists(self.caminho_saida):
            messagebox.showwarning("Atenção", "Gere a planilha reduzida antes de enviar.")
            return

        senha = self.entry_senha.get()
        if not senha:
            messagebox.showwarning("Atenção", "Digite a senha do email antes de enviar.")
            return

        confirmar = messagebox.askyesno(
            "Confirmar envio",
            "Enviar o arquivo:\n{}\n\nPara: {}\n\nConfirma?".format(
                os.path.basename(self.caminho_saida), EMAIL_DESTINATARIO
            ),
        )
        if not confirmar:
            return

        self.btn_enviar.config(state="disabled")

        def tarefa():
            try:
                enviar_email(self.caminho_saida, senha, log_callback=self.log)
                messagebox.showinfo("Sucesso", "Email enviado com sucesso!")
                self.entry_senha.delete(0, "end")
            except smtplib.SMTPAuthenticationError as exc:
                codigo = getattr(exc, "smtp_code", "?")
                detalhe = getattr(exc, "smtp_error", b"")
                if isinstance(detalhe, bytes):
                    detalhe = detalhe.decode("utf-8", errors="replace")
                self.log("ERRO DE AUTENTICAÇÃO (código {}):".format(codigo))
                self.log(detalhe)
                messagebox.showerror(
                    "Erro de autenticação",
                    "O servidor recusou a autenticação.\n\n"
                    "Código: {}\nDetalhe do servidor:\n{}\n\n"
                    "Veja a janela de Status para o texto completo. "
                    "Normalmente indica que é preciso usar uma 'senha de "
                    "aplicativo' (MFA ativo) ou que o SMTP AUTH está "
                    "desabilitado pelo TI.".format(codigo, detalhe[:300]),
                )
            except smtplib.SMTPException as exc:
                self.log("ERRO SMTP: {} - {}".format(type(exc).__name__, exc))
                messagebox.showerror("Erro ao enviar email", "{}: {}".format(type(exc).__name__, exc))
            except Exception as exc:
                self.log("ERRO: {} - {}".format(type(exc).__name__, exc))
                messagebox.showerror("Erro ao enviar email", str(exc))
            finally:
                self.btn_enviar.config(state="normal")

        threading.Thread(target=tarefa, daemon=True).start()


# --------------------------------------------------------------------------- #
# INTEGRAÇÃO COM O MENU PRINCIPAL
# --------------------------------------------------------------------------- #

def _build_content(win):
    """Monta a tela dentro de uma janela (Toplevel) já criada pelo
    menu_principal.py, seguindo o padrão dos outros módulos do app."""
    win.title("Redutor de Notas Fiscais em Atraso")
    win.geometry("620x560")
    win.minsize(560, 480)
    win.configure(bg=T("BG"))

    frame = NotasEmAtrasoFrame(win)
    frame.pack(fill="both", expand=True)
    return frame


if __name__ == "__main__":
    # Execução independente (fora do app principal), útil para testes.
    root = tk.Tk()
    root.title("Redutor de Notas Fiscais em Atraso")
    root.geometry("620x560")
    root.minsize(560, 480)
    root.configure(bg=T("BG"))
    NotasEmAtrasoFrame(root).pack(fill="both", expand=True)
    root.mainloop()