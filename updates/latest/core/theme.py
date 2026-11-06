"""
theme.py
────────
Sistema de tema unificado do Cofre.

Absorve ui_theme.py — todos os módulos devem importar apenas este arquivo.
Para retrocompatibilidade, ui_theme.py agora é um shim que re-exporta daqui.

Duas camadas de cor coexistem:
  • Constantes fixas (PRIMARY, BG_APP, CARD_BG …)  — usadas pelos módulos
    Brinks, Prossegur e Depósitos, que têm layout estático.
  • Paleta dinâmica via T("CHAVE")  — usada pelos menus e conciliações,
    que suportam troca de tema em tempo de execução.
"""

import tkinter as tk
from tkinter import ttk


# ══════════════════════════════════════════════════════════════════════════════
#  1. CONSTANTES FIXAS  (antigas ui_theme.py)
#     Usadas como:  theme.PRIMARY, theme.BG_APP, theme.CARD_BG …
# ══════════════════════════════════════════════════════════════════════════════

# Paleta do Prisma V22 (theme.py do Prisma): azul-carvão, acento gelo e ciano de assinatura.
PRIMARY        = "#294c60"      # botão principal / faixas (ACENTO2 do Prisma)
PRIMARY_HOVER  = "#33607a"
PRIMARY_ACTIVE = "#1f3b4b"

BG_APP    = "#0f1117"           # fundo
CARD_BG   = "#181c27"           # cards e painéis (SURFACE)
SURFACE2  = "#1e2333"           # inputs, linhas alternadas, botão secundário
FG_TEXT   = "#dde3f0"
FG_MUTED  = "#7a8aaa"
FG_DIM    = "#3a4460"           # desabilitado
GRID_LINE = "#2a3050"           # bordas (BORDA)
SEL_BG    = "#1f2a48"           # linha selecionada

ACCENT        = "#adb6c5"       # acento principal (texto/realce)
ACCENT_CYAN   = "#00d4ff"       # assinatura: títulos de seção e cabeçalhos
ACCENT_BLUE   = "#4f7cff"
ACCENT_GREEN  = "#00c97a"       # OK / conciliado
ACCENT_ORANGE = "#ffb830"       # atenção
ACCENT_RED    = "#ff4d6a"       # divergência / erro

FONT_BASE  = ("Segoe UI", 10)
FONT_SEMI  = ("Segoe UI Semibold", 10)
FONT_TITLE = ("Segoe UI", 15, "bold")


def apply_theme(root) -> dict:
    """
    Aplica o tema escuro fixo a uma janela Toplevel ou Tk.
    Retorna a paleta de cores para uso direto (ex.: self.palette = apply_theme(self)).
    """
    root.configure(bg=BG_APP)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(
        ".", background=BG_APP, foreground=FG_TEXT,
        fieldbackground=CARD_BG, font=FONT_BASE,
    )

    # Widgets Tk "clássicos" (Entry, Text, Listbox, menus…) criados DEPOIS desta chamada
    # nascem no tema; quem define cores próprias continua mandando.
    for opcao, valor in {
        "*Entry.background": SURFACE2, "*Entry.foreground": FG_TEXT,
        "*Entry.insertBackground": FG_TEXT, "*Entry.selectBackground": SEL_BG,
        "*Entry.selectForeground": FG_TEXT, "*Entry.relief": "flat",
        "*Entry.highlightThickness": 1, "*Entry.highlightBackground": GRID_LINE,
        "*Entry.highlightColor": ACCENT,
        "*Spinbox.background": SURFACE2, "*Spinbox.foreground": FG_TEXT,
        "*Spinbox.insertBackground": FG_TEXT, "*Spinbox.relief": "flat",
        "*Text.background": CARD_BG, "*Text.foreground": FG_TEXT,
        "*Text.insertBackground": FG_TEXT, "*Text.selectBackground": SEL_BG,
        "*Text.selectForeground": FG_TEXT,
        "*Listbox.background": SURFACE2, "*Listbox.foreground": FG_TEXT,
        "*Listbox.selectBackground": SEL_BG, "*Listbox.selectForeground": FG_TEXT,
        "*Listbox.relief": "flat", "*Listbox.highlightThickness": 0,
        "*Menubutton.background": SURFACE2, "*Menubutton.foreground": FG_TEXT,
        "*Menubutton.activeBackground": GRID_LINE, "*Menubutton.activeForeground": FG_TEXT,
        "*Menubutton.relief": "flat",
        "*Menu.background": CARD_BG, "*Menu.foreground": FG_TEXT,
        "*Menu.activeBackground": SEL_BG, "*Menu.activeForeground": FG_TEXT,
        "*TCombobox*Listbox.background": SURFACE2, "*TCombobox*Listbox.foreground": FG_TEXT,
        "*TCombobox*Listbox.selectBackground": SEL_BG, "*TCombobox*Listbox.selectForeground": FG_TEXT,
    }.items():
        root.option_add(opcao, valor)

    style.configure("TProgressbar", troughcolor=SURFACE2, background=ACCENT_CYAN,
                    bordercolor=GRID_LINE, lightcolor=ACCENT_CYAN, darkcolor=ACCENT_CYAN,
                    thickness=8)

    # Frames
    style.configure("App.TFrame",     background=BG_APP)
    style.configure("Card.TFrame",    background=CARD_BG)
    style.configure("Toolbar.TFrame", background=CARD_BG)
    style.configure("Status.TFrame",  background=BG_APP)

    # Labels
    style.configure("TLabel",         background=BG_APP, foreground=FG_TEXT,  font=FONT_BASE)
    style.configure("Muted.TLabel",   background=BG_APP, foreground=FG_MUTED)
    style.configure("Section.TLabel", background=BG_APP, foreground=FG_TEXT,  font=FONT_SEMI)
    style.configure("Toolbar.TLabel", background=CARD_BG, foreground=FG_TEXT,
                    font=("Segoe UI Semibold", 11))
    style.configure("Title.TLabel",   background=BG_APP, foreground=FG_TEXT,  font=FONT_TITLE)

    # Buttons
    style.configure("TButton",
                    background=SURFACE2, foreground=FG_TEXT,
                    bordercolor=GRID_LINE, lightcolor=SURFACE2, darkcolor=SURFACE2,
                    borderwidth=1, padding=(12, 8))
    style.map("TButton",
              background=[("active", GRID_LINE), ("pressed", GRID_LINE)])

    style.configure("Primary.TButton",
                    background=PRIMARY, foreground="white",
                    borderwidth=0, padding=(14, 8))
    style.map("Primary.TButton",
              background=[("active", PRIMARY_HOVER), ("pressed", PRIMARY_ACTIVE)])

    style.configure("Ghost.TButton",
                    background=BG_APP, foreground=FG_TEXT,
                    borderwidth=0, padding=(8, 6))
    style.map("Ghost.TButton",
              background=[("active", SURFACE2)])

    # Combobox
    style.configure("TCombobox",
                    fieldbackground=SURFACE2, background=SURFACE2,
                    foreground=FG_TEXT, arrowcolor=FG_MUTED, bordercolor=GRID_LINE)
    style.map("TCombobox",
              fieldbackground=[("readonly", SURFACE2), ("focus", SURFACE2)],
              bordercolor=[("focus", ACCENT)],
              arrowcolor=[("active", FG_TEXT)])

    # Treeview
    style.configure("Treeview",
                    background=CARD_BG, fieldbackground=CARD_BG,
                    foreground=FG_TEXT, bordercolor=GRID_LINE,
                    lightcolor=GRID_LINE, darkcolor=GRID_LINE, rowheight=30)
    style.configure("Treeview.Heading",
                    background=BG_APP, foreground=ACCENT_CYAN,
                    font=("Segoe UI", 8, "bold"), padding=(8, 7),
                    borderwidth=0, relief="flat")
    style.map("Treeview.Heading", background=[("active", SURFACE2)])
    style.map("Treeview",
              background=[("selected", SEL_BG)],
              foreground=[("selected", FG_TEXT)])

    # Entry
    style.configure("TEntry",
                    fieldbackground=SURFACE2, background=SURFACE2,
                    foreground=FG_TEXT, bordercolor=GRID_LINE,
                    insertcolor=FG_TEXT, padding=(8, 6))
    style.map("TEntry",
              fieldbackground=[("disabled", BG_APP), ("focus", CARD_BG)],
              foreground=[("disabled", FG_DIM)],
              bordercolor=[("focus", ACCENT)],
              lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])

    # LabelFrame
    style.configure("TLabelframe",
                    background=BG_APP, bordercolor=GRID_LINE,
                    lightcolor=GRID_LINE, darkcolor=GRID_LINE)
    style.configure("TLabelframe.Label",
                    background=BG_APP, foreground=FG_MUTED, font=FONT_SEMI)
    style.configure("Card.TLabelframe",
                    background=CARD_BG, bordercolor=GRID_LINE,
                    lightcolor=GRID_LINE, darkcolor=GRID_LINE)
    style.configure("Card.TLabelframe.Label",
                    background=CARD_BG, foreground=FG_MUTED, font=FONT_SEMI)

    # Notebook
    style.configure("TNotebook", background=BG_APP, borderwidth=0)
    style.configure("TNotebook.Tab",
                    background=BG_APP, foreground=FG_MUTED,
                    bordercolor=GRID_LINE, lightcolor=BG_APP, darkcolor=BG_APP,
                    padding=(20, 9), font=FONT_SEMI, borderwidth=1)
    style.map("TNotebook.Tab",
              background=[("selected", CARD_BG), ("active", SURFACE2)],
              foreground=[("selected", ACCENT_CYAN), ("active", FG_TEXT)],
              expand=[("selected", (0, 0, 0, 0))])

    # Checkbutton (usado em filtros eventuais)
    style.configure("TCheckbutton",
                    background=BG_APP, foreground=FG_TEXT, font=FONT_BASE)
    style.map("TCheckbutton", background=[("active", BG_APP)])

    # Scrollbar
    style.configure("Vertical.TScrollbar",
                    background=GRID_LINE, troughcolor=BG_APP,
                    bordercolor=BG_APP, arrowcolor=FG_MUTED)
    style.map("Vertical.TScrollbar",
              background=[("active", ACCENT)])
    style.configure("Horizontal.TScrollbar",
                    background=GRID_LINE, troughcolor=BG_APP,
                    bordercolor=BG_APP, arrowcolor=FG_MUTED)
    style.map("Horizontal.TScrollbar",
              background=[("active", ACCENT)])

    # Separator
    style.configure("TSeparator", background=GRID_LINE)

    return {
        "PRIMARY":   PRIMARY,
        "BG_APP":    BG_APP,
        "CARD_BG":   CARD_BG,
        "FG_TEXT":   FG_TEXT,
        "FG_MUTED":  FG_MUTED,
        "GRID_LINE": GRID_LINE,
        "GREEN":     ACCENT_GREEN,
        "ORANGE":    ACCENT_ORANGE,
        "RED":       ACCENT_RED,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  2. PALETA DINÂMICA  (antigas theme.py)
#     Usadas como:  T("BG"), T("ACENTO"), alternar_tema(), botao_tema() …
# ══════════════════════════════════════════════════════════════════════════════

_TEMAS: dict[str, dict[str, str]] = {
    "prisma": {
        "BG":        "#0f1117",
        "PAINEL":    "#181c27",
        "BORDA":     "#2a3050",
        "ACENTO":    "#adb6c5",
        "ACENTO2":   "#294c60",
        "TEXTO":     "#dde3f0",
        "TEXTO_SEC": "#7a8aaa",
        "VERDE":     "#00c97a",
        "AMARELO":   "#ffb830",
        "VERMELHO":  "#ff4d6a",
        "CINZA":     "#7a8aaa",
        "AZUL":      "#4f7cff",
        "LARANJA":   "#ffb830",
        "ROXO":      "#adb6c5",
        "CIANO":     "#00d4ff",
        "SURFACE2":  "#1e2333",
        "TEXTO_DES": "#3a4460",
        "TAG_CONCILIADO_BG": "#0f2a24",
        "TAG_CONCILIADO_FG": "#00c97a",
        "TAG_PARCIAL_BG":    "#2b2410",
        "TAG_PARCIAL_FG":    "#ffb830",
        "TAG_PENDENTE_BG":   "#2c1520",
        "TAG_PENDENTE_FG":   "#ff4d6a",
        "TAG_IGNORADO_BG":   "#181c27",
        "TAG_IGNORADO_FG":   "#7a8aaa",
        "TAG_SEL_BG":        "#1f2a48",
        "TAG_SEL_FG":        "#dde3f0",
        "TAG_ZEBRA_BG":      "#1e2333",
    },
    "original": {
        "BG":        "#0906bb",
        "PAINEL":    "#0b2cc2",
        "BORDA":     "#3a3a5c",
        "ACENTO":    "#e0e0f0",
        "ACENTO2":   "#0906bb",
        "TEXTO":     "#e0e0f0",
        "TEXTO_SEC": "#6d6d6d",
        "VERDE":     "#2ecc71",
        "AMARELO":   "#f39c12",
        "VERMELHO":  "#e74c3c",
        "CINZA":     "#7f8c8d",
        "AZUL":      "#3498db",
        "LARANJA":   "#e67e22",
        "ROXO":      "#ffffff",
        "TAG_CONCILIADO_BG": "#1a3a2a",
        "TAG_CONCILIADO_FG": "#2ecc71",
        "TAG_PARCIAL_BG":    "#3a3010",
        "TAG_PARCIAL_FG":    "#f39c12",
        "TAG_PENDENTE_BG":   "#3a1010",
        "TAG_PENDENTE_FG":   "#e74c3c",
        "TAG_IGNORADO_BG":   "#2a2a2a",
        "TAG_IGNORADO_FG":   "#7f8c8d",
        "TAG_SEL_BG":        "#1a1a5e",
        "TAG_SEL_FG":        "#ffffff",
        "TAG_ZEBRA_BG":      "#252535",
    },
    "escuro": {
        "BG":        "#0a0a0a",
        "PAINEL":    "#141414",
        "BORDA":     "#2a2a2a",
        "ACENTO":    "#ffffff",
        "ACENTO2":   "#bbbbbb",
        "TEXTO":     "#e8e8e8",
        "TEXTO_SEC": "#888888",
        "VERDE":     "#2ecc71",
        "AMARELO":   "#f39c12",
        "VERMELHO":  "#e74c3c",
        "CINZA":     "#666666",
        "AZUL":      "#5dade2",
        "LARANJA":   "#e67e22",
        "ROXO":      "#af7ac5",
        "TAG_CONCILIADO_BG": "#0d2118",
        "TAG_CONCILIADO_FG": "#2ecc71",
        "TAG_PARCIAL_BG":    "#1e1a00",
        "TAG_PARCIAL_FG":    "#f39c12",
        "TAG_PENDENTE_BG":   "#1e0808",
        "TAG_PENDENTE_FG":   "#e74c3c",
        "TAG_IGNORADO_BG":   "#161616",
        "TAG_IGNORADO_FG":   "#666666",
        "TAG_SEL_BG":        "#0d0d3a",
        "TAG_SEL_FG":        "#ffffff",
        "TAG_ZEBRA_BG":      "#111111",
    },
    "claro": {
        "BG":        "#f5f5f5",
        "PAINEL":    "#ffffff",
        "BORDA":     "#cccccc",
        "ACENTO":    "#4a3bbf",
        "ACENTO2":   "#6c5ce7",
        "TEXTO":     "#1a1a2e",
        "TEXTO_SEC": "#555577",
        "VERDE":     "#1a8a4a",
        "AMARELO":   "#c07a00",
        "VERMELHO":  "#c0392b",
        "CINZA":     "#6c7a7d",
        "AZUL":      "#1d6fa4",
        "LARANJA":   "#b85c00",
        "ROXO":      "#7d3c98",
        "TAG_CONCILIADO_BG": "#d5f5e3",
        "TAG_CONCILIADO_FG": "#1a6a35",
        "TAG_PARCIAL_BG":    "#fef9e7",
        "TAG_PARCIAL_FG":    "#7d5a00",
        "TAG_PENDENTE_BG":   "#fdedec",
        "TAG_PENDENTE_FG":   "#922b21",
        "TAG_IGNORADO_BG":   "#eaeaea",
        "TAG_IGNORADO_FG":   "#555555",
        "TAG_SEL_BG":        "#d2d8f7",
        "TAG_SEL_FG":        "#1a1a2e",
        "TAG_ZEBRA_BG":      "#eeeeee",
    },
}

_tema_atual    = "prisma"
_ORDEM_TEMAS   = ["prisma", "original", "escuro", "claro"]
_ICONES_TEMA   = {"prisma": "🔷", "original": "🟣", "escuro": "⚫", "claro": "⚪"}
_NOMES_TEMA    = {"prisma": "Prisma", "original": "Original", "escuro": "Escuro", "claro": "Claro"}
_callbacks_registrados: list = []


def T(chave: str) -> str:
    """Retorna a cor da chave no tema atual."""
    tema = _TEMAS[_tema_atual]
    if chave in tema:
        return tema[chave]
    # chaves extras (CIANO, SURFACE2, TEXTO_DES) só existem no tema "prisma"
    return {"CIANO": tema["ACENTO"], "SURFACE2": tema["PAINEL"], "TEXTO_DES": tema["TEXTO_SEC"]}[chave]


def tema_atual() -> str:
    """Retorna o nome do tema ativo ('original', 'escuro' ou 'claro')."""
    return _tema_atual


def aplicar_estilos_ttk(style: ttk.Style) -> None:
    """Aplica os estilos ttk da paleta dinâmica (menus e conciliações)."""
    style.theme_use("clam")
    style.configure(
        "Treeview",
        background=T("PAINEL"),
        fieldbackground=T("PAINEL"),
        foreground=T("TEXTO"),
        rowheight=24,
        font=("Segoe UI", 8),
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "Treeview.Heading",
        background=T("BG"),
        foreground=T("CIANO"),
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padding=(6, 6),
        borderwidth=0,
    )
    style.map("Treeview.Heading", background=[("active", T("BORDA"))])
    style.map(
        "Treeview",
        background=[("selected", T("TAG_SEL_BG"))],
        foreground=[("selected", T("TAG_SEL_FG"))],
    )
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.configure(
        "TScrollbar",
        background=T("BORDA"),
        troughcolor=T("BG"),
        arrowcolor=T("TEXTO_SEC"),
        bordercolor=T("BG"),
        borderwidth=0,
        relief="flat",
    )
    style.map("TScrollbar", background=[("active", T("ACENTO"))])
    style.configure(
        "TCombobox",
        fieldbackground=T("SURFACE2"),
        background=T("SURFACE2"),
        foreground=T("TEXTO"),
        arrowcolor=T("TEXTO_SEC"),
        bordercolor=T("BORDA"),
        borderwidth=1,
        relief="flat",
        padding=4,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", T("SURFACE2"))],
        foreground=[("readonly", T("TEXTO"))],
    )


def aplicar_tags_tree(tree: ttk.Treeview) -> None:
    """Configura as tags de cor nas Treeviews da paleta dinâmica."""
    tree.tag_configure("conciliado",
                       background=T("TAG_CONCILIADO_BG"),
                       foreground=T("TAG_CONCILIADO_FG"))
    tree.tag_configure("parcial",
                       background=T("TAG_PARCIAL_BG"),
                       foreground=T("TAG_PARCIAL_FG"))
    tree.tag_configure("pendente",
                       background=T("TAG_PENDENTE_BG"),
                       foreground=T("TAG_PENDENTE_FG"))
    tree.tag_configure("ignorado",
                       background=T("TAG_IGNORADO_BG"),
                       foreground=T("TAG_IGNORADO_FG"))
    tree.tag_configure("selecionado",
                       background=T("TAG_SEL_BG"),
                       foreground=T("TAG_SEL_FG"))
    tree.tag_configure("zebra",
                       background=T("TAG_ZEBRA_BG"))


def registrar_callback(fn) -> None:
    """Registra função a ser chamada após cada troca de tema."""
    if fn not in _callbacks_registrados:
        _callbacks_registrados.append(fn)


def alternar_tema() -> str:
    """Avança para o próximo tema e dispara todos os callbacks registrados."""
    global _tema_atual
    idx = _ORDEM_TEMAS.index(_tema_atual)
    _tema_atual = _ORDEM_TEMAS[(idx + 1) % len(_ORDEM_TEMAS)]
    for fn in list(_callbacks_registrados):
        try:
            fn()
        except Exception as e:
            print(f"[theme] callback error: {e}")
    return _tema_atual


def botao_tema(parent: tk.Widget, callback=None) -> tk.Button:
    """Cria e retorna o botão de alternância de tema."""
    btn = tk.Button(
        parent,
        text=_label_tema(),
        bg=T("BORDA"),
        fg=T("TEXTO"),
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=10,
        pady=6,
        cursor="hand2",
    )

    def _trocar():
        alternar_tema()
        btn.config(text=_label_tema(), bg=T("BORDA"), fg=T("TEXTO"))
        if callback:
            callback()

    btn.config(command=_trocar)
    return btn


def _label_tema() -> str:
    proximo = _ORDEM_TEMAS[(_ORDEM_TEMAS.index(_tema_atual) + 1) % len(_ORDEM_TEMAS)]
    return f"{_ICONES_TEMA[proximo]} {_NOMES_TEMA[proximo]}"


def recolorir_widget(widget: tk.Widget,
                     bg_key: str = "BG",
                     fg_key: str = "TEXTO") -> None:
    """Reaplica cores de tema a um widget e todos os seus filhos recursivamente."""
    try:
        widget.config(bg=T(bg_key))
    except Exception:
        pass
    try:
        widget.config(fg=T(fg_key))
    except Exception:
        pass
    for child in widget.winfo_children():
        recolorir_widget(child, bg_key, fg_key)