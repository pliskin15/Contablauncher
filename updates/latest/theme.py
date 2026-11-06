import tkinter as tk
from tkinter import ttk

_TEMAS = {
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
    "slate": {
        "BG":        "#0f1521",
        "PAINEL":    "#161d2c",
        "BORDA":     "#2a3344",
        "ACENTO":    "#5b9bf5",
        "ACENTO2":   "#3d7fe0",
        "TEXTO":     "#e3e8f0",
        "TEXTO_SEC": "#7c8aa0",
        "VERDE":     "#4ec98c",
        "AMARELO":   "#e0a23c",
        "VERMELHO":  "#e0586c",
        "CINZA":     "#5d6a80",
        "AZUL":      "#5b9bf5",
        "LARANJA":   "#e08a4c",
        "ROXO":      "#9d8cf0",
        "TAG_CONCILIADO_BG": "#132a22",
        "TAG_CONCILIADO_FG": "#4ec98c",
        "TAG_PARCIAL_BG":    "#2b2310",
        "TAG_PARCIAL_FG":    "#e0a23c",
        "TAG_PENDENTE_BG":   "#2b1418",
        "TAG_PENDENTE_FG":   "#e0586c",
        "TAG_IGNORADO_BG":   "#1a2030",
        "TAG_IGNORADO_FG":   "#5d6a80",
        "TAG_SEL_BG":        "#1c3a66",
        "TAG_SEL_FG":        "#ffffff",
        "TAG_ZEBRA_BG":      "#131a27",
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

_tema_atual = "slate"
_ORDEM_TEMAS = ["slate", "original", "escuro", "claro"]
_ICONES_TEMA = {"slate": "🔷", "original": "🟣", "escuro": "⚫", "claro": "⚪"}
_NOMES_TEMA  = {"slate": "Slate", "original": "Original", "escuro": "Escuro", "claro": "Claro"}

_callbacks_registrados: list = []


def T(chave: str) -> str:

    return _TEMAS[_tema_atual][chave]


def tema_atual() -> str:

    return _tema_atual

def aplicar_estilos_ttk(style: ttk.Style) -> None:

    style.theme_use("clam")
    style.configure(
        "Treeview",
        background=T("BG"),
        fieldbackground=T("BG"),
        foreground=T("TEXTO"),
        rowheight=24,
        font=("Segoe UI", 8),
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "Treeview.Heading",
        background=T("PAINEL"),
        foreground=T("ACENTO"),
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padding=(6, 5),
        borderwidth=0,
    )
    style.map(
        "Treeview.Heading",
        background=[("active", T("BORDA"))],
    )
    style.map(
        "Treeview",
        background=[("selected", T("ACENTO2"))],
        foreground=[("selected", "#ffffff")],
    )
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.configure(
        "TScrollbar",
        background=T("PAINEL"),
        troughcolor=T("BG"),
        arrowcolor=T("TEXTO_SEC"),
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "TCombobox",
        fieldbackground=T("BG"),
        background=T("BG"),
        foreground=T("TEXTO"),
        arrowcolor=T("TEXTO_SEC"),
        borderwidth=0,
        relief="flat",
        padding=4,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", T("BG"))],
        foreground=[("readonly", T("TEXTO"))],
    )


def aplicar_tags_tree(tree: ttk.Treeview) -> None:

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

    if fn not in _callbacks_registrados:
        _callbacks_registrados.append(fn)


def alternar_tema() -> str:

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


# ── Unificação ───────────────────────────────────────────────────────────────
# Este arquivo tinha paleta e estado de tema próprios, separados de core/theme.py
# (o Cofre e as conciliações usavam temas diferentes). Daqui em diante tudo vem de
# core/theme.py (identidade visual do Prisma V22); as definições acima ficam só por
# compatibilidade e são substituídas por estas.
from core.theme import (  # noqa: E402,F811
    T,
    alternar_tema,
    aplicar_estilos_ttk,
    aplicar_tags_tree,
    botao_tema,
    recolorir_widget,
    registrar_callback,
    tema_atual,
)