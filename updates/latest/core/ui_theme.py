"""
ui_theme.py
───────────
SHIM de retrocompatibilidade — não edite este arquivo.

Todo o conteúdo foi migrado para theme.py.
Este módulo apenas re-exporta de lá para que imports antigos
do tipo  `import ui_theme as theme`  continuem funcionando
sem nenhuma alteração nos módulos consumidores.

Quando todos os módulos forem atualizados para importar theme
diretamente, este arquivo pode ser removido.
"""

from core.theme import (
    # Constantes de cor fixas
    PRIMARY,
    PRIMARY_HOVER,
    PRIMARY_ACTIVE,
    BG_APP,
    CARD_BG,
    FG_TEXT,
    FG_MUTED,
    GRID_LINE,
    SURFACE2,
    FG_DIM,
    SEL_BG,
    ACCENT,
    ACCENT_CYAN,
    ACCENT_BLUE,
    ACCENT_GREEN,
    ACCENT_ORANGE,
    ACCENT_RED,
    FONT_BASE,
    FONT_SEMI,
    FONT_TITLE,
    # Função de aplicação de tema
    apply_theme,
)

__all__ = [
    "PRIMARY", "PRIMARY_HOVER", "PRIMARY_ACTIVE",
    "BG_APP", "CARD_BG", "FG_TEXT", "FG_MUTED", "GRID_LINE",
    "SURFACE2", "FG_DIM", "SEL_BG", "ACCENT", "ACCENT_CYAN", "ACCENT_BLUE",
    "ACCENT_GREEN", "ACCENT_ORANGE", "ACCENT_RED",
    "FONT_BASE", "FONT_SEMI", "FONT_TITLE",
    "apply_theme",
]
