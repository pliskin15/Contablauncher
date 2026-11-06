"""
core/versao.py — versão do programa (fonte única: o arquivo versao.txt).

O versao.txt fica ao lado do .exe (e da raiz do projeto em desenvolvimento) e é
atualizado pelo tools/release.py a cada release. As telas só chamam `texto()`.
"""
import os
import sys

_PADRAO = "2.1"


def _candidatos() -> list[str]:
    pastas = []
    if getattr(sys, "frozen", False):
        pastas.append(os.path.dirname(sys.executable))
        if hasattr(sys, "_MEIPASS"):
            pastas.append(sys._MEIPASS)
    pastas.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))   # raiz do projeto
    return [os.path.join(p, "versao.txt") for p in pastas]


def obter() -> str:
    """Versão atual, ex.: '1.9'. Se o arquivo não existir, devolve a versão padrão."""
    for caminho in _candidatos():
        try:
            with open(caminho, encoding="utf-8") as f:
                v = f.read().strip()
            if v:
                return v
        except OSError:
            continue
    return _PADRAO


def texto() -> str:
    """Versão pronta para exibir: 'v1.9'."""
    return f"v{obter()}"
