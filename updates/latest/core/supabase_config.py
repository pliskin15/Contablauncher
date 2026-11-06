"""
core/supabase_config.py
Cliente Supabase compartilhado.

O cliente é criado de forma lazy (na primeira chamada a get_supabase()),
evitando falha na importação quando o .env ainda não foi carregado
ou quando o módulo é importado fora do contexto do app.
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from supabase import create_client, Client

def _localizar_env() -> Path:
    """Procura o .env ao lado do .exe (release), no bundle e na raiz do projeto."""
    candidatos = []
    if getattr(sys, "frozen", False):
        candidatos.append(Path(sys.executable).resolve().parent / ".env")
        if hasattr(sys, "_MEIPASS"):
            candidatos.append(Path(sys._MEIPASS) / ".env")
    candidatos.append(Path(__file__).resolve().parent.parent / ".env")
    for c in candidatos:
        if c.is_file():
            return c
    return candidatos[0]


_ENV_PATH = _localizar_env()
load_dotenv(dotenv_path=_ENV_PATH)

_client: Client | None = None


def get_supabase() -> Client:
    """Retorna o cliente Supabase, criando-o na primeira chamada."""
    global _client
    if _client is None:
        url = os.getenv("SUPABASE_URL")
        key = os.getenv("SUPABASE_ANON_KEY")
        if not url or not key:
            raise EnvironmentError(
                "SUPABASE_URL e SUPABASE_ANON_KEY não encontrados.\n"
                f"Verifique o arquivo .env em: {_ENV_PATH}"
            )
        _client = create_client(url, key)
    return _client


# Compatibilidade: módulos que fazem  `from core.supabase_config import supabase`
# continuam funcionando — supabase agora é um proxy lazy.
class _LazyClient:
    """Proxy que encaminha qualquer acesso para o cliente real."""
    def __getattr__(self, name):
        return getattr(get_supabase(), name)


supabase = _LazyClient()
