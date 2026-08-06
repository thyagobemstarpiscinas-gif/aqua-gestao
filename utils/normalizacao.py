import re
import unicodedata


def normalizar_texto_busca(valor: str) -> str:
    """Normaliza texto para comparação robusta: remove acento, caixa, espaços e símbolos."""
    valor = str(valor or "").strip().lower()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    valor = re.sub(r"[^a-z0-9]+", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def nomes_condominio_equivalentes(a: str, b: str) -> bool:
    """Compara nomes de condomínios tolerando acentos, espaços e pequenas variações."""
    na = normalizar_texto_busca(a)
    nb = normalizar_texto_busca(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na
