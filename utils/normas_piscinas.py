"""Critérios resumidos da ABNT NBR 10818:2026 para avaliação RT.

Faixas ideais são distintas dos limites normativos. Uma leitura isolada não
comprova a frequência de monitoramento nem a conformidade global da piscina.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Faixa:
    minimo: float | None
    maximo: float | None
    ideal_min: float | None = None
    ideal_max: float | None = None
    maximo_exclusivo: bool = False


# Tratamento simples com produtos clorados em piscina coletiva.
# Para tratamento complementar, CRL admite 0,5 ppm como mínimo; não inferir
# essa condição sem cadastro explícito do tratamento empregado.
FAIXAS_COLETIVA_SIMPLES = {
    "ph": Faixa(7.0, 7.8, 7.4, 7.6),
    "crl": Faixa(1.0, 5.0, 2.0, 4.0),
    "alcalinidade": Faixa(60.0, 180.0, 80.0, 120.0),
    "dureza": Faixa(100.0, 400.0, 180.0, 250.0),
    "cya": Faixa(None, 100.0, 30.0, 50.0),
    "cc": Faixa(0.0, 0.6),
    "turbidez": Faixa(0.0, 0.5, maximo_exclusivo=True),
    "orp": Faixa(650.0, None),
    "lsi": Faixa(-0.3, 0.5),
}


def avaliar_leitura(campo: str, valor: float | None, *, tratamento_complementar=False) -> str:
    """Retorna ausente, fora_limite, fora_ideal ou dentro_ideal.

    Não classifica CT isoladamente, TDS sem linha de base/gerador ou outros
    parâmetros sem contexto normativo suficiente.
    """
    if valor is None:
        return "ausente"
    faixa = FAIXAS_COLETIVA_SIMPLES.get(campo)
    if faixa is None:
        return "sem_avaliacao"
    minimo = 0.5 if campo == "crl" and tratamento_complementar else faixa.minimo
    if (minimo is not None and valor < minimo) or (
        faixa.maximo is not None and
        (valor >= faixa.maximo if faixa.maximo_exclusivo else valor > faixa.maximo)
    ):
        return "fora_limite"
    if (faixa.ideal_min is not None and valor < faixa.ideal_min) or (
        faixa.ideal_max is not None and valor > faixa.ideal_max
    ):
        return "fora_ideal"
    return "dentro_ideal"


def cloro_combinado(crl: float | None, ct: float | None) -> tuple[float | None, str | None]:
    if crl is None or ct is None:
        return None, None
    if crl < 0 or ct < 0 or ct < crl:
        return None, "Cloro total inferior ao livre: conferir medição ou lançamento."
    return round(ct - crl, 6), None


NORMA_QUALIDADE = "ABNT NBR 10818:2026"
REFERENCIA_FAIXAS = (
    "ABNT NBR 10818:2026, Tabela 3, piscina coletiva com tratamento clorado simples"
)

# Rótulos e unidades usados nos documentos. Mantidos aqui para que todos os
# geradores (PDF/DOCX/HTML) mostrem os mesmos limites.
UNIDADES = {
    "ph": "",
    "crl": "mg/L",
    "alcalinidade": "mg/L CaCO3",
    "dureza": "mg/L CaCO3",
    "cya": "mg/L",
    "cc": "mg/L",
    "turbidez": "NTU",
    "orp": "mV",
    "lsi": "",
}

# Aliases de campos encontrados nos formulários, planilhas e lançamentos.
ALIASES = {
    "ph": "ph",
    "crl": "crl", "cloro_livre": "crl", "cl": "crl",
    "alcalinidade": "alcalinidade", "alc": "alcalinidade",
    "dureza": "dureza", "dc": "dureza",
    "cya": "cya", "cianurico": "cya",
    "cc": "cc", "cloro_combinado": "cc",
    "turbidez": "turbidez", "ntu": "turbidez",
    "orp": "orp",
    "lsi": "lsi",
}


def _num_br(v: float) -> str:
    if abs(v) >= 10 and float(v).is_integer():
        return str(int(v))
    return f"{v:.1f}".replace(".", ",")


def campo_normativo(campo: str) -> str | None:
    return ALIASES.get(str(campo or "").strip().lower())


def limites(campo: str, *, tratamento_complementar: bool = False) -> tuple[float | None, float | None]:
    """Limites aceitáveis (não as metas ideais) do parâmetro."""
    chave = campo_normativo(campo)
    faixa = FAIXAS_COLETIVA_SIMPLES.get(chave) if chave else None
    if faixa is None:
        return None, None
    minimo = 0.5 if chave == "crl" and tratamento_complementar else faixa.minimo
    return minimo, faixa.maximo


def texto_limite(campo: str, *, tratamento_complementar: bool = False) -> str:
    """Faixa aceitável em texto pt-BR, p.ex. "7,0–7,8", "até 100", "< 0,5".

    Usa apenas caracteres WinAnsi para renderizar com as fontes padrão do PDF.
    """
    chave = campo_normativo(campo)
    faixa = FAIXAS_COLETIVA_SIMPLES.get(chave) if chave else None
    if faixa is None:
        return "—"
    mn, mx = limites(chave, tratamento_complementar=tratamento_complementar)
    if chave in ("cc",):
        return f"até {_num_br(mx)}"
    if faixa.maximo_exclusivo and mx is not None:
        return f"< {_num_br(mx)}"
    if mn is None and mx is not None:
        return f"até {_num_br(mx)}"
    if mx is None and mn is not None:
        return f"mín. {_num_br(mn)}"
    if mn < 0:
        return f"-{_num_br(abs(mn))} a +{_num_br(mx)}"
    return f"{_num_br(mn)}–{_num_br(mx)}"


def texto_ideal(campo: str) -> str:
    chave = campo_normativo(campo)
    faixa = FAIXAS_COLETIVA_SIMPLES.get(chave) if chave else None
    if faixa is None or faixa.ideal_min is None or faixa.ideal_max is None:
        return "—"
    return f"{_num_br(faixa.ideal_min)}–{_num_br(faixa.ideal_max)}"


def fora_limite(campo: str, valor: float | None, *, tratamento_complementar: bool = False) -> bool:
    chave = campo_normativo(campo)
    if chave is None or valor is None:
        return False
    return avaliar_leitura(chave, valor, tratamento_complementar=tratamento_complementar) == "fora_limite"


# ── ORP: regra única dos relatórios do escritório e de campo ───────────────
# Critério: ABNT NBR 10818:2026, seção 6 / Tabela 3, piscina coletiva com tratamento
# clorado simples: "ORP mínimo 650 mV quando aplicável" (conferência do texto integral
# registrada em docs/RT_NORMAS_20261101.md). O material aprovado não define quando o
# ORP é aplicável nem atribui não conformidade à leitura abaixo da referência. Por isso:
# - abaixo de 650 mV: ponto de atenção; não altera o status nem entra nas não conformidades;
# - 650 mV ou mais (inclusive o limite): atende a referência;
# - sem leitura: ausente; nunca vira zero, nem conformidade, nem leitura avaliável.
ORP_AUSENTE = "ausente"
ORP_ABAIXO_REFERENCIA = "abaixo_referencia"
ORP_ATENDE_REFERENCIA = "atende_referencia"
RECOMENDACAO_ORP = ("Quando o ORP for aplicável ao tratamento, conferir a demanda oxidante e repetir a "
                    "leitura com o medidor calibrado.")


def ler_orp(valor) -> float | None:
    """Leitura de ORP em mV, aceitando "650", "650,5", "1.234,5"; vazio/traço/texto → None (nunca 0)."""
    if isinstance(valor, bool) or valor is None:
        return None
    if isinstance(valor, (int, float)):
        return None if valor != valor else float(valor)  # NaN = sem leitura
    s = str(valor).strip().replace(" ", "")
    if s.lower().endswith("mv"):
        s = s[:-2]
    if not s or s in {"—", "-", "–", "None", "nan"}:
        return None
    if s.count(",") == 1 and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return None
    return None if v != v else v


def avaliar_orp(valor) -> str:
    """ausente | abaixo_referencia | atende_referencia (ver regra acima)."""
    v = ler_orp(valor)
    if v is None:
        return ORP_AUSENTE
    return ORP_ABAIXO_REFERENCIA if fora_limite("orp", v) else ORP_ATENDE_REFERENCIA


def texto_atencao_orp(valor, onde: str = "") -> str:
    """Texto do ponto de atenção (mesmo nos dois relatórios); `onde` já vem formatado (" — data, piscina")."""
    v = ler_orp(valor)
    fmt = lambda x: f"{x:g}".replace(".", ",")  # noqa: E731
    return (f"ORP={fmt(v)} mV abaixo da referência mín. {fmt(limites('orp')[0])} mV "
            f"(NBR 10818:2026, quando aplicável) — ponto de atenção{onde}")


def texto_resumo_limites() -> str:
    """Linha única com os limites usados nos relatórios (sem alegar conformidade)."""
    partes = [
        f"pH {texto_limite('ph')} (ideal {texto_ideal('ph')})",
        f"CRL {texto_limite('crl')} mg/L (ideal {texto_ideal('crl')})",
        f"CC {texto_limite('cc')} mg/L",
        f"alcalinidade {texto_limite('alcalinidade')} mg/L",
        f"dureza cálcica {texto_limite('dureza')} mg/L",
        f"CYA {texto_limite('cya')} mg/L",
        f"turbidez {texto_limite('turbidez')} NTU",
        f"ORP {texto_limite('orp')} mV, quando aplicável",
    ]
    return (
        f"Limites de referência - {REFERENCIA_FAIXAS}: " + "; ".join(partes) +
        ". Tratamento complementar admite CRL mínimo de 0,5 mg/L somente quando cadastrado. "
        "TDS e temperatura não são classificados por leitura isolada."
    )
