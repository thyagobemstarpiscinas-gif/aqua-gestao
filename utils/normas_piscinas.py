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
