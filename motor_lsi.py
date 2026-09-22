"""Motor matematico isolado para calculo do indice de saturacao de Langelier."""

import math


FORMULA_VERSAO = "LSI_EPA_PHTA_1.0"
_FATORES_CYA = (
    (7.0, 0.23),
    (7.2, 0.27),
    (7.4, 0.31),
    (7.6, 0.33),
    (7.8, 0.35),
    (8.0, 0.36),
)


def _numero(valor, nome):
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise ValueError(f"{nome} deve ser um numero.")
    if not math.isfinite(valor):
        raise ValueError(f"{nome} deve ser um numero finito.")
    return float(valor)


def _fator_cya(ph):
    if ph <= _FATORES_CYA[0][0]:
        return _FATORES_CYA[0][1], ph < _FATORES_CYA[0][0]
    if ph >= _FATORES_CYA[-1][0]:
        return _FATORES_CYA[-1][1], ph > _FATORES_CYA[-1][0]

    for (ph_anterior, fator_anterior), (ph_proximo, fator_proximo) in zip(
        _FATORES_CYA, _FATORES_CYA[1:]
    ):
        if ph <= ph_proximo:
            proporcao = (ph - ph_anterior) / (ph_proximo - ph_anterior)
            return fator_anterior + proporcao * (fator_proximo - fator_anterior), False

    raise RuntimeError("Nao foi possivel determinar o fator de correcao do CYA.")


def _apresentar(valor):
    return round(valor, 5)


def calcular_lsi(
    ph,
    alcalinidade_total,
    dureza_calcica,
    temperatura_c,
    tds,
    cya=0,
    limite_inferior=-0.2,
    limite_superior=0.2,
):
    """Calcula o LSI e retorna os valores calculados em formato serializavel."""
    ph = _numero(ph, "pH")
    alcalinidade_total = _numero(alcalinidade_total, "Alcalinidade total")
    dureza_calcica = _numero(dureza_calcica, "Dureza calcica")
    temperatura_c = _numero(temperatura_c, "Temperatura")
    tds = _numero(tds, "TDS")
    cya = _numero(cya, "CYA")
    limite_inferior = _numero(limite_inferior, "Limite inferior")
    limite_superior = _numero(limite_superior, "Limite superior")

    if tds <= 0:
        raise ValueError("TDS deve ser maior que zero.")
    if dureza_calcica <= 0:
        raise ValueError("Dureza calcica deve ser maior que zero.")
    if alcalinidade_total <= 0:
        raise ValueError("Alcalinidade total deve ser maior que zero.")
    if cya < 0:
        raise ValueError("CYA nao pode ser negativo.")
    if limite_inferior > limite_superior:
        raise ValueError("O limite inferior nao pode ser maior que o limite superior.")
    if temperatura_c + 273 <= 0:
        raise ValueError("Temperatura deve ser maior que -273 graus Celsius.")

    fator_cya, fator_limitado = _fator_cya(ph)
    alcalinidade_carbonatada = alcalinidade_total - (cya * fator_cya)
    if alcalinidade_carbonatada <= 0:
        raise ValueError("Alcalinidade carbonatada deve ser maior que zero.")

    fator_tds_a = (math.log10(tds) - 1) / 10
    fator_temperatura_b = -13.12 * math.log10(temperatura_c + 273) + 34.55
    fator_dureza_c = math.log10(dureza_calcica) - 0.4
    fator_alcalinidade_d = math.log10(alcalinidade_carbonatada)
    phs = (9.3 + fator_tds_a + fator_temperatura_b) - (
        fator_dureza_c + fator_alcalinidade_d
    )
    lsi = ph - phs

    if lsi < limite_inferior:
        classificacao = "tendência subsaturada"
    elif lsi > limite_superior:
        classificacao = "tendência incrustante"
    else:
        classificacao = "próximo do equilíbrio"

    avisos = []
    if fator_limitado and cya > 0:
        avisos.append(
            "pH fora da faixa de 7.0 a 8.0: foi usado o fator CYA do extremo mais próximo."
        )

    return {
        "lsi": _apresentar(lsi),
        "phs": _apresentar(phs),
        "classificacao": classificacao,
        "alcalinidade_total": _apresentar(alcalinidade_total),
        "alcalinidade_carbonatada": _apresentar(alcalinidade_carbonatada),
        "fator_correcao_cya": _apresentar(fator_cya),
        "cya": _apresentar(cya),
        "fator_tds_a": _apresentar(fator_tds_a),
        "fator_temperatura_b": _apresentar(fator_temperatura_b),
        "fator_dureza_c": _apresentar(fator_dureza_c),
        "fator_alcalinidade_d": _apresentar(fator_alcalinidade_d),
        "formula_versao": FORMULA_VERSAO,
        "avisos": avisos,
        "parametros_utilizados": {
            "ph": _apresentar(ph),
            "temperatura_c": _apresentar(temperatura_c),
            "tds": _apresentar(tds),
            "dureza_calcica": _apresentar(dureza_calcica),
            "limite_inferior": _apresentar(limite_inferior),
            "limite_superior": _apresentar(limite_superior),
        },
    }