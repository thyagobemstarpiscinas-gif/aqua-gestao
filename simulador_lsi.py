"""Simulacao de cenarios e analise comparativa de sensibilidade do LSI."""

import math

from motor_lsi import calcular_lsi


_PARAMETROS = (
    "ph",
    "alcalinidade_total",
    "dureza_calcica",
    "temperatura_c",
    "tds",
    "cya",
)
_VARIACOES_PADRAO = {
    "ph": 0.10,
    "alcalinidade_total": 10,
    "dureza_calcica": 10,
    "temperatura_c": 1,
    "tds": 100,
    "cya": 10,
}
_AVISO_METODOLOGICO = (
    "Análise comparativa baseada nas amplitudes de variação informadas. "
    "Não representa previsão exata nem relação causal."
)


def _validar_parametros(parametros):
    if not isinstance(parametros, dict):
        raise ValueError("parametros_atuais deve ser um dicionário.")
    desconhecidas = set(parametros) - set(_PARAMETROS)
    if desconhecidas:
        nomes = ", ".join(sorted(desconhecidas))
        raise ValueError(f"Chave(s) desconhecida(s) em parametros_atuais: {nomes}.")
    ausentes = [nome for nome in _PARAMETROS if nome not in parametros]
    if ausentes:
        nomes = ", ".join(ausentes)
        raise ValueError(f"Parâmetro(s) ausente(s) em parametros_atuais: {nomes}.")
    return dict(parametros)


def _validar_alteracoes(alteracoes):
    if not isinstance(alteracoes, dict):
        raise ValueError("alteracoes deve ser um dicionário.")
    desconhecidas = set(alteracoes) - set(_PARAMETROS)
    if desconhecidas:
        nomes = ", ".join(sorted(desconhecidas))
        raise ValueError(f"Chave(s) desconhecida(s) em alteracoes: {nomes}.")


def _calcular(parametros, limite_inferior, limite_superior):
    return calcular_lsi(
        parametros["ph"],
        parametros["alcalinidade_total"],
        parametros["dureza_calcica"],
        parametros["temperatura_c"],
        parametros["tds"],
        parametros["cya"],
        limite_inferior=limite_inferior,
        limite_superior=limite_superior,
    )


def _resumo(cenario_atual, cenario_simulado):
    classificacao_atual = cenario_atual["classificacao"]
    classificacao_simulada = cenario_simulado["classificacao"]
    if classificacao_atual != classificacao_simulada:
        return (
            "O cenário simulado alterou a classificação de "
            f"{classificacao_atual} para {classificacao_simulada}."
        )
    if cenario_simulado["lsi"] > cenario_atual["lsi"]:
        return "O cenário simulado apresentou aumento do LSI."
    if cenario_simulado["lsi"] < cenario_atual["lsi"]:
        return "O cenário simulado apresentou redução do LSI."
    return "O cenário simulado permaneceu próximo do equilíbrio."


def simular_cenario(
    parametros_atuais,
    alteracoes,
    limite_inferior=-0.2,
    limite_superior=0.2,
):
    """Calcula o cenário original e outro com valores finais simulados."""
    atuais = _validar_parametros(parametros_atuais)
    _validar_alteracoes(alteracoes)
    simulados = dict(atuais)
    simulados.update(alteracoes)

    cenario_atual = _calcular(atuais, limite_inferior, limite_superior)
    cenario_simulado = _calcular(simulados, limite_inferior, limite_superior)
    alteracoes_aplicadas = {
        parametro: simulados[parametro]
        for parametro in _PARAMETROS
        if simulados[parametro] != atuais[parametro]
    }
    return {
        "cenario_atual": cenario_atual,
        "cenario_simulado": cenario_simulado,
        "parametros_atuais": dict(atuais),
        "parametros_simulados": dict(simulados),
        "alteracoes_aplicadas": alteracoes_aplicadas,
        "delta_lsi": cenario_simulado["lsi"] - cenario_atual["lsi"],
        "mudou_classificacao": (
            cenario_atual["classificacao"] != cenario_simulado["classificacao"]
        ),
        "resumo": _resumo(cenario_atual, cenario_simulado),
        "simulador_versao": "SIMULADOR_LSI_1.0",
    }


def _validar_variacoes(variacoes):
    if not isinstance(variacoes, dict):
        raise ValueError("variacoes deve ser um dicionário.")
    desconhecidas = set(variacoes) - set(_PARAMETROS)
    if desconhecidas:
        nomes = ", ".join(sorted(desconhecidas))
        raise ValueError(f"Chave(s) desconhecida(s) em variacoes: {nomes}.")
    for parametro, variacao in variacoes.items():
        if isinstance(variacao, bool) or not isinstance(variacao, (int, float)):
            raise ValueError(f"A variação de {parametro} deve ser um numero.")
        if not math.isfinite(variacao):
            raise ValueError(f"A variação de {parametro} deve ser um numero finito.")


def analisar_sensibilidade(
    parametros_atuais,
    variacoes=None,
    limite_inferior=-0.2,
    limite_superior=0.2,
):
    """Compara o efeito numérico de variações isoladas nos seis parâmetros."""
    atuais = _validar_parametros(parametros_atuais)
    variacoes_utilizadas = dict(_VARIACOES_PADRAO if variacoes is None else variacoes)
    _validar_variacoes(variacoes_utilizadas)
    if set(variacoes_utilizadas) != set(_PARAMETROS):
        ausentes = [nome for nome in _PARAMETROS if nome not in variacoes_utilizadas]
        variacoes_utilizadas.update(
            {nome: _VARIACOES_PADRAO[nome] for nome in ausentes}
        )

    resultado_atual = _calcular(atuais, limite_inferior, limite_superior)
    resultados = []
    for parametro in _PARAMETROS:
        variados = dict(atuais)
        variados[parametro] += variacoes_utilizadas[parametro]
        resultado_variado = _calcular(variados, limite_inferior, limite_superior)
        delta_lsi = resultado_variado["lsi"] - resultado_atual["lsi"]
        resultados.append(
            {
                "parametro": parametro,
                "valor_original": atuais[parametro],
                "valor_variado": variados[parametro],
                "lsi_variado": resultado_variado["lsi"],
                "delta_lsi": delta_lsi,
                "impacto_absoluto": abs(delta_lsi),
                "classificacao_variada": resultado_variado["classificacao"],
            }
        )

    resultados.sort(key=lambda item: (-item["impacto_absoluto"], item["parametro"]))
    ranking_influencia = [item["parametro"] for item in resultados]
    return {
        "lsi_atual": resultado_atual["lsi"],
        "classificacao_atual": resultado_atual["classificacao"],
        "variacoes_utilizadas": dict(variacoes_utilizadas),
        "resultados": resultados,
        "ranking_influencia": ranking_influencia,
        "principal_influencia": ranking_influencia[0],
        "aviso_metodologico": _AVISO_METODOLOGICO,
        "sensibilidade_versao": "SENSIBILIDADE_LSI_1.0",
    }