"""Estimativas de correção da água por piscina, com o teor do produto indicado."""
from math import isfinite

from utils.normas_piscinas import FAIXAS_COLETIVA_SIMPLES as FAIXAS

CLOROS = {"Hipoclorito de cálcio 65%": 0.65, "Dicloro 56%": 0.56}
META_CRL = 3.0
META_ALC = 100.0
META_DC = 225.0

# 1 ppm em 1 m³ corresponde a 1 g. AT e DC são expressas como CaCO3.
G_NAHCO3_POR_G_CACO3 = 84.01 / (100.09 / 2)
G_CACL2_POR_G_CACO3 = 110.98 / 100.09


def _orientacao(prioridade, produto, acao, motivo):
    return dict(prioridade=prioridade, produto=produto, acao=acao, justificativa=motivo,
                quantidade=0, unidade="")


def _massa(prioridade, produto, acao, massa_ativa_g, teor, memoria):
    if teor is None or not isfinite(teor) or not 0 < teor <= 100:
        return _orientacao(prioridade, produto, acao, "Informe o teor percentual indicado no rótulo para calcular a massa.")
    gramas = massa_ativa_g * 100 / teor
    gramas = round(gramas) if gramas >= 1 else round(gramas, 2)
    return dict(prioridade=prioridade, produto=produto, acao=acao,
                justificativa=f"{memoria} ÷ ({teor:g}/100) = {gramas:g} g de produto.",
                quantidade=gramas / 1000 if gramas >= 1000 else gramas,
                unidade="kg" if gramas >= 1000 else "g")


def calcular_sugestoes_dosagem(ph, crl, alc, dc, cya, volume_m3, *, produto_cloro="Hipoclorito de cálcio 65%",
                                teor_bicarbonato=None, teor_cloreto=None):
    """O volume é individual. Massas são teóricas, sem fator de demanda presumido."""
    if volume_m3 is None or not isfinite(volume_m3) or volume_m3 <= 0:
        return []
    if any(v is not None and (not isfinite(v) or v < 0) for v in (ph, crl, alc, dc, cya)) or (ph is not None and ph > 14):
        return [_orientacao(1, "Conferir medições", "Revise os valores da coleta.", "Há uma leitura inválida para o cálculo.")]

    sugestoes = []
    ph_alto = ph is not None and ph > FAIXAS["ph"].maximo
    alc_alta = alc is not None and alc > FAIXAS["alcalinidade"].ideal_max
    if ph_alto or alc_alta:
        sugestoes.append(_orientacao(1, "Avaliar redução de pH ou alcalinidade",
            "Conferir pH e alcalinidade antes de dosar ácido muriático 31%.",
            "Para reduzir pH, use o teste de demanda ácida e a tabela do kit para o teor do rótulo. "
            "Para alcalinidade alta, o RT deve definir as etapas e acompanhar o pH. A mesma aplicação afeta ambos."))
    if ph is not None and ph < FAIXAS["ph"].minimo:
        sugestoes.append(_orientacao(1, "Ajuste de pH", "Conferir alcalinidade e demanda de base.",
            "Defina o corretor e a quantidade conforme o rótulo e a medição."))

    if crl is not None and crl < META_CRL:
        delta = META_CRL - crl
        acao = f"Elevar cloro livre de {crl:g} para {META_CRL:g} ppm (meta operacional)"
        if produto_cloro not in CLOROS:
            sugestoes.append(_orientacao(2, "Cloro", acao, "Selecione o produto e o teor de cloro disponível."))
        elif produto_cloro == "Dicloro 56%" and cya is not None and cya >= FAIXAS["cya"].maximo:
            sugestoes.append(_orientacao(2, produto_cloro, "Selecionar cloro sem estabilizante.",
                "O dicloro também acrescenta ácido cianúrico e o CYA já atingiu o máximo."))
        else:
            sugestoes.append(_massa(2, produto_cloro, acao, delta * volume_m3, CLOROS[produto_cloro] * 100,
                                     f"{volume_m3:g} m³ × {delta:g} ppm"))
    elif crl is not None and crl > FAIXAS["crl"].maximo:
        sugestoes.append(_orientacao(2, "Cloro livre elevado", "Suspender a adição de cloro e medir novamente.",
            "Avaliar o residual com o RT antes da liberação da piscina."))

    if alc is not None and alc < FAIXAS["alcalinidade"].ideal_min:
        delta = META_ALC - alc
        sugestoes.append(_massa(3, "Bicarbonato de sódio", f"Elevar alcalinidade de {alc:g} para {META_ALC:g} ppm",
                                volume_m3 * delta * G_NAHCO3_POR_G_CACO3, teor_bicarbonato,
                                f"{volume_m3:g} m³ × {delta:g} ppm como CaCO₃ × {G_NAHCO3_POR_G_CACO3:.4f}"))
    if dc is not None and dc < FAIXAS["dureza"].ideal_min:
        delta = META_DC - dc
        sugestoes.append(_massa(4, "Cloreto de cálcio", f"Elevar dureza cálcica de {dc:g} para {META_DC:g} ppm",
                                volume_m3 * delta * G_CACL2_POR_G_CACO3, teor_cloreto,
                                f"{volume_m3:g} m³ × {delta:g} ppm como CaCO₃ × {G_CACL2_POR_G_CACO3:.4f}"))
    if dc is not None and dc > FAIXAS["dureza"].maximo:
        sugestoes.append(_orientacao(4, "Renovação parcial de água", f"Dureza cálcica de {dc:g} ppm.",
            "Meça a dureza da água de reposição e defina a meta para calcular o volume a renovar."))
    if cya is not None and cya > FAIXAS["cya"].maximo:
        sugestoes.append(_orientacao(5, "Renovação parcial de água", f"CYA de {cya:g} ppm, acima do limite.",
            "Verifique o CYA da água de reposição e defina a meta para calcular o volume a renovar."))
    elif cya is not None and (cya > FAIXAS["cya"].ideal_max or 0 < cya < FAIXAS["cya"].ideal_min):
        sugestoes.append(_orientacao(5, "Monitorar CYA", f"CYA de {cya:g} ppm.",
            "Para água estabilizada, a faixa ideal é 30–50 ppm e o máximo é 100 ppm."))
    return sugestoes
