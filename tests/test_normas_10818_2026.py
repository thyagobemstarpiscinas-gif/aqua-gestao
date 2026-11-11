from utils.normas_piscinas import avaliar_leitura, cloro_combinado
from rt_monthly_report import (
    _collect_summary, _recommendations_from_deviations,
    prepare_monitoring_table_rows,
)


def test_limites_e_metas_sao_distintos():
    assert avaliar_leitura("ph", 7.0) == "fora_ideal"
    assert avaliar_leitura("ph", 7.5) == "dentro_ideal"
    assert avaliar_leitura("ph", 7.9) == "fora_limite"
    assert avaliar_leitura("crl", 0.7) == "fora_limite"
    assert avaliar_leitura("crl", 0.7, tratamento_complementar=True) == "fora_ideal"
    assert avaliar_leitura("turbidez", 0.5) == "fora_limite"
    assert avaliar_leitura("cc", 0.6) == "dentro_ideal"
    assert avaliar_leitura("cc", 0.61) == "fora_limite"
    assert avaliar_leitura("tds", 1600) == "sem_avaliacao"


def test_ct_inferior_ao_livre_nao_vira_zero():
    assert cloro_combinado(2.5, 2.0)[0] is None
    row = prepare_monitoring_table_rows([{"data": "01/11/2026", "crl": 2.5, "ct": 2.0}])[0]
    assert row[4] == {"text": "Inconsistente", "status": "out"}
    assert prepare_monitoring_table_rows([{"crl": 1.0, "ct": 1.7}])[0][4]["status"] == "out"


def test_ausencia_de_medicao_nao_e_condicao_estavel():
    resumo = _collect_summary([{"data": "01/11/2026"}])
    assert resumo["status_flag"] == "info"
    assert "Sem medições" in resumo["status"]


def test_ct_nao_tem_limite_independente_e_cc_gera_recomendacao():
    row = {"data": "01/11/2026", "crl": 3.0, "ct": 3.7, "turbidez": 0.5}
    tabela = prepare_monitoring_table_rows([row])[0]
    assert tabela[3]["status"] == "ok"  # CT isolado não é não conformidade
    assert tabela[4]["status"] == "out"  # CC = 0,7 ppm
    recs = " ".join(_recommendations_from_deviations([row]))
    assert "cloro combinado" in recs
    assert "turbidez" in recs


def test_relatorio_mensal_usa_limites_2026_e_nao_metas_antigas():
    from rt_monthly_report import _status_for_value, _collect_deviation_data
    # Dentro dos limites 2026, mas fora das faixas operacionais antigas.
    for campo, valor in [("ph", 7.1), ("crl", 4.0), ("alcalinidade", 70), ("dureza", 120), ("cya", 80)]:
        assert _status_for_value(campo, valor) == "ok", (campo, valor)
    # Fora dos limites 2026, mas dentro das faixas antigas (CRL 0,5–3,0).
    assert _status_for_value("crl", 0.8) == "out"
    assert _status_for_value("orp", 600) == "attention"  # referência "quando aplicável": ponto de atenção
    assert _status_for_value("tds", 5000) == "ok"  # TDS não é classificado isoladamente
    desvios = _collect_deviation_data([{"data": "02/11/2026", "crl": 0.8, "ph": 7.1}])
    assert desvios == {"02/11/2026": ["CRL"]}


def test_textos_de_limite_usam_apenas_caracteres_winansi():
    from utils.normas_piscinas import texto_limite, texto_ideal, texto_resumo_limites
    assert texto_limite("ph") == "7,0–7,8" and texto_ideal("ph") == "7,4–7,6"
    assert texto_limite("crl") == "1,0–5,0"
    assert texto_limite("crl", tratamento_complementar=True) == "0,5–5,0"
    assert texto_limite("cya") == "até 100"
    assert texto_limite("cc") == "até 0,6"
    assert texto_limite("turbidez") == "< 0,5"
    assert texto_limite("orp") == "mín. 650"
    assert texto_limite("lsi") == "-0,3 a +0,5"
    texto = texto_resumo_limites()
    texto.encode("cp1252")  # renderizável com Helvetica (WinAnsi)
    assert "NBR 10818:2026" in texto


def test_consolidado_nao_classifica_cloro_total_e_usa_limites_2026():
    from utils.rt_relatorio import consolidar, PARAMETROS
    faixas = {k: (mn, mx) for k, _r, mn, mx, _u in PARAMETROS}
    assert faixas["ph"] == (7.0, 7.8) and faixas["cloro_livre"] == (1.0, 5.0)
    assert faixas["cloro_total"] == (None, None)
    cons = consolidar([
        {"data": "03/11/2026", "piscina": "Adulto", "ph": "7,1", "cloro_livre": "0,8", "cloro_total": "9"},
        {"data": "10/11/2026", "piscina": "Adulto", "ph": "7,5", "cloro_livre": "2,5", "cloro_total": "2,6"},
    ], [(11, 2026)])
    params = cons["piscinas"]["Adulto"]["parametros"]
    assert params["ph"]["conformidade"] == 100.0
    assert params["cloro_livre"]["conformidade"] == 50.0
    assert params["cloro_total"]["conformidade"] is None


# O executor oficial (tests/rodar_isolado.py) usa unittest, que não coleta
# funções soltas: expõe os mesmos testes numa TestCase.
import unittest as _unittest


class Normas10818(_unittest.TestCase):
    pass


for _nome, _fn in list(globals().items()):
    if _nome.startswith("test_") and callable(_fn):
        setattr(Normas10818, _nome, (lambda f: lambda self: f())(_fn))

del _nome, _fn
