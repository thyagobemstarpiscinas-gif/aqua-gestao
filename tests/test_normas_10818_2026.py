from utils.normas_piscinas import avaliar_leitura, cloro_combinado
from rt_monthly_report import _collect_summary, prepare_monitoring_table_rows


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
