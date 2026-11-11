"""ORP: regra única nos relatórios do escritório e de campo (dados fictícios, sem rede).

Critério (utils.normas_piscinas): ABNT NBR 10818:2026, seção 6 / Tabela 3 — "ORP mínimo
650 mV quando aplicável" (docs/RT_NORMAS_20261101.md). O material aprovado não atribui
não conformidade à leitura abaixo da referência: abaixo de 650 mV é ponto de atenção,
650 mV ou mais atende, e leitura ausente não vira zero, conformidade nem leitura avaliável.

Geradores cobertos: escritório (avaliar_conformidade_analises + gerar_textos_automaticos_relatorio
+ interpretacao_resultados), campo/operador (build_monthly_rt_pdf_bytes do app.py) e o
módulo rt_monthly_report.
"""
import unittest

import fitz

from test_mensal_rt_app_pdf import _carregar
from utils.normas_piscinas import (
    ORP_ABAIXO_REFERENCIA, ORP_ATENDE_REFERENCIA, ORP_AUSENTE, RECOMENDACAO_ORP,
    avaliar_orp, ler_orp,
)

# Demais parâmetros dentro dos limites da NBR 10818:2026.
BASE = {"data": "10/11/2026", "piscina": "Piscina Adulto", "ph": "7,4", "cloro_livre": "2,0",
        "cloro_total": "2,3", "alcalinidade": "100", "dureza": "200", "cianurico": "40", "turbidez": "0,3"}
CASOS = {  # leitura → estado esperado
    "649": ORP_ABAIXO_REFERENCIA, "649,9": ORP_ABAIXO_REFERENCIA,
    "650": ORP_ATENDE_REFERENCIA, "650,0": ORP_ATENDE_REFERENCIA,
    "651": ORP_ATENDE_REFERENCIA, "780": ORP_ATENDE_REFERENCIA,
    "": ORP_AUSENTE,
}


def _com_orp(orp, **extra):
    return {**BASE, "orp": orp, **extra}


class RegraCompartilhada(unittest.TestCase):
    def test_abaixo_no_limite_acima_e_ausente(self):
        for valor, esperado in CASOS.items():
            self.assertEqual(avaliar_orp(valor), esperado, valor)
        self.assertEqual(avaliar_orp(649.99), ORP_ABAIXO_REFERENCIA)
        self.assertEqual(avaliar_orp(650), ORP_ATENDE_REFERENCIA)
        self.assertEqual(avaliar_orp("1.234,5"), ORP_ATENDE_REFERENCIA)

    def test_ausencia_nunca_vira_zero(self):
        for vazio in (None, "", "  ", "—", "-", "None", "sem leitura", float("nan")):
            self.assertIsNone(ler_orp(vazio), repr(vazio))
            self.assertEqual(avaliar_orp(vazio), ORP_AUSENTE, repr(vazio))
        self.assertEqual(ler_orp("0"), 0.0)  # zero digitado é leitura real (abaixo da referência)
        self.assertEqual(avaliar_orp("0"), ORP_ABAIXO_REFERENCIA)


class Escritorio(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _carregar()

    def _avaliar(self, orp, **extra):
        analises = [_com_orp(orp, **extra)]
        aval = self.ns["avaliar_conformidade_analises"](analises)
        return aval, self.ns["gerar_textos_automaticos_relatorio"](analises, aval)

    def test_abaixo_e_atencao_sem_mudar_status(self):
        aval, textos = self._avaliar("640")
        self.assertEqual(aval["status"], "CONFORME")
        self.assertEqual(aval["detalhes"], [])
        self.assertEqual(len(aval["atencoes"]), 1)
        self.assertIn("ORP=640 mV abaixo da referência mín. 650 mV (NBR 10818:2026, quando aplicável) — ponto de atenção",
                      aval["atencoes"][0])
        self.assertIn(RECOMENDACAO_ORP, textos["recomendacoes"])

    def test_limite_e_acima_sem_atencao(self):
        for orp in ("650", "651", "780"):
            aval, textos = self._avaliar(orp)
            self.assertEqual((aval["status"], aval["atencoes"]), ("CONFORME", []), orp)
            self.assertNotIn(RECOMENDACAO_ORP, textos["recomendacoes"], orp)

    def test_ausente_nao_avalia_nem_confere(self):
        aval, _ = self._avaliar("")
        self.assertEqual((aval["status"], aval["atencoes"]), ("CONFORME", []))  # pelos demais parâmetros
        self.assertNotIn("orp", aval["informativos"])
        so_orp_vazio = self.ns["avaliar_conformidade_analises"]([{"data": "10/11/2026", "orp": ""}])
        self.assertEqual(so_orp_vazio["status"], "NÃO AVALIADO")
        so_orp = self.ns["avaliar_conformidade_analises"]([{"data": "10/11/2026", "orp": "700"}])
        self.assertEqual(so_orp["status"], "NÃO AVALIADO")  # ORP sozinho não torna o período avaliável

    def test_outros_parametros_continuam_classificados(self):
        aval, textos = self._avaliar("640", ph="6,8", turbidez="0,8")
        self.assertEqual(aval["status"], "NÃO CONFORME")
        detalhes = " | ".join(aval["detalhes"])
        self.assertIn("pH=6,8", detalhes)
        self.assertIn("turbidez=0,8", detalhes)
        self.assertNotIn("ORP", detalhes)
        self.assertEqual(len(aval["atencoes"]), 1)

    def test_interpretacao_do_pdf(self):
        from utils.rt_relatorio import interpretacao_resultados
        bloco = dict(interpretacao_resultados([_com_orp("640"), _com_orp("650", data="17/11/2026")]))
        texto = bloco["Turbidez, ORP, TDS e temperatura"]
        self.assertIn("abaixo da referência mín. 650 mV (quando aplicável, ponto de atenção) em: 10/11/2026", texto)
        self.assertNotIn("17/11/2026", texto)  # 650 mV atende
        bloco = dict(interpretacao_resultados([_com_orp("650")]))
        self.assertIn("sem leitura abaixo da referência", bloco["Turbidez, ORP, TDS e temperatura"])


class CampoOperador(unittest.TestCase):
    """build_monthly_rt_pdf_bytes do app.py (relatório RT da área do operador)."""

    @classmethod
    def setUpClass(cls):
        cls.ns = _carregar()

    def _texto(self, registros):
        pdf = self.ns["build_monthly_rt_pdf_bytes"](registros, "CONDOMINIO FICTICIO ORP", "11", "2026", "30/11/2026")
        with fitz.open(stream=pdf, filetype="pdf") as doc:
            return " ".join(" ".join(p.get_text() for p in doc).split())

    def test_abaixo_e_atencao_nao_nao_conformidade(self):
        texto = self._texto([_com_orp("640")])
        self.assertIn("AVALIAÇÃO PARCIAL — SEM DESVIOS OPERACIONAIS IDENTIFICADOS", texto)
        self.assertNotIn("HÁ DESVIOS", texto)
        self.assertNotIn("verificar: ORP", texto)
        self.assertIn("Não conformidades identificadas: Não foram identificadas não conformidades relevantes", texto)
        self.assertIn("ORP=640 mV abaixo da referência mín. 650 mV (NBR 10818:2026, quando aplicável) — ponto de atenção", texto)
        self.assertIn(RECOMENDACAO_ORP, texto)

    def test_limite_e_acima_sem_atencao(self):
        for orp in ("650", "651"):
            texto = self._texto([_com_orp(orp)])
            self.assertIn("AVALIAÇÃO PARCIAL", texto, orp)
            self.assertNotIn("ponto de atenção", texto, orp)
            self.assertNotIn(RECOMENDACAO_ORP, texto, orp)

    def test_ausente_nao_vira_zero_nem_avaliacao(self):
        texto = self._texto([_com_orp("")])
        self.assertIn("AVALIAÇÃO PARCIAL", texto)
        self.assertNotIn("ponto de atenção", texto)
        self.assertNotIn("ORP=0", texto)
        so_orp = self._texto([{"data": "10/11/2026", "orp": "700"}])
        self.assertIn("Não avaliado", so_orp)  # ORP sozinho: sem leitura avaliável

    def test_outros_parametros_continuam_desvios(self):
        texto = self._texto([_com_orp("640", ph="6,8", turbidez="0,8")])
        self.assertIn("HÁ DESVIOS", texto)
        self.assertIn("verificar: pH, turbidez (limites da NBR 10818:2026)", texto)
        self.assertNotIn("ORP (limites", texto)
        self.assertIn("ponto de atenção", texto)


class ModuloRtMonthlyReport(unittest.TestCase):
    def test_mesma_regra(self):
        from rt_monthly_report import (_collect_deviation_data, _collect_summary,
                                       _recommendations_from_deviations, _status_for_value)
        linha = {"data": "10/11/2026", "ph": "7,4", "crl": "2,0", "ct": "2,3", "orp": "640"}
        self.assertEqual(_status_for_value("orp", "640"), "attention")
        self.assertEqual(_status_for_value("orp", "650"), "ok")
        self.assertEqual(_status_for_value("orp", ""), "missing")
        self.assertEqual(_collect_deviation_data([linha]), {})
        resumo = _collect_summary([linha])
        self.assertEqual(resumo["status_flag"], "ok")
        self.assertEqual(len(resumo["attention"]), 1)
        self.assertIn(RECOMENDACAO_ORP, _recommendations_from_deviations([linha]))
        self.assertEqual(_collect_summary([{"data": "10/11/2026", "orp": "700"}])["status_flag"], "info")
        # Outros parâmetros inalterados.
        self.assertEqual(_collect_deviation_data([{**linha, "crl": "0,8", "ct": "1,0"}]), {"10/11/2026": ["CRL"]})


if __name__ == "__main__":
    unittest.main()
