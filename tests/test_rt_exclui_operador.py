"""Relatório mensal RT: só visitas técnicas RT entram; visitas de operador ficam fora,
inclusive numa terça-feira.

Achado da homologação do PR #16: no escritório, o filtro só considerava "terça-feira +
uma por semana", e a V00002 (operador fictício, terça 06/10/2026) entrou no mensal.
Planilha falsa em memória, sem rede.
"""
import json
import sys
import unittest

import test_rt_relatorio_apptest as _base
from rt_monthly_report import exclude_operator_records

AppTest = _base.AppTest
ALFA = "Condomínio Alfa Fictício"


def _payload(data, ph, crl, rt):
    """Mesmo formato que o modo campo grava (app.py, salvar lançamento)."""
    return {"data": data, "condominio": ALFA, "status": "Concluída",
            "operador": "Responsável Técnico Fictício" if rt else "João",
            "visita_rt_semanal": rt, "tipo_visita": "RT semanal" if rt else "Operador",
            "piscinas": [{"nome": "Piscina Adulto", "ph": ph, "cloro_livre": crl}]}


def _linha(id_visita, data, ph, crl, rt=None):
    """Linha da aba 🔬 Visitas; rt=None é linha antiga, sem Payload JSON nem marca."""
    operador = "João" if not rt else "Responsável Técnico Fictício"
    payload = "" if rt is None else json.dumps(_payload(data, ph, crl, rt), ensure_ascii=False)
    return ["", id_visita, data, "C1", ALFA, ph, crl, "", "", "", ""] + [""] * 8 + \
        ["Concluída", operador, "", payload]


class TestFiltroVisitasRT(unittest.TestCase):
    def test_operador_em_terca_e_linha_sem_marca_ficam_fora(self):
        terca_operador = _payload("06/10/2026", "7,1", "0,8", rt=False)
        terca_rt = _payload("13/10/2026", "7,5", "2,0", rt=True)
        sem_marca = {"data": "20/10/2026", "condominio": ALFA, "ph": "7,3", "operador": "João"}
        rt_manual = {"data": "27/10/2026", "tipo_visita": "Relatório mensal RT", "is_rt": True}
        self.assertEqual(exclude_operator_records([terca_operador, terca_rt, sem_marca, rt_manual]),
                         [terca_rt, rt_manual])


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestMensalRTExcluiOperadorNoApp(unittest.TestCase):
    setUp = _base.TestRelatorioRTProfissional.setUp
    _entrar_e_carregar = _base.TestRelatorioRTProfissional._entrar_e_carregar
    _mes = _base.TestRelatorioRTProfissional._mes
    _carregar = _base.TestRelatorioRTProfissional._carregar
    _preparar_alfa = _base.TestRelatorioRTProfissional._preparar_alfa

    def test_visita_de_operador_em_terca_nao_entra_no_mensal_rt(self):
        import streamlit as st_real
        self._preparar_alfa()
        vis = sys.modules["gspread"].PLANILHA.worksheet("🔬 Visitas")
        vis.linhas = vis.linhas[:1] + [
            _linha("V1", "06/10/2026", "7,1", "0,8", rt=False),  # terça, operador
            _linha("V2", "13/10/2026", "7,5", "2,0", rt=True),   # terça, RT
            _linha("V3", "20/10/2026", "7,3", "1,1"),            # terça, linha antiga sem marca
            _linha("V4", "27/10/2026", "6,9", "0,5", rt=False),  # terça, operador antes do RT na mesma semana
            _linha("V5", "27/10/2026", "7,6", "2,4", rt=True),   # terça, RT: não perde a vaga da semana
        ]
        st_real.cache_data.clear()
        at = self._entrar_e_carregar()
        self._mes(at, "10")
        ss = at.session_state
        importadas = [(ss[f"rel_analise_data_{i}"], ss[f"rel_analise_ph_{i}"], ss[f"rel_analise_cl_{i}"])
                      for i in range(4)]
        self.assertEqual(sorted(importadas[:2]), [("13/10/2026", "7,5", "2,0"), ("27/10/2026", "7,6", "2,4")])
        self.assertEqual(importadas[2:], [("", "", ""), ("", "", "")])
        self.assertTrue(any("3 visita(s) do período ficaram fora do relatório RT" in i.value for i in at.info),
                        [i.value for i in at.info])
        # Nada é apagado da planilha.
        self.assertEqual([r[1] for r in vis.linhas[1:]], ["V1", "V2", "V3", "V4", "V5"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
