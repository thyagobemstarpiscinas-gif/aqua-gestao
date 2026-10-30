"""Painel real, duas variantes e reabertura; executor bloqueia rede e Drive."""
import copy
import os
import re
import sys
import unittest
from unittest.mock import patch

from test_relatorio_rt_campos import AppTest, _AppReal, TestRelatorioRTCamposNoAppReal as _Aux


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestVariantesNoPainel(_AppReal):
    _entrar_e_carregar = _Aux._entrar_e_carregar

    def test_reabertura_e_restauracao_explicita_de_ambas_variantes(self):
        p = patch.dict(os.environ, {"CODESPACES": ""})
        p.start()
        self.addCleanup(p.stop)
        at = self._entrar_e_carregar()
        at.text_area(key="rel_diagnostico").input("VERSAO_REMOTA_FICTICIA").run()
        self.assertFalse(at.exception, at.exception)
        chave = at.session_state["_rt_rasc_chave"]
        aba = sys.modules["gspread"].PLANILHA.worksheet("_Rascunhos_RT")
        remota = copy.deepcopy([row for row in aba.linhas if row and row[0] == chave])
        self.assertEqual(len(remota), 1)
        with patch.object(aba, "batch_update", side_effect=ConnectionError("FALHA_FICTICIA")):
            at.text_area(key="rel_diagnostico").input("LOCAL_PENDENTE_FICTICIA").run()
        self.assertFalse(at.exception, at.exception)
        self.assertIn(chave, at.session_state["_rt_rasc_sync_pendente"])
        at.session_state["_rt_rasc_proxima_tentativa"] = 0
        at.run()
        self.assertEqual([row for row in aba.linhas if row and row[0] == chave], remota)
        prefixo = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", chave)
        self.assertTrue(any((b.key or "").startswith(prefixo + "_local_") for b in at.button))
        self.assertTrue(any((b.key or "").startswith(prefixo + "_planilha_") for b in at.button))
        # Cada restauração parte de um novo AppTest: a pendência vem do disco,
        # não do session_state anterior. Nenhuma restauração é automática.
        for origem, esperado in (("local", "LOCAL_PENDENTE_FICTICIA"), ("planilha", "VERSAO_REMOTA_FICTICIA")):
            with self.subTest(origem=origem):
                novo = self._entrar_e_carregar()
                self.assertIn(chave, novo.session_state["_rt_rasc_sync_pendente"])
                self.assertNotEqual(novo.session_state["rel_diagnostico"], "LOCAL_PENDENTE_FICTICIA")
                botao = next(b for b in novo.button if (b.key or "").startswith(prefixo + "_" + origem + "_"))
                botao.click().run()
                self.assertFalse(novo.exception, novo.exception)
                self.assertEqual(novo.session_state["rel_diagnostico"], esperado)
                self.assertNotEqual(novo.session_state["_rt_rasc_chave"], chave)
                novo.run()
                self.assertEqual(novo.session_state["rel_diagnostico"], esperado)
                self.assertEqual([row for row in aba.linhas if row and row[0] == chave], remota)


del _Aux
