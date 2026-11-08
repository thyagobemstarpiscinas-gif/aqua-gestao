"""Sincronização do rascunho RT no app real (AppTest, planilha falsa, sem rede).

Sequência da homologação: salvar, falha transitória da planilha, salvar de novo,
restaurar explicitamente a cópia local, planilha volta. Antes, o aviso de
pendência ("Há cópias locais pendentes; versões divergentes…") nunca saía e o
salvamento manual dizia "Não foi possível confirmar o salvamento".
"""
import os
import re
import sys
import unittest
from unittest.mock import patch

from test_relatorio_rt_campos import AppTest, _AppReal, TestRelatorioRTCamposNoAppReal as _Aux


def _aviso_pendencia(at):
    return [e.value for e in at.error if "AINDA NÃO foram salvas na planilha" in e.value]


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class SincronizacaoRascunhoRT(_AppReal):
    _entrar_e_carregar = _Aux._entrar_e_carregar

    def setUp(self):
        super().setUp()
        p = patch.dict(os.environ, {"CODESPACES": ""})  # rascunho na planilha (falsa), como no Cloud
        p.start()
        self.addCleanup(p.stop)

    def _linhas(self):
        aba = sys.modules["gspread"].PLANILHA.worksheet("_Rascunhos_RT")
        return aba, [r for r in aba.linhas[1:] if r]

    def test_falha_transitoria_restauracao_e_confirmacao_limpam_o_aviso(self):
        at = self._entrar_e_carregar("Condomínio Alfa Fictício")
        at.text_area(key="rel_diagnostico").input("VERSAO_1_SALVA").run()
        self.assertFalse(_aviso_pendencia(at))
        aba, _ = self._linhas()
        falha = ConnectionError("COTA_FICTICIA")
        with patch.object(type(aba), "append_row", side_effect=falha), \
                patch.object(type(aba), "batch_update", side_effect=falha):
            at.text_area(key="rel_diagnostico").input("VERSAO_2_DURANTE_FALHA").run()
            at.run()  # o painel fica acima do autosave: o aviso aparece na execução seguinte
            self.assertTrue(_aviso_pendencia(at))
            next(b for b in at.button if b.label == "💾 Salvar rascunho").click().run()
            self.assertFalse([e.value for e in at.error if "Não foi possível confirmar o salvamento" in e.value])
            self.assertTrue([w.value for w in at.warning if "envio à planilha ficou pendente" in w.value])
            # O painel não trata a falha de conexão como versão divergente.
            self.assertFalse([w.value for w in at.warning if "Cópias divergentes" in w.value])
            # Restauração explícita da cópia local pendente.
            botao = next(b for b in at.button if (b.key or "").startswith("btn_restaurar_rascunho_rt_")
                         and "_local_" in (b.key or ""))
            botao.click().run()
            self.assertFalse(at.exception, at.exception)
            self.assertEqual(at.session_state["rel_diagnostico"], "VERSAO_2_DURANTE_FALHA")
            self.assertTrue(_aviso_pendencia(at))  # ainda sem confirmação da planilha
        # Planilha de volta: o próximo rerun reenvia e confirma por releitura.
        at.session_state["_rt_rasc_proxima_tentativa"] = 0
        at.run()
        next(b for b in at.button if b.label == "💾 Salvar rascunho").click().run()
        self.assertTrue([s.value for s in at.success if "preservado na planilha" in s.value])
        at.session_state["_rt_rasc_proxima_tentativa"] = 0
        at.run()
        self.assertFalse(_aviso_pendencia(at), _aviso_pendencia(at))
        self.assertNotIn("_rt_rasc_sync_pendente", at.session_state)
        _, linhas = self._linhas()
        self.assertTrue(any("VERSAO_2_DURANTE_FALHA" in "".join(r) for r in linhas))
        # Nada de outro condomínio ou mês misturado.
        self.assertEqual(at.session_state["rel_nome_condominio"], "Condomínio Alfa Fictício")
        self.assertEqual(at.session_state["rel_origem_dados"].rsplit("|", 1)[0], "aqua_gestao|id:C1")

    def test_conflito_real_mantem_aviso_e_nao_sobrescreve(self):
        at = self._entrar_e_carregar("Condomínio Alfa Fictício")
        at.text_area(key="rel_diagnostico").input("VERSAO_1_SALVA").run()
        chave = at.session_state["_rt_rasc_chave"]
        aba, _ = self._linhas()
        falha = ConnectionError("COTA_FICTICIA")
        with patch.object(type(aba), "append_row", side_effect=falha), \
                patch.object(type(aba), "batch_update", side_effect=falha):
            at.text_area(key="rel_diagnostico").input("VERSAO_2_LOCAL").run()
        # Outro gravador altera a linha remota desta chave (conflito real).
        indice = next(i for i, r in enumerate(aba.linhas) if r and r[0] == chave)
        import json
        remoto = json.loads(aba.linhas[indice][2])
        remoto["campos"]["rel_diagnostico"] = "ALTERACAO_DE_OUTRO_ACESSO"
        aba.linhas[indice][2] = json.dumps(remoto, ensure_ascii=False)
        at.session_state["_rt_rasc_proxima_tentativa"] = 0
        at.run()
        self.assertTrue(_aviso_pendencia(at))
        self.assertIn("ALTERACAO_DE_OUTRO_ACESSO", aba.linhas[indice][2])  # não sobrescrito
        self.assertTrue([w.value for w in at.warning if "Cópias divergentes" in w.value])
        self.assertFalse(at.exception, at.exception)


del _Aux
