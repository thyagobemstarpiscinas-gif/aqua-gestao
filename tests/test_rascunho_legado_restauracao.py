"""Restaurar rascunho RT não grava nada até a primeira edição do usuário (garantia do PR #3).

Regressão do d469f90: limpar() gravava "" nos campos de escolha (status dos EPI,
NBR 11238, tipo de atendimento); na execução seguinte os widgets os corrigiam para o
padrão e o salvamento automático tomava a correção por edição, gravando sozinho uma
versão nova do rascunho logo após a restauração.

App real (AppTest) em modo produção com planilha FALSA em memória: nenhum acesso a
Google Sheets, Drive ou rede; dados fictícios.

Executar:  python tests/rodar_isolado.py --modules test_rascunho_legado_restauracao
"""

import json
import os
import sys
import unittest
from unittest import mock

from test_relatorio_rt_campos import AppTest, _AppReal
from utils.rt_contexto import limpar

# Padrões que os próprios widgets usariam (definidos aqui, independentes da implementação).
PADROES_ESCOLHA = {
    "rel_tipo_atendimento": "Contrato ativo",
    "rel_status_agua": "CONFORME",
    **{f"rel_epi_{e}_status": "Conforme" for e in ("luvas", "oculos", "respirador", "botas")},
    **{f"rel_nbr11238_{c}": "Sim" for c in ("profundidade", "retrolavagem", "skimmers", "circulacao", "chuveiro")},
}

LEGADO = {"salvo_em": "24/10/2026 18:00:00", "motivo": "autosave", "rel_analises_total": 12,
          "campos": {"rel_nome_condominio": "Condomínio Legado Fictício", "rel_representante": "Síndico Fictício",
                     "rel_diagnostico": "Diagnóstico antigo",
                     **{f"rel_analise_ph_{i}": f"7.{i}" for i in range(4)}}}


class TestLimparCamposDeEscolha(unittest.TestCase):
    def test_campos_de_escolha_voltam_ao_padrao_do_widget(self):
        ss = {"rel_epi_luvas_status": "Pendente", "rel_nbr11238_chuveiro": "Não", "rel_tipo_atendimento": "Inspeção técnica",
              "rel_status_agua": "NÃO CONFORME", "rel_diagnostico": "texto", "rel_analise_ph_0": "7,1"}
        limpar(ss)
        for k, v in PADROES_ESCOLHA.items():
            self.assertEqual(ss[k], v, k)
        self.assertEqual((ss["rel_diagnostico"], ss["rel_analise_ph_0"]), ("", ""))
        self.assertNotIn("", [ss[k] for k in PADROES_ESCOLHA])  # vazio não é opção válida


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestRestaurarRascunhoLegadoNaoGrava(_AppReal):
    def setUp(self):
        super().setUp()
        p = mock.patch.dict(os.environ, {"CODESPACES": ""})  # modo produção: rascunhos na planilha (falsa)
        p.start()
        self.addCleanup(p.stop)
        planilha = sys.modules["gspread"].PLANILHA
        self.aba = planilha.add_worksheet("_Rascunhos_RT", 100, 3)
        self.aba.linhas = [["Usuario", "Salvo em", "Dados JSON"],
                           ["thyago", LEGADO["salvo_em"], json.dumps(LEGADO, ensure_ascii=False)]]
        self.legado_original = list(self.aba.linhas[1])
        self.gravacoes = []
        for nome in ("append_row", "update", "batch_update"):
            original = getattr(self.aba, nome)

            def registrar(*a, _nome=nome, _original=original, **k):
                self.gravacoes.append(_nome)
                return _original(*a, **k)
            setattr(self.aba, nome, registrar)

    def _entrar(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": "pin-ficticio-teste"}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao("🔵 Aqua Gestão").click().run()
        at.text_input(key="pin_admin_input").input("pin-ficticio-teste")
        botao("Entrar no escritório").click().run()
        return at

    def test_restaurar_legado_nao_grava_e_primeira_edicao_vai_para_linha_propria(self):
        at = self._entrar()
        at.run()
        at.run()
        self.assertEqual(self.gravacoes, [])  # abrir o escritório não grava
        next(b for b in at.button if (b.key or "") == "btn_restaurar_rascunho_rt_thyago").click().run()
        at.run()
        at.run()
        ss = at.session_state
        self.assertEqual(ss["rel_nome_condominio"], "Condomínio Legado Fictício")
        self.assertEqual([ss[f"rel_analise_ph_{i}"] for i in range(4)], ["7.0", "7.1", "7.2", "7.3"])
        for k, v in PADROES_ESCOLHA.items():
            if k in ss:
                self.assertEqual(ss[k], v, k)
        self.assertEqual(self.gravacoes, [], "restaurar não pode gravar sem edição do usuário")
        self.assertFalse(at.exception, at.exception)

        at.text_area(key="rel_diagnostico").input("Diagnóstico editado 1").run()
        self.assertEqual(self.gravacoes, ["append_row"])  # primeira edição: linha nova desta sessão
        self.assertEqual(self.aba.linhas[1], self.legado_original)  # linha legada intacta
        at.text_area(key="rel_diagnostico").input("Diagnóstico editado 2").run()
        self.assertEqual(self.gravacoes[0], "append_row")
        self.assertNotIn("append_row", self.gravacoes[1:])  # depois, só atualiza a própria linha
        self.assertEqual(self.aba.linhas[1], self.legado_original)


if __name__ == "__main__":
    unittest.main()
