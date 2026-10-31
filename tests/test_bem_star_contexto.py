"""Impede que uma troca de cliente reaproveite análises e parecer de outro."""
import unittest
from test_fotos_isolamento import RAIZ
from utils.contexto_bem_star import preparar_contexto_bem_star, preencher_campo_agenda_bem_star


class TestContextoBemStar(unittest.TestCase):
    def setUp(self):
        self.ss = {"op_fotos_extras": "preservar", "rt_ph": "7.2"}
        preparar_contexto_bem_star(self.ss, "Tríad TESTE", "09", "2026", "C1")
        self.ss.update({"bs_agenda_ph_07092026": "7.5",
                        "bs_rel_parecer_tecnico": "Parecer Tríad TESTE",
                        "csr_obs_rel": "Observação Tríad TESTE",
                        "bs_rel_dias_semana": ["Segunda-feira"]})

    def test_troca_automatica_de_cliente_nao_reaproveita_dados(self):
        self.assertTrue(preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2"))
        for k in ["bs_agenda_ph_07092026", "bs_rel_parecer_tecnico", "csr_obs_rel"]:
            self.assertEqual(self.ss[k], "")
        self.assertEqual(self.ss["op_fotos_extras"], "preservar")
        self.assertEqual(self.ss["rt_ph"], "7.2")

    def test_retorno_ao_cliente_recupera_seu_rascunho(self):
        preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2")
        self.ss["bs_agenda_ph_07092026"] = "8.0"
        preparar_contexto_bem_star(self.ss, "Tríad TESTE", "09", "2026", "C1")
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "7.5")
        self.assertEqual(self.ss["bs_rel_parecer_tecnico"], "Parecer Tríad TESTE")

    def test_troca_de_mes_e_ano_separa_parecer(self):
        for mes, ano in [("10", "2026"), ("09", "2027")]:
            with self.subTest(mes=mes, ano=ano):
                preparar_contexto_bem_star(self.ss, "Tríad TESTE", mes, ano, "C1")
                self.assertEqual(self.ss["bs_rel_parecer_tecnico"], "")

    def test_rerun_mesmo_contexto_preserva_edicoes(self):
        self.assertFalse(preparar_contexto_bem_star(self.ss, "Tríad TESTE", "9", "2026", "C1"))
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "7.5")

    def test_mesmo_nome_com_outro_id_nao_herda_rascunho(self):
        preparar_contexto_bem_star(self.ss, "Tríad TESTE", "09", "2026", "C3")
        self.assertEqual(self.ss["bs_rel_parecer_tecnico"], "")

    def test_widgets_recebem_padrao_explicito_em_vez_de_recuperar_frontend_antigo(self):
        preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2")
        # Um widget com a mesma identidade usa o estado do navegador se não
        # houver valor atribuído no servidor. O padrão deve prevalecer.
        valor_widget = self.ss.get("csr_obs_rel", "Observação Tríad TESTE")
        self.assertEqual(valor_widget, "")
        self.assertEqual(self.ss["bs_rel_dias_semana"], [])
        self.assertEqual(self.ss["bs_rel_freq_semana"], 1)

    def test_agenda_do_novo_cliente_recebe_seus_parametros(self):
        preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2")
        preencher_campo_agenda_bem_star(self.ss, "bs_agenda_ph_07092026", "8.0")
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "8.0")
        preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2")
        self.ss["bs_agenda_ph_07092026"] = "7.2"
        preencher_campo_agenda_bem_star(self.ss, "bs_agenda_ph_07092026", "8.0")
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "7.2")

    def test_fonte_nao_sobrescreve_edicao_restaurada(self):
        preparar_contexto_bem_star(self.ss, "Gama TESTE", "09", "2026", "C2")
        preparar_contexto_bem_star(self.ss, "Tríad TESTE", "09", "2026", "C1")
        preencher_campo_agenda_bem_star(self.ss, "bs_agenda_ph_07092026", "8.0")
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "7.5")


if __name__ == "__main__":
    unittest.main()
