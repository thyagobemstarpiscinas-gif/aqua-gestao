"""Impede que uma troca de cliente reaproveite análises e parecer de outro."""
import unittest
from test_fotos_isolamento import RAIZ
from utils.contexto_bem_star import preparar_contexto_bem_star


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
            self.assertNotIn(k, self.ss)
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
                self.assertNotIn("bs_rel_parecer_tecnico", self.ss)

    def test_rerun_mesmo_contexto_preserva_edicoes(self):
        self.assertFalse(preparar_contexto_bem_star(self.ss, "Tríad TESTE", "9", "2026", "C1"))
        self.assertEqual(self.ss["bs_agenda_ph_07092026"], "7.5")

    def test_mesmo_nome_com_outro_id_nao_herda_rascunho(self):
        preparar_contexto_bem_star(self.ss, "Tríad TESTE", "09", "2026", "C3")
        self.assertNotIn("bs_rel_parecer_tecnico", self.ss)


if __name__ == "__main__":
    unittest.main()
