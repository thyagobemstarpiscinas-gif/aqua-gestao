"""Trocas reais de widgets Streamlit, sem Sheets ou dados de clientes reais."""
import unittest
from test_fotos_isolamento import RAIZ
from streamlit.testing.v1 import AppTest


TELA = '''
import streamlit as st
from utils.contexto_bem_star import preparar_contexto_bem_star, preencher_campo_agenda_bem_star
cliente = st.selectbox("Cliente", ["Tríad TESTE", "Gama TESTE"], key="csr_sel_relatorio")
mes = st.selectbox("Mês", ["09", "10"], key="csr_mes_rel")
preparar_contexto_bem_star(st.session_state, cliente, mes, "2026", cliente)
st.text_area("Observação", key="csr_obs_rel")
st.text_area("Parecer", key="bs_rel_parecer_tecnico")
st.number_input("Frequência", min_value=1, max_value=7, key="bs_rel_freq_semana")
st.multiselect("Dias", ["Segunda", "Quarta", "Sexta"], key="bs_rel_dias_semana")
preencher_campo_agenda_bem_star(st.session_state, "bs_agenda_ph_07092026",
    "7.5" if cliente == "Tríad TESTE" else "8.0")
st.text_input("pH", key="bs_agenda_ph_07092026")
'''


class TestWidgetsBemStar(unittest.TestCase):
    def tela(self):
        at = AppTest.from_string(TELA).run()
        self.assertEqual(len(at.exception), 0)
        at.text_area(key="csr_obs_rel").set_value("Observação Tríad setembro")
        at.text_area(key="bs_rel_parecer_tecnico").set_value("Parecer Tríad setembro")
        at.text_input(key="bs_agenda_ph_07092026").set_value("7.2")
        at.multiselect(key="bs_rel_dias_semana").select("Segunda")
        at.number_input(key="bs_rel_freq_semana").set_value(3)
        at.run()
        self.assertEqual(len(at.exception), 0)
        return at

    def test_trocar_cliente_limpa_widgets_e_restaurar_recupera_edicoes(self):
        at = self.tela()
        at.selectbox(key="csr_sel_relatorio").select("Gama TESTE").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.text_area(key="csr_obs_rel").value, "")
        self.assertEqual(at.text_area(key="bs_rel_parecer_tecnico").value, "")
        self.assertEqual(at.text_input(key="bs_agenda_ph_07092026").value, "8.0")
        self.assertEqual(at.multiselect(key="bs_rel_dias_semana").value, [])
        self.assertEqual(at.number_input(key="bs_rel_freq_semana").value, 1)
        at.selectbox(key="csr_sel_relatorio").select("Tríad TESTE").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.text_area(key="csr_obs_rel").value, "Observação Tríad setembro")
        self.assertEqual(at.text_area(key="bs_rel_parecer_tecnico").value, "Parecer Tríad setembro")
        self.assertEqual(at.text_input(key="bs_agenda_ph_07092026").value, "7.2")
        self.assertEqual(at.multiselect(key="bs_rel_dias_semana").value, ["Segunda"])
        self.assertEqual(at.number_input(key="bs_rel_freq_semana").value, 3)

    def test_trocar_mes_nao_reaproveita_parecer_e_retorno_preserva(self):
        at = self.tela()
        at.selectbox(key="csr_mes_rel").select("10").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.text_area(key="bs_rel_parecer_tecnico").value, "")
        at.selectbox(key="csr_mes_rel").select("09").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.text_area(key="bs_rel_parecer_tecnico").value, "Parecer Tríad setembro")


if __name__ == "__main__":
    unittest.main()
