"""Widgets reais e contextos fictícios, sem rede."""
import unittest
from streamlit.testing.v1 import AppTest
from utils.bem_star_mensal import chave_campo, token_pool

TELA = '''
import streamlit as st
from utils.contexto_bem_star import preparar_contexto_bem_star,preencher_campo_agenda_bem_star
from utils.bem_star_mensal_ui import renderizar_piscinas,renderizar_agenda,renderizar_fotos
cliente=st.selectbox('Cliente',['Cliente A TESTE','Cliente B TESTE'])
mes=st.selectbox('Mês',['09','10'])
preparar_contexto_bem_star(st.session_state,cliente,mes,'2026',cliente)
ls=[{'data':'07/'+mes+'/2026','piscinas':[
    {'nome':'Piscina adulta','ph':'7.4' if cliente=='Cliente A TESTE' else '7.8'},
    {'nome':'Piscina infantil','ph':'7.1' if cliente=='Cliente A TESTE' else ''}]}]
renderizar_piscinas(st,ls)
renderizar_agenda(st,ls,['07/'+mes+'/2026'],lambda k,v:preencher_campo_agenda_bem_star(st.session_state,k,v))
renderizar_fotos(st)
'''


class WidgetsMensal(unittest.TestCase):
    def test_duas_piscinas_nao_misturam_e_cliente_retorna_editado(self):
        at=AppTest.from_string(TELA).run()
        self.assertEqual(len(at.exception),0)
        ka=chave_campo('Piscina adulta','07/09/2026','ph');ki=chave_campo('Piscina infantil','07/09/2026','ph')
        self.assertEqual(at.text_input(key=ka).value,'7.4');self.assertEqual(at.text_input(key=ki).value,'7.1')
        at.text_input(key=ka).set_value('7.2').run()
        at.selectbox[0].select('Cliente B TESTE').run()
        self.assertEqual(len(at.exception),0)
        self.assertEqual(at.text_input(key=ka).value,'7.8');self.assertEqual(at.text_input(key=ki).value,'')
        at.selectbox[0].select('Cliente A TESTE').run()
        self.assertEqual(len(at.exception),0)
        self.assertEqual(at.text_input(key=ka).value,'7.2');self.assertEqual(at.text_input(key=ki).value,'7.1')

    def test_tipo_de_piscina_e_selecao_sao_separados_por_mes(self):
        at=AppTest.from_string(TELA).run(); k='bs_mensal_tipo_'+token_pool('Piscina adulta')
        at.selectbox(key=k).select('Outros').run()
        at.selectbox[1].select('10').run()
        self.assertEqual(len(at.exception),0)
        self.assertEqual(at.selectbox(key=k).value,'Adulto')
        at.selectbox[1].select('09').run()
        self.assertEqual(at.selectbox(key=k).value,'Outros')
