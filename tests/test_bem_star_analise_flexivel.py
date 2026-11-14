"""Escolha da coleta semanal pelo PIN de Thyago; dados fictícios e sem rede."""
import ast
import json
import unittest
from pathlib import Path
from types import SimpleNamespace

from streamlit.testing.v1 import AppTest

from utils.bem_star_mensal import (analise_semanal_no_formulario,
    filtrar_analises_semanais, restaurar_analise_semanal)
from test_bem_star_analise_semanal import visita
import test_bem_star_analise_semanal as semanal
from test_operador_formulario import _sessao
from test_homologacao_pr14 import _Homologacao, EPSILON


def acesso(nome='Thyago Silveira', rt=False):
    return {'op_pin_ok': True, 'op_dados_atual': {'nome': nome, 'is_rt': rt},
            'op_operador': nome, 'op_data_visita': '13/10/2026'}


class EscolhaSemanal(unittest.TestCase):
    def test_terca_marcada_substitui_segunda_desmarcada_sem_mudar_datas(self):
        segunda = dict(visita('12/10/2026'), analise_semanal_bem_star=False)
        terca = dict(visita('13/10/2026', vid='TERCA'), analise_semanal_bem_star=True)
        extra = visita('14/10/2026', vid='QUARTA')
        self.assertEqual(filtrar_analises_semanais([segunda, terca, extra]), [terca])
        self.assertEqual(terca['data'], '13/10/2026')
        self.assertEqual(segunda['data'], '12/10/2026')

    def test_legado_preservado_outros_operadores_e_valores_invalidos_excluidos(self):
        antigo = visita('12/10/2026')
        entradas = [antigo, visita('13/10/2026'),
                    dict(visita('13/10/2026', 'João'), analise_semanal_bem_star=True),
                    dict(visita('12/10/2026'), analise_semanal_bem_star='true'),
                    dict(visita('inválida'), analise_semanal_bem_star=True)]
        self.assertEqual(filtrar_analises_semanais(entradas), [antigo])

    def test_escolha_depende_do_acesso_autenticado_e_nao_do_nome_restaurado(self):
        for nome, rt, pin in [('João', False, True), ('Thyago Silveira', True, True),
                              ('Thyago Silveira', False, False)]:
            with self.subTest(nome=nome, rt=rt, pin=pin):
                ss = acesso(nome, rt)
                ss.update(op_pin_ok=pin, op_operador='Thyago Silveira',
                          op_analise_semanal_bem_star=True)
                self.assertFalse(analise_semanal_no_formulario(ss))

    def test_rascunho_restauracao_e_troca_nao_transferem_escolha(self):
        a = _sessao(); a['ss'].update(acesso(), op_analise_semanal_bem_star=True,
                                    op_adulto_ph='7.4')
        rasc = json.loads(json.dumps(a['coletar_rascunho_operador']('Alfa TESTE', ['Piscina Adulto'])))
        self.assertTrue(rasc['analise_semanal_bem_star'])
        b = _sessao(); b['ss'].update(acesso())
        b['restaurar_rascunho_operador'](rasc)
        self.assertTrue(analise_semanal_no_formulario(b['ss']))
        self.assertEqual(b['ss']['op_data_visita'], '13/10/2026')
        self.assertTrue(b['trocar_condominio_operador']('Alfa TESTE', 'Beta TESTE'))
        self.assertNotIn('op_analise_semanal_bem_star', b['ss'])
        self.assertFalse(analise_semanal_no_formulario(b['ss']))
        self.assertTrue(b['salvos'][-1][1]['analise_semanal_bem_star'])
        b['restaurar_rascunho_operador'](rasc)
        self.assertTrue(analise_semanal_no_formulario(b['ss']))
        # Rascunho de outro operador não transfere sua marca para o PIN de Thyago.
        restaurar_analise_semanal(b['ss'], dict(rasc, operador='João'))
        self.assertNotIn('op_analise_semanal_bem_star', b['ss'])

    def test_desmarcado_na_segunda_persiste_e_legado_nao_herda_escolha_anterior(self):
        a = _sessao(); a['ss'].update(acesso(), op_data_visita='12/10/2026',
                                    op_analise_semanal_bem_star=False)
        rasc = a['coletar_rascunho_operador']('Alfa TESTE', [])
        b = _sessao(); b['ss'].update(acesso(), op_analise_semanal_bem_star=True)
        b['restaurar_rascunho_operador'](rasc)
        self.assertFalse(analise_semanal_no_formulario(b['ss']))
        rasc.pop('analise_semanal_bem_star')
        b['restaurar_rascunho_operador'](rasc)
        self.assertTrue(analise_semanal_no_formulario(b['ss']))

    def test_mensal_inclui_terca_e_cc_sem_anunciar_todas_as_coletas_como_segunda(self):
        segunda = dict(visita('12/10/2026'), analise_semanal_bem_star=False)
        segunda['piscinas'][0]['ph'] = '9.9'
        terca = dict(visita('13/10/2026', vid='TERCA'), analise_semanal_bem_star=True)
        paginas = semanal.MensalPDFSemanal.gerar(self, [segunda, terca])
        adulto = next(p for p in paginas if 'Resumo numérico da piscina' in p and 'Piscina Adulto' in p)
        self.assertIn('13/10/2026', adulto)
        self.assertNotIn('12/10/2026', adulto)
        self.assertNotIn('9,9', adulto)
        self.assertIn('0,3', adulto)
        self.assertNotIn('Análises de segunda-feira', '\n'.join(paginas))


TELA = '''
import streamlit as st
from utils.bem_star_mensal import analise_semanal_no_formulario
from utils.bem_star_mensal_ui import renderizar_analise_semanal_operador
def salvar():
    st.session_state['guardado'] = {
        'data': st.session_state['op_data_visita'],
        'analise_semanal_bem_star': analise_semanal_no_formulario(st.session_state)}
st.text_input('Data da visita', key='op_data_visita')
renderizar_analise_semanal_operador(st, salvar)
st.text_input('Observação', key='observacao')
'''


class WidgetSemanal(unittest.TestCase):
    def test_padrao_segunda_escolha_terca_rerun_e_autosave(self):
        at = AppTest.from_string(TELA, default_timeout=30)
        for k, v in dict(acesso(), op_data_visita='12/10/2026').items():
            at.session_state[k] = v
        at.run(); self.assertFalse(at.exception)
        self.assertTrue(at.checkbox[0].value)
        self.assertNotIn('guardado', at.session_state)  # abrir não sobrescreve o rascunho
        at.text_input(key='op_data_visita').set_value('13/10/2026').run()
        self.assertFalse(at.checkbox[0].value)
        at.checkbox[0].check().run(); self.assertFalse(at.exception)
        self.assertEqual(at.session_state['guardado'],
                         {'data': '13/10/2026', 'analise_semanal_bem_star': True})
        at.text_input(key='observacao').set_value('Medição de terça').run()
        self.assertTrue(at.checkbox[0].value)
        at.checkbox[0].uncheck().run()
        self.assertFalse(at.session_state['guardado']['analise_semanal_bem_star'])

    def test_outro_pin_ou_rt_nao_exibe_opcao_e_nao_grava(self):
        for nome, rt in [('João', False), ('Thyago Silveira', True)]:
            with self.subTest(nome=nome, rt=rt):
                at = AppTest.from_string(TELA, default_timeout=30)
                for k, v in dict(acesso(nome, rt), op_analise_semanal_bem_star=True).items():
                    at.session_state[k] = v
                at.run(); self.assertFalse(at.exception)
                self.assertEqual(len(at.checkbox), 0)
                self.assertNotIn('op_analise_semanal_bem_star', at.session_state)
                self.assertNotIn('guardado', at.session_state)


class PersistenciaSemanalSheets(_Homologacao):
    def test_marca_real_do_formulario_vai_ao_payload_e_volta_do_sheets(self):
        # Usa a expressão do payload real, e as rotinas reais de gravação/leitura
        # com Sheets fictício. O nome digitado não consegue habilitar a marca.
        source = (Path(__file__).resolve().parents[1] / 'app.py').read_text()
        tree = ast.parse(source)
        payload = next(n.value for n in ast.walk(tree) if isinstance(n, ast.Assign)
                       and any(isinstance(t, ast.Name) and t.id == 'lancamento' for t in n.targets)
                       and isinstance(n.value, ast.Dict)
                       and any(isinstance(k, ast.Constant) and k.value == 'analise_semanal_bem_star'
                               for k in n.value.keys))
        expr = next(v for k, v in zip(payload.keys, payload.values)
                    if isinstance(k, ast.Constant) and k.value == 'analise_semanal_bem_star')
        ss = acesso(); ss['op_analise_semanal_bem_star'] = True
        marcado = eval(compile(ast.Expression(expr), 'payload_real', 'eval'),
                       {'st': SimpleNamespace(session_state=ss),
                        'analise_semanal_no_formulario': analise_semanal_no_formulario})
        v = self._lancamento_operador()
        v.update(data='13/10/2026', operador='Thyago Silveira', analise_semanal_bem_star=marcado)
        self.assertTrue(self.app['sheets_salvar_lancamento_campo'](v, EPSILON))
        recuperada = next(x for x in self.app['situacao_visitas_condominio'](EPSILON)['visitas']
                          if x['id_visita'] == v['id_visita'])
        self.assertIs(recuperada['analise_semanal_bem_star'], True)
        self.assertEqual(recuperada['data'], '13/10/2026')
        self.assertEqual(filtrar_analises_semanais([recuperada]), [recuperada])


if __name__ == '__main__':
    unittest.main()
