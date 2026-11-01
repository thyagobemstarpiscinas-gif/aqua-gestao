"""Regressão observada na produção: adulto de 173 m³ era tratado como não cadastrado."""
import ast
import unittest
from pathlib import Path
from utils.cadastro_piscinas import serializar

ROOT = Path(__file__).resolve().parents[1]
TREE = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
NS = {}
exec(compile(ast.Module(body=[next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and
                                n.name == 'resolver_piscinas_operador')], type_ignores=[]), 'app.py', 'exec'), NS)
resolver = NS['resolver_piscinas_operador']


class PiscinasOperador(unittest.TestCase):
    def test_adulto_remoto_sem_json_local(self):
        self.assertEqual(resolver({'vol_adulto': 173}, {}),
                         {'piscinas': ['Piscina Adulto'], 'configuradas': True, 'fonte': 'cadastro'})

    def test_adulto_local_explicito_tambem_e_configuracao(self):
        self.assertTrue(resolver({}, {'piscinas': ['Piscina Adulto']})['configuradas'])

    def test_remoto_prevalece_sobre_local_desatualizado(self):
        self.assertEqual(resolver({'vol_adulto': 173}, {'piscinas': ['Piscina Infantil']})['piscinas'],
                         ['Piscina Adulto'])

    def test_extras_e_tipos_sem_volume(self):
        cliente = {'cadastro_piscinas': serializar({'adulto': 'Adulto', 'extra1': 'SPA'},
                                                  [{'nome': 'SPA Cobertura', 'vol': 4, 'tipo': 'SPA'}])}
        self.assertEqual(resolver(cliente, {})['piscinas'], ['Piscina Adulto', 'SPA Cobertura'])

    def test_total_sem_distribuicao_nao_inventa_adulto(self):
        self.assertFalse(resolver({'vol_total': 173}, {})['configuradas'])

    def test_vazio_nao_e_configurado(self):
        self.assertFalse(resolver(None, {'piscinas': []})['configuradas'])

from streamlit.testing.v1 import AppTest
from test_relatorio_rt_campos import _AppReal
import sys


class PiscinasOperadorUI(_AppReal):
    def conferir_empresa(self, empresa):
        clientes = sys.modules['gspread'].PLANILHA.worksheet('👥 Clientes')
        clientes.linhas[1][9] = '173'
        clientes.linhas[1][12] = empresa
        at = AppTest.from_file(str(self.pasta / 'app' / 'app.py'), default_timeout=300)
        at.secrets['operador'] = {'pin_mestre': 'op-teste-123'}
        at.secrets['gcp_service_account'] = {'tipo': 'ficticio'}
        at.run()
        next(b for b in at.button if b.label == '📱 Acessar como Operador').click().run()
        next(w for w in at.text_input if w.label == 'PIN').input('op-teste-123')
        next(b for b in at.button if b.label == 'Entrar').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertTrue(any('1 piscina(s) deste local' in x.value for x in at.info), {'info':[x.value for x in at.info], 'seletores':[(x.label,x.value) for x in at.selectbox], 'erros':[x.value for x in at.error]})
        self.assertFalse(any(b.label == '💾 Salvar configuração de piscinas' for b in at.button))
        self.assertFalse(any('Modo RT ativo' in x.value for x in at.info))
        self.assertEqual([x for x in at.text_input if x.label == 'pH'].__len__(), 1)

    def test_operador_aqua_sem_pasta_local(self):
        self.conferir_empresa('Aqua Gestão')

    def test_operador_bem_star_sem_pasta_local(self):
        self.conferir_empresa('Bem Star Piscinas')
