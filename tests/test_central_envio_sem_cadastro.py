"""A central não pode impedir um novo relatório quando falta cadastro local."""
import ast
import json
from pathlib import Path
import tempfile
import unittest

from streamlit.testing.v1 import AppTest


class CentralSemCadastro(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        source = (Path(__file__).resolve().parents[1] / 'app.py').read_text()
        loader = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef) and n.name == 'carregar_dados_condominio')
        cls.loader = ast.get_source_segment(source, loader)
        marker = source.index('# CENTRAL DE ENVIO DE DOCUMENTOS — AQUA GESTÃO')
        start = source.rfind('if _empresa_ativa_codigo() != "bem_star":', 0, marker)
        end = source.index('    # PAINEL DE VENCIMENTOS', marker)
        cls.central = source[start:end]

    def executar(self, conteudo=None, com_pasta=True):
        with tempfile.TemporaryDirectory() as tmp:
            pasta = Path(tmp) / 'Cliente novo'
            if com_pasta:
                pasta.mkdir()
                if conteudo is not None:
                    (pasta / 'dados.json').write_text(conteudo)
            setup = '''
import streamlit as st
import json
from pathlib import Path
GENERATED_DIR = Path(''' + repr(tmp) + ''')
DADOS_JSON_NAME = 'dados.json'
def garantir_pastas_restauradas(): pass
def _empresa_ativa_codigo(): return 'aqua_gestao'
def ancora(*args): pass
def filtrar_pastas_por_empresa(pastas, empresa): return pastas
def humanizar_nome_pasta(nome): return nome
def nomes_condominio_equivalentes(a, b): return a == b
def _coletar_documentos_email_aqua(pasta): return []
def _email_aqua_configurado(): return True, ''
def enviar_email_aqua_smtp(*args, **kwargs): raise AssertionError('Não deve enviar durante o carregamento')
st.session_state['email_cliente'] = 'outro-cliente@example.invalid'
'''
            app = AppTest.from_string(setup + '\n' + self.loader + '\n' + self.central + '\nst.subheader("Relatório mensal RT")\nst.text_input("Condomínio do novo relatório")\n').run()
            self.assertEqual([e.message for e in app.exception], [])
            self.assertTrue(any(h.value == 'Relatório mensal RT' for h in app.subheader))
            return next(t.value for t in app.text_input if t.label == 'Destinatário do e-mail')

    def test_pasta_nova_sem_json_continua_ate_relatorio(self):
        self.assertEqual(self.executar(), '')

    def test_json_invalido_ou_nulo_nao_bloqueia_formulario(self):
        for conteudo in ['null', '[]', '{incompleto']:
            with self.subTest(conteudo=conteudo):
                self.assertEqual(self.executar(conteudo), '')

    def test_sem_pasta_local_continua_ate_relatorio(self):
        self.assertEqual(self.executar(com_pasta=False), '')

    def test_cadastro_existente_preserva_destinatario_selecionado(self):
        self.assertEqual(self.executar(json.dumps({'email_cliente':'selecionado@example.invalid'})), 'selecionado@example.invalid')


if __name__ == '__main__': unittest.main()
