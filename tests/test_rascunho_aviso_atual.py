"""Abertura do painel confirma a fila antes do aviso, sem editar o relatório."""
import ast
import copy
import unittest
from unittest.mock import Mock, patch

from test_rascunho_rt import APP, TREE, Ambiente


class AvisoRascunhoAtual(unittest.TestCase):
    def setUp(self):
        self.amb = Ambiente()
        self.s = self.amb.sessao()
        self.s['_rascunho_usa_sheets'] = lambda: True

    def pendencia(self):
        self.s['ss']['rel_diagnostico'] = 'Conteúdo fictício preservado'
        self.amb.criar_aba().falhar = True
        self.assertFalse(self.s['_relatorio_rt_salvar_rascunho']())
        self.amb.aba().falhar = False
        return self.s['ss']['_rt_rasc_chave']

    def painel(self, s):
        st = s['st']
        for nome in ('caption', 'warning', 'error', 'success', 'info', 'markdown'):
            setattr(st, nome, Mock())
        func = next(n for n in TREE.body if isinstance(n, ast.FunctionDef)
                    and n.name == '_relatorio_rt_renderizar_painel_rascunho')
        exec(compile(ast.Module(body=[func], type_ignores=[]), str(APP), 'exec'), s)
        s['_relatorio_rt_renderizar_painel_rascunho']()
        self.assertFalse(any('indisponível' in str(c) for c in st.caption.call_args_list))
        return st

    def test_reabrir_sem_edicao_confirma_antes_de_renderizar(self):
        chave = self.pendencia()
        nova = self.amb.sessao()
        nova['_rascunho_usa_sheets'] = lambda: True
        campos = copy.deepcopy(nova['ss'])
        st = self.painel(nova)
        self.assertEqual(nova['_relatorio_rt_gestor_sync']().ler(chave)['estado'], 'confirmado')
        self.assertEqual(len(self.amb.linhas(chave)), 1)
        self.assertNotIn('_rt_rasc_aviso', nova['ss'])
        st.error.assert_not_called()
        for k, v in campos.items():
            if k.startswith('rel_'):
                self.assertEqual(nova['ss'][k], v)

    def test_falha_real_mantem_aviso_e_copia_local(self):
        chave = self.pendencia()
        nova = self.amb.sessao()
        nova['_rascunho_usa_sheets'] = lambda: True
        self.amb.aba().falhar = True
        st = self.painel(nova)
        self.assertEqual(nova['_relatorio_rt_gestor_sync']().ler(chave)['estado'], 'pendente')
        self.assertTrue(st.error.called)
        self.assertEqual(self.amb.linhas(chave), [])
        self.assertIn('falha simulada', str(st.error.call_args_list))

    def test_painel_respeita_intervalo_entre_tentativas(self):
        chave = self.pendencia()
        self.s['ss']['_rt_rasc_proxima_tentativa'] = float('inf')
        st = self.painel(self.s)
        self.assertEqual(self.s['_relatorio_rt_gestor_sync']().ler(chave)['estado'], 'pendente')
        self.assertTrue(st.error.called)
        self.assertEqual(self.amb.linhas(chave), [])

    def test_backend_local_nao_consulta_rede(self):
        self.s['_rascunho_backend'] = lambda: 'local'
        self.s['_rascunho_usa_sheets'] = lambda: False
        with patch.dict(self.s, {'_relatorio_rt_sincronizar_pendentes': Mock(side_effect=AssertionError)}):
            self.painel(self.s)
            self.s['_relatorio_rt_sincronizar_pendentes'].assert_not_called()

    def test_salvar_formulario_vazio_informa_ausencia_de_edicao(self):
        self.assertEqual(self.s['relatorio_rt_resultado_salvamento_manual'](), 'vazio')
        st = self.s['st']
        for nome in ('info', 'warning', 'error', 'success'):
            setattr(st, nome, Mock())
        st.button = Mock(return_value=True)
        source = APP.read_text()
        start = source.index('if st.button("💾 Salvar rascunho", use_container_width=True):')
        end = source.index('\nst.markdown("---")', start)
        exec(compile(source[start:end], str(APP), 'exec'), self.s)
        st.info.assert_called_once_with('O formulário está vazio; não há novas alterações para salvar.')
        st.warning.assert_not_called()
        st.error.assert_not_called()
        self.assertEqual(self.amb.linhas(), [])
        self.assertEqual(list(self.amb.dir.glob('*.json')), [])


if __name__ == '__main__':
    unittest.main()
