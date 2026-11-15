"""Gestão de status nas duas empresas, com planilha fictícia e rede bloqueada."""
import copy
import sys
from unittest.mock import patch

from streamlit.testing.v1 import AppTest
from test_relatorio_rt_campos import _AppReal, PIN_FICTICIO


class ClientesInativos(_AppReal):
    def setUp(self):
        super().setUp()
        self.clientes = sys.modules['gspread'].PLANILHA.worksheet('👥 Clientes')
        self.clientes.linhas[2][12] = 'Bem Star Piscinas'
        for cid, nome, empresa, status in [
            ('C10', 'Cliente Antigo Aqua TESTE', 'Aqua Gestão', 'Inativo'),
            ('C11', 'Outro Antigo Aqua TESTE', 'Aqua Gestão', 'Desativado'),
            ('C20', 'Cliente Antigo Bem TESTE', 'Bem Star Piscinas', 'Inativo'),
        ]:
            self.clientes.linhas.append(['', cid, nome, '20', '', 'Contato TESTE',
                                         'Rua TESTE', '', status, '20', '', '', empresa, '', ''])

    def entrar(self, empresa):
        at = AppTest.from_file(str(self.pasta / 'app/app.py'), default_timeout=120)
        at.secrets['admin'] = {'pin': PIN_FICTICIO}
        at.secrets['gcp_service_account'] = {'tipo': 'ficticio'}
        at.run()
        at.button(key='btn_admin_limpo').click().run()
        at.button(key='btn_admin_login_bemstar' if empresa == 'bem_star' else 'btn_admin_login_aqua').click().run()
        at.text_input(key='pin_admin_input').input(PIN_FICTICIO)
        at.button(key='btn_pin_admin_ok').click().run()
        at.radio(key='cc_modo_acao').set_value('✏️ Editar cliente existente').run()
        self.assertFalse(at.exception, at.exception)
        return at

    def linha(self, cid):
        return next(r for r in self.clientes.linhas if len(r) > 1 and r[1] == cid)

    def test_aqua_filtros_reativacao_confirma_por_cliente_e_preserva_historico(self):
        at = self.entrar('aqua_gestao')
        self.assertEqual(at.selectbox(key='cc_sel_editar').options, ['Condomínio Alfa Fictício'])
        at.selectbox(key='cc_filtro_status').select('Todos').run()
        self.assertEqual(len(at.selectbox(key='cc_sel_editar').options), 3)
        self.assertNotIn('Cliente Antigo Aqua TESTE', at.selectbox(key='sel_cliente_rel').options)
        at.selectbox(key='cc_filtro_status').select('Inativos').run()
        self.assertEqual(at.selectbox(key='cc_sel_editar').options,
                         ['Cliente Antigo Aqua TESTE', 'Outro Antigo Aqua TESTE'])
        self.assertTrue(any('Cliente inativo' in t.value for t in at.info))
        self.assertTrue(at.button(key='btn_reativar_cliente').disabled)
        at.checkbox(key='cc_confirmar_reativar_aqua_gestao_C10').check().run()
        at.selectbox(key='cc_sel_editar').select('Outro Antigo Aqua TESTE').run()
        self.assertTrue(at.button(key='btn_reativar_cliente').disabled)
        at.selectbox(key='cc_sel_editar').select('Cliente Antigo Aqua TESTE').run()
        antes = copy.deepcopy(self.clientes.linhas)
        visitas = sys.modules['gspread'].PLANILHA.worksheet('🔬 Visitas')
        historico = copy.deepcopy(visitas.linhas)
        at.checkbox(key='cc_confirmar_reativar_aqua_gestao_C10').check().run()
        at.button(key='btn_reativar_cliente').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertTrue(any('reativado' in t.value for t in at.success))
        esperado = copy.deepcopy(antes)
        next(r for r in esperado if len(r) > 1 and r[1] == 'C10')[8] = 'Ativo'
        self.assertEqual(self.clientes.linhas, esperado)
        self.assertEqual(visitas.linhas, historico)
        self.assertNotIn('Cliente Antigo Aqua TESTE', at.selectbox(key='cc_sel_editar').options)
        self.assertIn('Cliente Antigo Aqua TESTE', at.selectbox(key='sel_cliente_rel').options)
        at.selectbox(key='cc_filtro_status').select('Ativos').run()
        self.assertIn('Cliente Antigo Aqua TESTE', at.selectbox(key='cc_sel_editar').options)

    def test_bem_star_edicao_e_transferencia_nao_reativam_inativo(self):
        at = self.entrar('bem_star')
        self.assertEqual(at.selectbox(key='cc_sel_editar').options, ['Residencial Beta Fictício'])
        at.selectbox(key='cc_filtro_status').select('Inativos').run()
        self.assertEqual(at.selectbox(key='cc_sel_editar').options, ['Cliente Antigo Bem TESTE'])
        at.button(key='btn_carregar_editar').click().run()
        at.text_input(key='cc_contato').input('Contato atualizado TESTE')
        next(b for b in at.button if b.label == '💾 Salvar alterações').click().run()
        self.assertEqual(self.linha('C20')[8], 'Inativo')
        self.assertEqual(self.linha('C20')[12], 'Bem Star Piscinas')
        at.selectbox(key='cc_transferir_empresa_destino_bem_star_C20').select('Aqua Gestão').run()
        at.button(key='btn_transferir_cliente_empresa').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.linha('C20')[8], 'Inativo')
        self.assertEqual(self.linha('C20')[12], 'Aqua Gestão')
        self.assertTrue(any('Nenhum cliente no filtro inativos' in t.value for t in at.info))

    def test_falha_na_reativacao_preserva_status_e_nao_informa_sucesso(self):
        at = self.entrar('bem_star')
        at.selectbox(key='cc_filtro_status').select('Inativos').run()
        at.checkbox(key='cc_confirmar_reativar_bem_star_C20').check().run()
        antes = copy.deepcopy(self.clientes.linhas)
        with patch.object(self.clientes, 'update', side_effect=RuntimeError('Falha fictícia de escrita')):
            at.button(key='btn_reativar_cliente').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.clientes.linhas, antes)
        self.assertTrue(any('Não foi possível reativar' in t.value for t in at.error))
        self.assertFalse(any('reativado' in t.value for t in at.success))

    def test_bem_star_reativa_e_desativa_sem_alterar_historico_ou_empresa(self):
        at = self.entrar('bem_star')
        antes = copy.deepcopy(self.clientes.linhas)
        self.assertNotIn('Cliente Antigo Bem TESTE', at.selectbox(key='csr_sel_relatorio').options)

        at.selectbox(key='cc_filtro_status').select('Inativos').run()
        at.checkbox(key='cc_confirmar_reativar_bem_star_C20').check().run()
        at.button(key='btn_reativar_cliente').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.linha('C20')[8], 'Ativo')
        self.assertIn('Cliente Antigo Bem TESTE', at.selectbox(key='csr_sel_relatorio').options)
        at.selectbox(key='cc_filtro_status').select('Ativos').run()
        at.selectbox(key='cc_sel_editar').select('Cliente Antigo Bem TESTE').run()
        self.assertTrue(at.button(key='btn_desativar_cliente').disabled)
        at.checkbox(key='cc_confirmar_desativar_bem_star_C20').check().run()
        at.button(key='btn_desativar_cliente').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.clientes.linhas, antes)
        self.assertNotIn('Cliente Antigo Bem TESTE', at.selectbox(key='csr_sel_relatorio').options)

    def test_clientes_homonimos_mantem_ids_e_confirmacoes_distintos(self):
        self.linha('C11')[2] = self.linha('C10')[2]
        at = self.entrar('aqua_gestao')
        at.selectbox(key='cc_filtro_status').select('Inativos').run()
        opcoes = at.selectbox(key='cc_sel_editar').options
        self.assertEqual(len(set(opcoes)), 2)
        self.assertIn('C10', opcoes[0])
        self.assertIn('C11', opcoes[1])
        at.selectbox(key='cc_sel_editar').select(opcoes[1]).run()
        at.checkbox(key='cc_confirmar_reativar_aqua_gestao_C11').check().run()
        at.button(key='btn_reativar_cliente').click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self.linha('C10')[8], 'Inativo')
        self.assertEqual(self.linha('C11')[8], 'Ativo')
