"""Menu do escritório: destinos reais, empresa fixa e formulários preservados."""
import re
import shutil
import sys
import unittest
from unittest.mock import patch

from utils.navegacao import empresa_do_escritorio, html_menu, links_menu
from test_relatorio_rt_campos import _AppReal, AppTest, PIN_FICTICIO


class MenuPorPerfil(unittest.TestCase):
    def test_empresa_fixa_prevalece_e_entrada_operador_nao_recebem_menu(self):
        ss = dict(modo_atual='escritorio', admin_logado=True,
                  admin_empresa_fixa='bem_star', empresa_ativa='aqua_gestao')
        self.assertEqual(empresa_do_escritorio(ss), 'bem_star')
        for modo, admin in [('entrada', False), ('operador', False), ('operador', True), ('escritorio', False)]:
            self.assertIsNone(empresa_do_escritorio(dict(ss, modo_atual=modo, admin_logado=admin)))
        self.assertIsNone(empresa_do_escritorio(dict(ss, admin_empresa_fixa='desconhecida')))

    def test_destinos_por_empresa_e_links_sem_nova_pagina(self):
        aqua = html_menu('aqua_gestao')
        bem = html_menu('bem_star')
        self.assertIn('Relatório mensal RT', aqua)
        self.assertNotIn('Relatório mensal RT', bem)
        self.assertIn('Relatório mensal operacional', bem)
        self.assertNotIn('nav-mensal-bs', aqua)
        for empresa, html in [('aqua_gestao', aqua), ('bem_star', bem)]:
            destinos = re.findall(r'href="([^"]+)"', html)
            self.assertEqual(len(destinos), len(set(destinos)))
            self.assertEqual(destinos, ['#' + alvo for alvo, _ in links_menu(empresa)])
            self.assertEqual(html.count('target="_self"'), len(destinos))
            self.assertNotIn('<script', html)
        self.assertEqual(html_menu('inexistente'), '')


class MenuNoApp(_AppReal):
    def setUp(self):
        # O ambiente Python não pertence à cópia fictícia nem à prova do app.
        ignorar = shutil.ignore_patterns
        with patch('shutil.ignore_patterns', side_effect=lambda *nomes: ignorar(*nomes, '.venv')):
            super().setUp()
        sys.modules['gspread'].PLANILHA.worksheet('👥 Clientes').linhas.append(
            ['', 'C3', 'Condomínio Gama Fictício', '60', '', 'Síndico Gama', 'Rua C, 3', '',
             'Ativo', '60', '', '', 'Bem Star Piscinas', '', ''])

    def _app(self):
        at = AppTest.from_file(str(self.pasta / 'app/app.py'), default_timeout=120)
        at.secrets['admin'] = {'pin': PIN_FICTICIO}
        at.secrets['gcp_service_account'] = {'tipo': 'ficticio'}
        return at

    def _entrar(self, empresa):
        at = self._app().run()
        self.assertFalse(at.exception)
        self.assertNotIn('Navegação do escritório', ''.join(x.value for x in at.markdown))
        at.button(key='btn_admin_limpo').click().run()
        at.button(key='btn_admin_login_bemstar' if empresa == 'bem_star' else 'btn_admin_login_aqua').click().run()
        at.text_input(key='pin_admin_input').input(PIN_FICTICIO)
        at.button(key='btn_pin_admin_ok').click().run()
        self.assertFalse(at.exception, at.exception)
        return at

    def _conferir_destinos(self, at, empresa):
        html = '\n'.join(x.value for x in at.markdown)
        for alvo, _ in links_menu(empresa):
            self.assertEqual(html.count(f'id="{alvo}"'), 1, alvo)
        self.assertEqual(sum(b.label == '🚪 Sair do admin' for b in at.button), 1)
        self.assertEqual(next(e for e in at.expander if e.label == 'Histórico e arquivos').proto.expanded, False)
        self.assertEqual(next(e for e in at.expander if e.label == 'Diagnóstico do sistema').proto.expanded, False)

    def test_aqua_destinos_estado_formulario_e_saida(self):
        at = self._entrar('aqua_gestao')
        self._conferir_destinos(at, 'aqua_gestao')
        at.text_area(key='rel_diagnostico').input('PARECER FICTÍCIO PRESERVADO').run()
        at.text_input(key='busca_rapida').input('Alfa').run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(at.text_area(key='rel_diagnostico').value, 'PARECER FICTÍCIO PRESERVADO')
        self._conferir_destinos(at, 'aqua_gestao')
        at.button(key='btn_trocar_empresa_login_admin').click().run()
        self.assertFalse(at.exception)
        self.assertNotIn('Navegação do escritório', ''.join(x.value for x in at.markdown))

    def test_bem_star_destinos_e_isolamento(self):
        at = self._entrar('bem_star')
        self._conferir_destinos(at, 'bem_star')
        html = '\n'.join(x.value for x in at.markdown)
        self.assertNotIn('id="sec-relatorio-rt"', html)
        self.assertNotRegex(html, r'<a\b[^>]*href="#sec-relatorio-rt"')
        at.text_input(key='csr_mes_rel').input('10').run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(at.text_input(key='csr_mes_rel').value, '10')

    def test_operador_nao_renderiza_widgets_do_escritorio(self):
        at = self._app()
        at.session_state['modo_atual'] = 'operador'
        at.session_state['op_pin_ok'] = True
        at.session_state['op_dados_atual'] = {'nome': 'Operador Fictício', 'condomínios': ['Condomínio Alfa Fictício']}
        at.run()
        self.assertFalse(at.exception, at.exception)
        self.assertNotIn('Navegação do escritório', ''.join(x.value for x in at.markdown))
        self.assertNotIn('filtro_historico', [x.key for x in at.text_input])
        self.assertTrue(any(b.key == 'btn_voltar_inicio' for b in at.button))


if __name__ == '__main__':
    unittest.main()
