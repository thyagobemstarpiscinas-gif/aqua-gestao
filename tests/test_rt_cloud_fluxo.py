"""Fluxo do renderizador/aplicadores reais por AST, Sheets e documentos simulados.

Não é AppTest, navegador nem homologação com credenciais reais.
"""
import copy
from datetime import datetime
from pathlib import Path
import shutil
import unittest
from unittest.mock import Mock

import test_ultimos_sheets as suporte
from test_rt_ultimos_regressoes import carregar

class TestFluxoCloud(unittest.TestCase):
    setUp = suporte.TestUltimosSheets.setUp

    def contexto(self, caso='sucesso', raiz=None, chave=None):
        def docx(_template, destino, *_args, **_kwargs):
            if caso == 'docx_falha':
                raise OSError('DOCX fictício indisponível')
            destino.write_bytes(b'DOCX FICTICIO')
        def pdf(_dados, _fotos, destino):
            if caso != 'pdf_falha':
                destino.write_bytes(b'PDF FICTICIO')
            return caso != 'pdf_falha', 'falha fictícia' if caso == 'pdf_falha' else ''
        ns = carregar(['_renderizar_relatorio_rt', 'recuperar_ultimos_valores_rt',
                       'aplicar_dosagens_ultimas_no_relatorio'],
            GENERATED_DIR=raiz or self.raiz, datetime=datetime,
            conectar_sheets=self.conectar,
            validar_contexto_rt=lambda: ['Contexto inválido'] if caso == 'contexto' else [],
            montar_dados_relatorio=lambda: {'nome_condominio': 'Cliente Fictício'},
            validar_relatorio_mensal=lambda _: [],
            obter_snapshot_relatorio_independente=lambda: copy.deepcopy(self.snapshot),
            salvar_dados_condominio=lambda *_: None,
            _resolver_fotos_relatorio_rt=lambda *a, **k: ([], 'teste'),
            limpar_nome_arquivo=lambda n: n, preencher_relatorio_mensal_docx=docx,
            TEMPLATE_RELATORIO=None, gerar_pdf_relatorio_rt_premium_reportlab=pdf,
            converter_docx_para_pdf=lambda *_: (False, 'fallback fictício indisponível'),
            registrar_documento_manifest=Mock(side_effect=OSError('Manifest fictício') if caso == 'manifest_falha' else None),
            obter_status_art_texto=lambda _: 'fictício')
        ns['st'].warning = Mock()
        ns['st'].session_state.update(rel_origem_dados=chave or self.chave,
            rel_analises_total=12, rel_analise_data_0='15/10/2026',
            rel_analise_ph_0='ATUAL', rel_dos_produto_0='ATUAL')
        return ns

    def test_final_nova_sessao_sem_disco_recupera_ambos(self):
        ns = self.contexto()
        resultado = ns['_renderizar_relatorio_rt']()
        self.assertTrue(resultado['pdf_ok'])
        self.assertEqual(len(self.sh.aba.rows), 2)
        # A simulação perde todo o disco da sessão que gerou o relatório.
        shutil.rmtree(self.raiz)
        novo = self.raiz / 'servidor_novo'
        ns2 = self.contexto(raiz=novo, chave='aqua_gestao|id:C1|10/2026')
        ss = ns2['st'].session_state
        for tipo in ('parametros', 'dosagens'):
            ns2['recuperar_ultimos_valores_rt'](tipo)
            self.assertIn('09/2026', ss['_rt_ultimos_aviso'])
        self.assertEqual(ss['rel_analise_ph_0'], '7.3')
        self.assertEqual(ss['rel_dos_produto_0'], 'Cloro')
        self.assertEqual(ss['rel_analise_data_0'], '15/10/2026')
        self.assertFalse(novo.exists())
        self.assertEqual(len(self.sh.aba.rows), 2)

    def test_previa_falhas_e_contexto_nao_publicam(self):
        self.contexto()['_renderizar_relatorio_rt']()
        antes = copy.deepcopy(self.sh.aba.rows)
        for caso in ('previa', 'docx_falha', 'pdf_falha', 'manifest_falha', 'contexto'):
            with self.subTest(caso=caso):
                ns = self.contexto(caso=caso)
                if caso in ('docx_falha', 'manifest_falha'):
                    with self.assertRaises(OSError):
                        ns['_renderizar_relatorio_rt']()
                else:
                    ns['_renderizar_relatorio_rt'](preview=caso == 'previa')
                self.assertEqual(self.sh.aba.rows, antes)

    def test_falha_sheets_preserva_documentos_e_avisa(self):
        ns = self.contexto()
        self.sh.aba.falhar = True
        resultado = ns['_renderizar_relatorio_rt']()
        self.assertTrue(resultado['docx'].is_file())
        self.assertTrue(resultado['pdf'].is_file())
        ns['st'].warning.assert_called_once()
        self.assertIn('últimos valores não foram atualizados', ns['st'].warning.call_args.args[0])
        self.assertEqual(len(self.sh.aba.rows), 1)

    def test_recuperacao_apos_final_nao_contamina_outros_clientes(self):
        self.contexto()['_renderizar_relatorio_rt']()
        for chave in ('bem_star|id:C1|10/2026', 'aqua_gestao|id:C2|10/2026'):
            for tipo in ('parametros', 'dosagens'):
                with self.subTest(chave=chave, tipo=tipo):
                    ns = self.contexto(chave=chave)
                    ss = ns['st'].session_state
                    antes = copy.deepcopy(dict(ss))
                    ns['recuperar_ultimos_valores_rt'](tipo)
                    self.assertIn('Não foi possível', ss.pop('_rt_ultimos_aviso'))
                    self.assertEqual(dict(ss), antes)

if __name__ == '__main__':
    unittest.main(verbosity=2)
