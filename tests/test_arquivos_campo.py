"""Persistência de fotos sem Drive e recuperação após perda do disco; dados fictícios."""
import base64
from copy import deepcopy
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from utils import arquivos_campo as ac
from utils.pdf_bem_star_mensal import serie_grafico
from test_google_oauth_homologacao import PlanilhaFalsa, WorksheetNotFound
from test_homologacao_pr14 import _Homologacao, EPSILON, _imagens_pdf
from test_operador_formulario import _sessao
from utils.formatacao import slugify_nome


class ArquivosCampo(unittest.TestCase):
    def setUp(self):
        compat = patch('gspread.WorksheetNotFound', WorksheetNotFound)
        compat.start(); self.addCleanup(compat.stop)
        self.sh = PlanilhaFalsa('Teste privado fictício', None)
        self.ctx = ['visita', 'Alfa TESTE', 'VTESTE']

    def test_texto_longo_unicode_e_foto_integra_sem_celula_excedida(self):
        dados = {'texto': 'Medição de água 💧\n' * 9000,
                 'foto': base64.b64encode(bytes(range(256)) * 1000).decode()}
        ref = ac.salvar(self.sh, self.ctx, dados)
        linhas = self.sh.worksheet(ac.ABA).get_all_values()
        self.assertGreater(len(linhas), 6)
        self.assertLessEqual(max(len(r[5].encode('utf-16-le'))//2 for r in linhas), 50000)
        self.assertEqual(ac.carregar(ref, self.ctx, linhas), dados)
        with self.assertRaises(ValueError):
            ac.carregar(ref, ['visita', 'Beta TESTE', 'VTESTE'], linhas)

    def test_parte_ausente_duplicada_ou_alterada_bloqueia_sem_retornar_versao_antiga(self):
        ref = ac.salvar(self.sh, self.ctx, {'texto': 'a' * 95000})
        originais = self.sh.worksheet(ac.ABA).get_all_values()
        for linhas in [originais[:-1], originais + originais[:1]]:
            with self.assertRaises(ValueError):
                ac.carregar(ref, self.ctx, linhas)
        alteradas = deepcopy(originais); alteradas[0][5] += 'x'
        with self.assertRaises(ValueError):
            ac.carregar(ref, self.ctx, alteradas)

    def test_escrita_parcial_nao_entrega_referencia_nem_altera_revisao_anterior(self):
        antiga = ac.salvar(self.sh, self.ctx, {'texto': 'versão preservada'})
        aba = self.sh.worksheet(ac.ABA)
        append = aba.append_rows
        def interromper(linhas, **kw):
            append(linhas[:-1], **kw)
            raise OSError('interrupção de rede simulada')
        with patch.object(aba, 'append_rows', interromper), self.assertRaises(OSError):
            ac.salvar(self.sh, self.ctx, {'texto': 'novo' * 30000})
        self.assertEqual(ac.carregar(antiga, self.ctx, aba.get_all_values()), {'texto': 'versão preservada'})

    def test_limite_bloqueia_antes_de_gravar(self):
        with patch.object(ac, 'LIMITE_BYTES', 30), self.assertRaises(ValueError):
            ac.salvar(self.sh, self.ctx, {'texto': 'a' * 100})
        self.assertEqual(self.sh.abas, {})

    def test_foto_do_rascunho_nao_escapa_da_pasta_nem_sobrescreve_local(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasta = Path(tmp)
            dados = {'fotos_rascunho': {'antes': ['../fora.jpg']}}
            with self.assertRaises(ValueError):
                ac.anexar_fotos_rascunho(dados, pasta)
            (pasta/'rasc_antes_teste.jpg').write_bytes(b'local novo')
            dados = {'fotos_rascunho': {'antes': ['rasc_antes_teste.jpg']},
                     '_fotos_campo': {'rasc_antes_teste.jpg': base64.b64encode(b'antiga').decode()}}
            with self.assertRaises(ValueError):
                ac.restaurar_fotos_rascunho(dados, pasta)
            self.assertEqual((pasta/'rasc_antes_teste.jpg').read_bytes(), b'local novo')

    def test_grafico_cc_calcula_por_leitura_e_preserva_lacunas(self):
        linhas = [dict(data=f'{d:02d}/10/2026', cloro_livre=cl, cloro_total=ct, cloro_combinado='999')
                  for d, cl, ct in [(5,'2','2.3'), (12,'2','2'), (19,'','3'), (26,'3','2')]]
        self.assertEqual([v for _, v in serie_grafico(linhas, 'cloro_combinado')], [.3, 0, None, None])


class FluxoSemDrive(_Homologacao):
    def test_visita_fotos_e_origem_voltam_do_sheets_sem_disco_ou_drive(self):
        visita = self._lancamento_operador()
        visita.update(fotos_antes_ids=[], analise_semanal_bem_star=True,
                      acesso_analise_bem_star={'nome':'Thyago Silveira','autenticado':True,'perfil':'operador'})
        foto = visita['fotos_antes_b64'][0]
        self.assertTrue(self.app['sheets_salvar_lancamento_campo'](visita, EPSILON))
        # A linha principal contém referência pequena, sem descartar o arquivo na aba auxiliar.
        row = next(r for r in self.planilha.worksheet('🔬 Visitas').get_all_values() if len(r)>1 and r[1]==visita['id_visita'])
        self.assertNotIn(foto, row[22])
        self.assertIn('arquivo_campo_ref', json.loads(row[22]))
        shutil.rmtree(self.app['GENERATED_DIR'], ignore_errors=True)
        self.app['GENERATED_DIR'].mkdir()
        recuperada = next(v for v in self.app['situacao_visitas_condominio'](EPSILON)['visitas'] if v['id_visita']==visita['id_visita'])
        self.assertEqual(recuperada['fotos_antes_b64'], [foto])
        self.assertEqual(recuperada['acesso_analise_bem_star'], visita['acesso_analise_bem_star'])
        pdf = self.app['gerar_pdf_relatorio_visita_bem_star'](recuperada, EPSILON)
        self.assertTrue(_imagens_pdf(pdf))

    def test_falha_de_anexo_nao_confirma_visita_sem_fotos(self):
        antes = self.planilha.worksheet('🔬 Visitas').get_all_values()
        with patch.object(ac, 'salvar', side_effect=OSError('cópia não confirmada')):
            self.assertFalse(self.app['sheets_salvar_lancamento_campo'](self._lancamento_operador(), EPSILON))
        self.assertEqual(self.planilha.worksheet('🔬 Visitas').get_all_values(), antes)
        incompleta = self._lancamento_operador(); incompleta['fotos_antes_b64'] = []
        self.assertFalse(self.app['sheets_salvar_lancamento_campo'](incompleta, EPSILON))
        self.assertEqual(self.planilha.worksheet('🔬 Visitas').get_all_values(), antes)

    def test_rascunho_grande_e_fotos_restaurados_apos_perda_do_disco(self):
        with tempfile.TemporaryDirectory() as tmp:
            ns = _sessao(pasta=Path(tmp))
            ns.update(_rascunho_usa_sheets=lambda: True, conectar_sheets=lambda:self.planilha,
                      persistir_pasta_condominio=lambda *a, **kw:None)
            pasta = Path(tmp)/slugify_nome(EPSILON)/'fotos_rascunho'; pasta.mkdir(parents=True)
            nome = 'rasc_antes_teste.jpg'; (pasta/nome).write_bytes(self.foto)
            dados = {'operador':'Thyago Silveira','obs':'água e medições\n'*7000,
                     'assinatura_responsavel_b64':'ASSINATURA_FICTICIA', 'fotos_rascunho':{'antes':[nome]}}
            self.assertTrue(ns['salvar_rascunho_operador_real'](EPSILON, dados, salvar_sheets=True))
            referencia_anterior = deepcopy(self.planilha.worksheet('_Rascunhos').get_all_values())
            with patch.object(ac, 'salvar', side_effect=OSError('rede indisponível')):
                self.assertFalse(ns['salvar_rascunho_operador_real'](EPSILON, dict(dados,obs='outra'*20000), salvar_sheets=True))
            self.assertEqual(self.planilha.worksheet('_Rascunhos').get_all_values(), referencia_anterior)
            shutil.rmtree(pasta.parent)
            recuperado = ns['carregar_rascunho_operador'](EPSILON)
            self.assertEqual(recuperado['obs'], dados['obs'])
            self.assertEqual(recuperado['assinatura_responsavel_b64'], dados['assinatura_responsavel_b64'])
            self.assertEqual((pasta/nome).read_bytes(), self.foto)


if __name__ == '__main__':
    unittest.main()
