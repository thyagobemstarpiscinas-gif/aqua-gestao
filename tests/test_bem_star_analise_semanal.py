"""Análises semanais e CC do mensal Bem Star; somente dados fictícios isolados."""
import io
import re
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import Mock

import fitz
from PIL import Image

from utils.bem_star_mensal import (expandir, unir_agenda, cloro_combinado,
    texto_cloro_combinado, filtrar_analises_semanais, deduplicar_lancamentos)
from utils.pdf_bem_star_mensal import gerar_pdf, serie_grafico
from test_fotos_isolamento import _carregar

POOLS = ['Piscina Adulto', 'Piscina Infantil']


def visita(data='05/10/2026', operador='Thyago Silveira', vid='VTESTE1'):
    return dict(id_visita=vid, data=data, operador=operador, salvo_em=data+' 09:00:00',
        piscinas=[
            dict(nome=POOLS[0], ph='7.4', cloro_livre='2.0', cloro_total='2.3',
                 alcalinidade='90', dureza='160', cianurico='40'),
            dict(nome=POOLS[1], ph='7.5', cloro_livre='1.8', cloro_total='2.1',
                 alcalinidade='92', dureza='170', cianurico='42')])


class PiscinasEOrigem(unittest.TestCase):
    def test_duas_piscinas_sobrevivem_a_contexto_pdf_e_reexpansao(self):
        local = visita(); local.pop('id_visita')
        sheets = visita()
        linhas = unir_agenda([local, sheets], [], POOLS)
        self.assertEqual(len(linhas), 2)
        for _ in range(3):
            linhas = expandir(linhas, POOLS)
            self.assertEqual([(r['piscina'], r['ph']) for r in linhas],
                             [(POOLS[0], '7.4'), (POOLS[1], '7.5')])
            self.assertTrue(all(r['id_visita'] == 'VTESTE1' for r in linhas))

    def test_copias_por_piscina_nao_duplicam_e_visitas_distintas_permanecem(self):
        a = expandir([visita()])
        b = expandir([visita(vid='VTESTE2')])
        resultado = expandir(a + deepcopy(a) + b)
        self.assertEqual(len(resultado), 4)
        self.assertEqual(sum(r['piscina'] == POOLS[0] for r in resultado), 2)
        self.assertEqual(len(deduplicar_lancamentos([visita(), visita()])), 1)

    def test_seleciona_segunda_e_autoria_exata_sem_herdar_nome_do_formulario(self):
        entradas = [visita(),
            visita('12/10/2026', ' THYAGO  FERNANDO DA SILVEIRA ', 'V2'),
            visita('07/10/2026', 'Thyago Silveira', 'V3'),
            visita('05/10/2026', 'João', 'V4'),
            visita('05/10/2026', '', 'V5'),
            visita('05/10/2026', 'Thyago Silveira Junior', 'V6'),
            visita('inválida', 'Thyago Silveira', 'V7'),
            dict(visita('05/10/2026', vid='V8'), fonte='Agenda mensal Bem Star')]
        original = deepcopy(entradas)
        self.assertEqual([r['id_visita'] for r in filtrar_analises_semanais(entradas)], ['VTESTE1','V2'])
        self.assertEqual(entradas, original)

    def test_contexto_preserva_operacao_mas_nao_usa_edicao_manual_nas_analises(self):
        semana, outro = visita(), visita('07/10/2026', 'João', 'VOUTRA')
        manual = dict(data='05/10/2026', piscina=POOLS[0], ph='9.9',
                      operador='Thyago Silveira', fonte='Agenda mensal Bem Star')
        ns = _carregar(['_coletar_contexto_relatorio_bem_star'], {
            '_carregar_clientes_bem_star_relatorio': lambda: {'Gama TESTE': {'id':'C1','nome':'Gama TESTE'}},
            'carregar_dados_condominio': lambda *a: {},
            'situacao_visitas_condominio': lambda *a: {'status':'ok','visitas':[semana,outro],'bloqueio':'','aviso':''},
            'origem_validada': lambda *a: True,
            'chave_validacao_origem': lambda *a: 'FICTICIO',
            '_filtrar_lancamentos_preview_por_mes': lambda ls,*a: ls,
            '_piscinas_cadastradas_cliente': lambda *a: POOLS,
            '_bs_coletar_lancamentos_agenda_manual': lambda: [manual],
            '_coletar_fotos_bem_star_preview': lambda *a: [],
            'deduplicar_fotos': lambda ls: ls,
        })
        ns['st'].session_state.update(csr_sel_relatorio='Gama TESTE', csr_mes_rel='10',
                                     csr_ano_rel='2026', csr_operador_rel='Nome no formulário')
        ctx = ns['_coletar_contexto_relatorio_bem_star']()
        self.assertTrue(ctx['ok'])
        self.assertEqual(len(ctx['lancamentos']), 4)
        self.assertEqual(ctx['analises_semanais'], [semana])
        self.assertEqual(ctx['analises_semanais'][0]['piscinas'][0]['ph'], '7.4')


class CloroCombinado(unittest.TestCase):
    def test_calculo_zero_ausencia_e_inconsistencia(self):
        casos = [('2,0','2,3',.3,'0,3'), ('0','0',0,'0'), ('2','2',0,'0'),
                 ('','2',None,'Não calculado'), ('1',None,None,'Não calculado'),
                 ('NaN','3',None,'Não calculado'), ('3','2',None,'Inconsistente'),
                 ('-1','2',None,'Inconsistente')]
        for cl,ct,valor,texto in casos:
            with self.subTest(cl=cl,ct=ct):
                row={'cloro_livre':cl,'cloro_total':ct}
                self.assertEqual(cloro_combinado(row),valor)
                self.assertEqual(texto_cloro_combinado(row),texto)

    def test_resumo_automatico_usa_mesma_fonte_e_ordem_cc(self):
        ns = _carregar(['gerar_parecer_tecnico_bem_star'])
        a = expandir([visita()])[0]
        b = dict(a, data='07/10/2026', cloro_livre='99', cloro_total='100')
        texto = ns['gerar_parecer_tecnico_bem_star']([a,b], 'Gama TESTE','10','2026')
        self.assertIn('1 data(s) de análise semanal',texto)
        self.assertIn('Cloro Combinado: 1 medição(ões), mínimo 0,30',texto)
        self.assertLess(texto.index('Cloro total:'),texto.index('Cloro Combinado:'))
        self.assertLess(texto.index('Cloro Combinado:'),texto.index('Alcalinidade:'))
        self.assertNotIn('99,00',texto)


class FotosMensal(unittest.TestCase):
    def preparar(self):
        imagens = {}
        for fid, cor in [('depois1','green'),('depois2','blue'),('cmaq','brown'),('antes','red')]:
            b=io.BytesIO(); Image.new('RGB',(80,60),cor).save(b,format='PNG'); imagens[fid]=b.getvalue()
        baixar=Mock(side_effect=lambda fid: imagens[fid])
        ns = _carregar(['_materializar_fotos_remotas_bem_star'], {'drive_baixar_foto':baixar})
        return ns,baixar

    def test_ids_especificos_preservam_categorias_sem_duplicar(self):
        ns,baixar=self.preparar()
        v=dict(visita(), fotos_depois_ids=['depois1','depois2'],
               fotos_cmaq_ids=['cmaq'], fotos_drive_ids=['depois1','depois2','cmaq'])
        paths=ns['_materializar_fotos_remotas_bem_star']('Gama TESTE',[v])
        meta=ns['st'].session_state['_bs_fotos_metadados_visitas']
        self.assertEqual(len(paths),3);self.assertEqual(baixar.call_count,3)
        self.assertEqual([m['legenda'] for m in meta.values()],
                         ['Visita VTESTE1 — Após o tratamento']*2+['Visita VTESTE1 — Casa de máquinas'])

    def test_visita_antiga_so_lista_geral_preserva_compatibilidade(self):
        ns,baixar=self.preparar()
        paths=ns['_materializar_fotos_remotas_bem_star']('Gama TESTE',
                    [dict(visita(),fotos_drive_ids=['antes','antes'])])
        self.assertEqual(len(paths),1);self.assertEqual(baixar.call_count,1)
        self.assertEqual(next(iter(ns['st'].session_state['_bs_fotos_metadados_visitas'].values()))['legenda'],
                         'Visita VTESTE1 — Antes do tratamento')


class MensalPDFSemanal(unittest.TestCase):
    def gerar(self, visitas, analises=None):
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup)
        ctx=dict(cliente='Gama FICTÍCIO',mes='10',ano='2026',lancamentos=unir_agenda(visitas,[],POOLS),
                 piscinas=POOLS, piscinas_cadastradas=POOLS)
        if analises is not None:ctx['analises_semanais']=analises
        path=Path(tmp.name)/'mensal.pdf'
        ok,erro=gerar_pdf(path,ctx,cnpj='FICTÍCIO',responsavel='FICTÍCIO',logo_path=None)
        self.assertTrue(ok,erro)
        doc=fitz.open(path);paginas=[p.get_text() for p in doc];doc.close()
        return paginas

    def test_pdf_duas_piscinas_cc_na_ordem_e_estatisticas_somente_de_segundas(self):
        a=visita()
        b=visita('12/10/2026',vid='V2')
        b['piscinas'][0]['cloro_total']='2.7'
        c=visita('19/10/2026',vid='V3')
        c['piscinas'][0]['cloro_total']=''
        d=visita('26/10/2026',vid='V4')
        d['piscinas'][0]['cloro_total']='1.0'
        outro=visita('07/10/2026','João','V5')
        outro['observacao']='SERVIÇO DE OUTRO DIA PRESERVADO'
        outro['piscinas'][0]['ph']='9.9'
        paginas=self.gerar([a,b,c,d,outro])
        adulto=next(p for p in paginas if p.startswith('BEM STAR') and 'Resumo numérico da piscina' in p and 'Piscina Adulto' in p)
        infantil=next(p for p in paginas if 'Resumo numérico da piscina' in p and 'Piscina Infantil' in p)
        for pagina in [adulto,infantil]:
            self.assertRegex(pagina,r'Data\s+pH\s+CL\s+CT\s+CC\s+Alc\.')
            self.assertIn('05/10/2026',pagina)
        self.assertIn('7,4',adulto);self.assertIn('7,5',infantil)
        self.assertIn('Não calculado',' '.join(adulto.split()));self.assertIn('Inconsistente',adulto)
        stats=adulto.split('Resumo numérico da piscina',1)[1]
        self.assertRegex(stats,r'Cloro Combinado\s+0,3\s+0,5\s+0,7\s+2\s+Alcalinidade')
        self.assertNotIn('9,9',adulto)
        self.assertIn('SERVIÇO DE OUTRO DIA PRESERVADO','\n'.join(paginas))

    def test_sem_analise_semanal_nao_substitui_por_outro_operador(self):
        paginas=self.gerar([visita(operador='João')])
        texto='\n'.join(paginas)
        self.assertIn('Sem análise semanal',' '.join(texto.split()))
        self.assertNotIn('7,4',texto);self.assertNotIn('7,5',texto)

    def test_fonte_salva_prevalece_sobre_edicoes_de_parametros_no_formulario(self):
        original=visita(); editada=visita()
        editada['piscinas'][0]['ph']='9.9'
        paginas=self.gerar([editada],analises=[original])
        self.assertIn('7,4','\n'.join(paginas))
        self.assertNotIn('9,9','\n'.join(paginas))


if __name__ == '__main__':
    unittest.main()
