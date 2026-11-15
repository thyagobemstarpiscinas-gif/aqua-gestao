"""Cobertura semanal, preenchimento e consulta do mensal com leituras fictícias."""
from copy import deepcopy
from datetime import date
import unittest

from streamlit.testing.v1 import AppTest
from utils.conferencia_mensal_bem_star import conferir_analises

POOLS = ['Piscina Adulto', 'Piscina Infantil']


def visita(dia, pools=POOLS, **extras):
    return dict(data=f'{dia:02d}/10/2026', id_visita=f'V{dia}', operador='Thyago Silveira',
                analise_semanal_bem_star=True, piscinas=[dict(nome=nome, ph='7,4',
                cloro_livre='1,5', cloro_total='1,8', alcalinidade='100', dureza='200', cianurico='30')
                for nome in pools], **extras)


class Conferencia(unittest.TestCase):
    def conferir(self, ls, piscinas=POOLS, hoje=date(2026, 11, 15)):
        return conferir_analises(ls, piscinas, '10', '2026', hoje)

    def test_quatro_datas_adulto_tres_infantil_com_semana_exata_ausente(self):
        ls = [visita(d) for d in [5, 13, 19]] + [visita(27, pools=POOLS[:1])]
        antes = deepcopy(ls)
        r = self.conferir(ls)
        self.assertEqual([p['Datas com análise'] for p in r['piscinas']], [4, 3])
        self.assertEqual(r['piscinas'][1]['Semanas encerradas sem leitura'], '26/10 a 31/10')
        self.assertEqual(r['piscinas'][0]['Semanas encerradas sem leitura'], 'Nenhuma')
        self.assertIn('13/10/2026', r['piscinas'][0]['Datas disponíveis'])
        self.assertTrue(r['tem_pendencias'])
        self.assertEqual(ls, antes)

    def test_semana_em_andamento_futura_e_primeira_parcial_nao_sao_faltas(self):
        r = self.conferir([], hoje=date(2026, 10, 8))
        self.assertEqual(r['semanas_encerradas'], 0)
        self.assertEqual(r['semanas_em_aberto'], 4)
        self.assertFalse(r['tem_pendencias'])
        self.assertFalse(self.conferir([], hoje=date(2026, 9, 20))['tem_pendencias'])
        self.assertEqual(self.conferir([], hoje=date(2026, 10, 12))['piscinas'][0]['semanas_faltantes'], 1)

    def test_mes_fechado_cinco_segundas_inclui_semana_final_parcial(self):
        r = conferir_analises([], POOLS, '11', '2026', date(2026, 12, 1))
        self.assertEqual(r['semanas_encerradas'], 5)
        self.assertIn('30/11 a 30/11', r['piscinas'][0]['Semanas encerradas sem leitura'])

    def test_deduplicacao_mes_autoria_e_marcacao_iguais_ao_relatorio(self):
        boa = visita(13)
        fora = dict(visita(12), data='12/09/2026')
        outro = dict(visita(19), operador='Outro operador')
        desmarcada = dict(visita(26), analise_semanal_bem_star=False)
        agenda = dict(visita(5), fonte='Agenda mensal Bem Star')
        r = self.conferir([boa, deepcopy(boa), fora, outro, desmarcada, agenda])
        self.assertEqual([p['Datas com análise'] for p in r['piscinas']], [1, 1])
        self.assertEqual(len(r['leituras']), 2)
        self.assertEqual(len(r['pendencias']), 0)

    def test_campos_ausentes_zero_valido_e_cc_inconsistente_na_mesma_piscina(self):
        v = visita(5)
        v['piscinas'][0].update(cianurico='0', dureza='')
        v['piscinas'][1].update(cloro_total='1,0', alcalinidade='texto')
        r = self.conferir([v])
        self.assertEqual(r['pendencias'][0]['Campos ausentes'], 'Dureza cálcica')
        self.assertEqual(r['pendencias'][1]['Valores inválidos'], 'Alcalinidade')
        self.assertEqual(r['pendencias'][1]['Cloro combinado'], 'Inconsistente')
        self.assertEqual(r['leituras'][0]['Cloro combinado (mg/L)'], '0,3')

    def test_linha_vazia_nao_conta_como_analise_e_piscina_cadastrada_sem_visita_aparece(self):
        v = visita(5, pools=POOLS[:1])
        v['piscinas'] = [{'nome': POOLS[0]}]
        r = self.conferir([v])
        self.assertEqual([p['Datas com análise'] for p in r['piscinas']], [0, 0])
        self.assertEqual([p['semanas_faltantes'] for p in r['piscinas']], [4, 4])
        self.assertEqual(len(r['pendencias']), 1)

    def test_todas_semanas_campos_completos_e_periodo_invalido(self):
        r = self.conferir([visita(d) for d in [5, 13, 19, 27]])
        self.assertFalse(r['tem_pendencias'])
        self.assertEqual(r['semanas_encerradas'], 4)
        self.assertTrue(conferir_analises([], [], '', '2026')['erro'])

    def test_ui_consulta_sem_campos_editaveis_e_aviso_com_datas(self):
        tela = '''
import streamlit as st
from datetime import date
from utils.conferencia_mensal_bem_star import renderizar_conferencia
from test_conferencia_mensal import visita, POOLS
renderizar_conferencia(st, [visita(5, pools=POOLS[:1])], POOLS, '10', '2026', date(2026, 11, 15))
'''
        at = AppTest.from_string(tela).run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(len(at.warning), 1)
        self.assertEqual(len(at.text_input), 0)
        self.assertEqual(len(at.dataframe), 2)
        self.assertEqual(at.dataframe[0].value.iloc[1]['Datas com análise'], 0)
        self.assertEqual(at.dataframe[1].value.iloc[0]['Data'], '05/10/2026')
