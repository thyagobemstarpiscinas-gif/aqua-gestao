"""Capa mensal: período, escolha da foto e complementos isolados, sem rede."""
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

import fitz
from PIL import Image
from streamlit.testing.v1 import AppTest

from test_bem_star_mensal import A, B, foto, visita
from utils.bem_star_mensal import anexar_foto, salvar_complementos, carregar_complementos, materializar_fotos
from utils.pdf_bem_star_mensal import gerar_pdf


UI = '''
import streamlit as st
from utils.contexto_bem_star import preparar_contexto_bem_star
from utils.bem_star_mensal_ui import renderizar_fotos
nome = st.selectbox('Cliente fictício', ['Condomínio TESTE A', 'Condomínio TESTE B'])
mes = st.selectbox('Mês fictício', ['09', '10'])
preparar_contexto_bem_star(st.session_state, nome, mes, '2026', 'id-a' if nome.endswith('A') else 'id-b')
renderizar_fotos(st)
st.text_input('Observação fictícia', key='csr_obs_rel')
'''


class CapaMensal(unittest.TestCase):
    def setUp(self):
        self.fotos = []
        anexar_foto(self.fotos, foto('green'), 'primeira.png', A)
        anexar_foto(self.fotos, foto('blue'), 'segunda.png', A)

    def gerar(self, pasta, paths, meta, escolhida):
        ctx = dict(cliente=A[1], mes='09', ano='2026', lancamentos=[visita()],
                   fotos=paths, fotos_metadados=meta, foto_capa=escolhida)
        destino = Path(pasta) / 'mensal.pdf'
        ok, erro = gerar_pdf(destino, ctx, cnpj='FICTÍCIO', responsavel='FICTÍCIO', logo_path=None)
        self.assertTrue(ok, erro)
        return fitz.open(destino)

    def test_capa_mensal_usa_segunda_foto_sem_remover_fotos_do_corpo(self):
        with tempfile.TemporaryDirectory() as d:
            paths, meta = materializar_fotos(self.fotos, A, d)
            with self.gerar(d, paths, meta, paths[1]) as pdf:
                capa = ' '.join(pdf[0].get_text().split())
                for texto in ['Relatório Mensal Operacional', 'Setembro de 2026',
                              '01/09/2026 a 30/09/2026', 'REGISTRO FOTOGRÁFICO DO PERÍODO']:
                    self.assertIn(texto, capa)
                for texto in ['DATA DA VISITA', 'REGISTRO DA VISITA', 'APÓS O TRATAMENTO']:
                    self.assertNotIn(texto, capa)
                pix = pdf[0].get_pixmap(clip=fitz.Rect(200, 430, 300, 530), alpha=False)
                r, g, b = Image.frombytes('RGB', (pix.width, pix.height), pix.samples).resize((1, 1)).getpixel((0, 0))
                self.assertGreater(b, r+100)
                self.assertGreater(b, g+100)
                ids_corpo = {i[0] for p in list(pdf)[1:] for i in p.get_images()}
                self.assertEqual(len(ids_corpo), 2)
                cores_corpo = [Image.open(io.BytesIO(pdf.extract_image(i)['image'])).convert('RGB')
                               .resize((1, 1)).getpixel((0, 0)) for i in ids_corpo]
                self.assertTrue(any(b > r+100 and b > g+100 for r, g, b in cores_corpo))
                self.assertTrue(any(g > r+60 and g > b+60 for r, g, b in cores_corpo))

    def test_sem_foto_ou_imagem_escolhida_indisponivel_nao_substitui_por_outra(self):
        with tempfile.TemporaryDirectory() as d:
            paths, meta = materializar_fotos(self.fotos, A, d)
            for escolha in [None, Path(d)/'ausente.jpg']:
                with self.subTest(escolha=escolha), self.gerar(d, paths, meta, escolha) as pdf:
                    self.assertFalse(pdf[0].get_images())
                    self.assertIn('PERÍODO DO RELATÓRIO', pdf[0].get_text())
                    self.assertTrue(any(p.get_images() for p in list(pdf)[1:]))

    def test_seletor_preserva_escolha_por_cliente_mes_e_limpa_foto_removida(self):
        at = AppTest.from_string(UI, default_timeout=30).run()
        at.session_state['bs_mensal_fotos'] = self.fotos
        at.run()
        chave = 'bs_mensal_capa_hash'
        segunda = self.fotos[1]['sha256']
        at.selectbox(key=chave).set_value(segunda).run()
        at.text_input(key='csr_obs_rel').set_value('Preservar observação').run()
        self.assertFalse(at.exception)
        self.assertEqual(at.selectbox(key=chave).value, segunda)
        at.selectbox[0].set_value(B[1]).run()
        self.assertEqual(at.selectbox(key=chave).value, '')
        self.assertEqual(at.selectbox(key=chave).options, ['Sem foto na capa'])
        at.selectbox[0].set_value(A[1]).run()
        self.assertEqual(at.selectbox(key=chave).value, segunda)
        self.assertEqual(at.session_state['csr_obs_rel'], 'Preservar observação')
        at.selectbox[1].set_value('10').run()
        self.assertEqual(at.selectbox(key=chave).value, '')
        at.selectbox[1].set_value('09').run()
        self.assertEqual(at.selectbox(key=chave).value, segunda)
        # Remoção usa o mesmo manifesto do uploader, sem escolher a outra foto.
        botao = next(b for b in at.button if segunda in b.key and b.key.endswith('remover'))
        botao.click().run()
        self.assertFalse(at.exception)
        self.assertEqual(at.selectbox(key=chave).value, '')
        self.assertEqual(len(at.session_state['bs_mensal_fotos']), 1)

    def test_complementos_restauram_escolha_e_foto_em_pasta_vazia(self):
        segunda = self.fotos[1]['sha256']
        ss = {'_bs_rel_contexto': A, 'bs_mensal_fotos': self.fotos,
              'bs_mensal_capa_hash': segunda, 'csr_obs_rel': 'Preservar'}
        aba = Mock()
        salvar_complementos(ss, aba)
        rows = aba.append_rows.call_args.args[0]
        campos = carregar_complementos(A, rows)
        self.assertEqual(campos['bs_mensal_capa_hash'], segunda)
        self.assertEqual(campos['csr_obs_rel'], 'Preservar')
        self.assertIsNone(carregar_complementos(B, rows))
        with tempfile.TemporaryDirectory() as d:
            paths, meta = materializar_fotos(campos['bs_mensal_fotos'], A, d)
            capa = next(p for p in paths if p.stem == campos['bs_mensal_capa_hash'])
            with self.gerar(d, paths, meta, capa) as pdf:
                self.assertEqual(len(pdf[0].get_images()), 1)


if __name__ == '__main__':
    unittest.main()
