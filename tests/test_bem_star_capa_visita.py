"""Capa Bem Star: seleção, rascunho, salvamento real e PDF após perder o disco.

Dados e imagens fictícios; nenhuma credencial, rede ou base compartilhada.
"""
import ast
import base64
import hashlib
import io
import json
import re
import tempfile
import textwrap
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import fitz
from PIL import Image
from streamlit.testing.v1 import AppTest

import test_operador_fotos_heic as heic
from test_operador_fotos_heic import Upload, _foto
from test_operador_formulario import _sessao
from test_visitas_pdf_marcas import ambiente
from test_homologacao_pr14 import _Homologacao, EPSILON
from utils.bem_star_visita_capa import opcoes_capa_depois, foto_capa_visita
from utils.formatacao import limpar_nome_arquivo, slugify_nome

ROOT=Path(__file__).resolve().parents[1]
SRC=(ROOT/'app.py').read_text(encoding='utf-8')
TREE=ast.parse(SRC)


def fonte(nome):
    return ast.get_source_segment(SRC,next(n for n in TREE.body
        if isinstance(n,ast.FunctionDef) and n.name==nome))


def hash_foto(data):
    return hashlib.sha256(data).hexdigest()


def cor_capa(pdf):
    d=fitz.open(stream=pdf,filetype='pdf')
    # Amostra do centro da foto, longe dos títulos e sem incluir o logo.
    pix=d[0].get_pixmap(clip=fitz.Rect(200,430,300,530),alpha=False)
    return Image.frombytes('RGB',(pix.width,pix.height),pix.samples).resize((1,1)).getpixel((0,0))


class EscolhaCapa(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        heic.FotosHEICOperador.setUpClass()
        cls.carregar=staticmethod(heic.FotosHEICOperador.carregar)
        cls.a=_foto('JPEG',cor=(30,155,65),tamanho=(80,60))
        cls.b=_foto('HEIF',cor=(25,70,190),tamanho=(96,64))

    def test_opcoes_validas_deduplicadas_sem_mudar_upload(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'rasc_depois_a.jpg';p.write_bytes(self.a)
            u=Upload('a.jpg',self.a);u.seek(10)
            op=opcoes_capa_depois([Upload('ruim.heic',b'ruim'),u,Upload('b.heic',self.b)],[p],self.carregar)
            self.assertEqual([x['chave'] for x in op],[hash_foto(self.a),hash_foto(self.b)])
            self.assertEqual(u.tell(),10)
            self.assertEqual(p.read_bytes(),self.a)

    def test_seletor_segunda_foto_rerun_restauracao_e_remocao(self):
        code='''
import streamlit as st
from utils.bem_star_visita_capa import renderizar_seletor_capa, _abrir_imagem
def carregar(f):
    return _abrir_imagem(f.getvalue())
def salvar():
    st.session_state['salvo'] = st.session_state.get('op_capa_depois_hash', '')
renderizar_seletor_capa(st, st.session_state['fotos'], [], carregar, salvar)
st.text_input('Observação', key='obs')
'''
        at=AppTest.from_string(code,default_timeout=30)
        at.session_state['fotos']=[Upload('a.jpg',self.a),Upload('b.heic',self.b)]
        at.run();self.assertFalse(at.exception)
        self.assertEqual(at.session_state['salvo'],hash_foto(self.a))
        at.selectbox[0].set_value(hash_foto(self.b)).run()
        self.assertFalse(at.exception)
        self.assertEqual(at.session_state['salvo'],hash_foto(self.b))
        at.text_input[0].set_value('Manter').run()
        self.assertEqual(at.selectbox[0].value,hash_foto(self.b))
        at2=AppTest.from_string(code,default_timeout=30)
        at2.session_state['fotos']=[Upload('a.jpg',self.a),Upload('b.heic',self.b)]
        at2.session_state['op_capa_depois_hash']=at.session_state['salvo']
        at2.run();self.assertEqual(at2.selectbox[0].value,hash_foto(self.b))
        at2.session_state['fotos']=[Upload('a.jpg',self.a)]
        at2.run();self.assertFalse(at2.exception)
        self.assertEqual(at2.session_state['salvo'],hash_foto(self.a))
        at2.session_state['fotos']=[]
        at2.run();self.assertFalse(at2.exception)
        self.assertEqual(at2.session_state['salvo'],'')
        self.assertEqual(len(at2.image),0)

    def test_rascunho_guarda_escolha_e_troca_isola_condominio(self):
        a=_sessao();a['ss']['op_capa_depois_hash']=hash_foto(self.b)
        d=json.loads(json.dumps(a['coletar_rascunho_operador']('Alfa',[])))
        b=_sessao();b['restaurar_rascunho_operador'](d)
        self.assertEqual(b['ss']['op_capa_depois_hash'],hash_foto(self.b))
        self.assertTrue(b['trocar_condominio_operador']('Alfa','Beta'))
        self.assertEqual(b['ss']['op_capa_depois_hash'],'')
        self.assertEqual(b['salvos'][-1][1]['capa_depois_hash'],hash_foto(self.b))
        b['restaurar_rascunho_operador'](d)
        self.assertEqual(b['ss']['op_capa_depois_hash'],hash_foto(self.b))

    def test_salvamento_real_de_upload_e_rascunho_preserva_referencia(self):
        # Executa o trecho real que salva todas as categorias, até a assinatura.
        start=SRC.index('                pasta_op = GENERATED_DIR / slugify_nome(op_nome_cond.strip())',SRC.index('def _salvar_foto_imediato'))
        end=SRC.index('                assinatura_responsavel_b64 = ',start)
        codigo=textwrap.dedent(SRC[start:end])
        for como_upload in (True,False):
            with self.subTest(upload=como_upload),tempfile.TemporaryDirectory() as tmp:
                pasta=Path(tmp); ss={'op_capa_depois_hash':hash_foto(self.b)}
                uploads=[Upload('a.jpg',self.a)]
                if como_upload:uploads.append(Upload('b.heic',self.b))
                else:
                    r=pasta/'Alfa'/'fotos_rascunho';r.mkdir(parents=True)
                    (r/'rasc_depois_b.heic').write_bytes(self.b)
                remoto={}
                def subir(**kw):
                    fid='id_'+str(len(remoto));remoto[fid]=kw['arquivo_bytes'];return fid
                ns={'st':SimpleNamespace(session_state=ss),'GENERATED_DIR':pasta,'slugify_nome':slugify_nome,
                    'limpar_nome_arquivo':limpar_nome_arquivo,'op_nome_cond':'Alfa','datetime':datetime,
                    '_op_is_rt':False,'Path':Path,'re':re,'drive_upload_foto':subir,
                    '_chave_fotos_operador':lambda c:'op_fotos_'+c,'op_fotos_antes':[],
                    'op_fotos_depois':uploads,'op_fotos_cmaq':[],
                    '_excluir_copia_persistente':lambda *a:None}
                exec(fonte('deduplicar_fotos'),ns)
                exec(codigo,ns)
                salvo=ns['_capa_depois_salva']
                self.assertTrue(salvo['referencia']['arquivo'].endswith('b.heic'))
                self.assertEqual(remoto[salvo['referencia']['drive_id']],self.b)
                self.assertEqual(len(ns['fotos_depois_nomes']),2)
                with Image.open(io.BytesIO(base64.b64decode(salvo['b64']))) as im:
                    self.assertEqual(im.format,'JPEG')

    def test_sem_depois_e_referencia_de_outra_categoria_nao_entram(self):
        b64=base64.b64encode(self.a).decode()
        for v in ({'fotos_cmaq_b64':[b64]},
                  {'capa_depois':{'drive_id':'cmaq'},'fotos_cmaq_ids':['cmaq'],'capa_depois_b64':b64},
                  {'capa_depois':{'arquivo':'../outra.jpg'},'fotos_depois':['../outra.jpg']}):
            self.assertIsNone(foto_capa_visita(v,[],lambda _:self.a,Path('/nao_existe')))

    def test_escolha_indisponivel_nao_e_trocada_silenciosamente(self):
        v={'capa_depois':{'drive_id':'segunda'},'fotos_depois_ids':['primeira','segunda']}
        self.assertIsNone(foto_capa_visita(v,[base64.b64encode(self.a).decode()],lambda _:None,Path('/nao_existe')))


class CapaPDF(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns=ambiente();exec(fonte('_imagem_foto_pdf_em_memoria'),cls.ns)
        cls.a=_foto('JPEG',cor=(25,160,60),tamanho=(96,64))
        cls.b=_foto('JPEG',cor=(30,75,195),tamanho=(120,80))

    def gerar(self,v,nome='Condomínio Fictício'):
        return self.ns['gerar_pdf_relatorio_visita_bem_star'](v,nome)

    def test_primeira_depois_valida_capa_e_categoria_preservadas(self):
        a,b=[base64.b64encode(x).decode() for x in (self.a,self.b)]
        pdf=self.gerar({'data':'14/11/2026','operador':'Fictício',
                       'fotos_depois_b64':['aWxlZ2l2ZWw=',a,b],'fotos_cmaq_b64':[b]})
        cor=cor_capa(pdf);self.assertGreater(cor[1],cor[2]+50)
        doc=fitz.open(stream=pdf,filetype='pdf');texto='\n'.join(p.get_text() for p in doc)
        self.assertIn('14 de novembro de 2026',doc[0].get_text())
        self.assertIn('Apos o tratamento',texto)
        self.assertIn('Casa de maquinas',texto)
        self.assertIn('332/2025',texto);self.assertIn('345/2026',texto)
        self.assertIn('item 4.5',texto)
        for proibido in ['contrate a','necessario contratar','documentacao exigida','332/2026']:
            self.assertNotIn(proibido,texto.lower())
        self.assertTrue(any(len(p.get_images())>=2 for p in list(doc)[1:]))

    def test_sem_foto_e_nome_longo_sem_corte(self):
        nome='Condomínio Fictício das Águas Claras e Jardins de Novembro - Residencial Norte'
        pdf=self.gerar({'data':'14/11/2026','operador':'Operador Fictício de Nome Comprido da Silva'},nome)
        doc=fitz.open(stream=pdf,filetype='pdf');capa=doc[0]
        self.assertIn('Residencial Norte',capa.get_text())
        self.assertIn('Limpeza e manutenção',capa.get_text())
        for b in capa.get_text('blocks'):
            self.assertGreaterEqual(b[0],0);self.assertLessEqual(b[2],capa.rect.width)
            self.assertGreaterEqual(b[1],0);self.assertLessEqual(b[3],capa.rect.height)


class CapaDepoisSheets(_Homologacao):
    def test_segunda_foto_e_preservada_no_sheets_e_pdf_sem_disco(self):
        from utils.bem_star_visita_capa import registrar_escolha_capa
        a=_foto('JPEG',cor=(25,160,60),tamanho=(96,64))
        b=_foto('JPEG',cor=(30,75,195),tamanho=(120,80))
        ids=[self.app['drive_upload_foto'](x,n,EPSILON,'2026-11') for x,n in [(a,'a.jpg'),(b,'b.jpg')]]
        v=self._lancamento_operador()
        v.update(fotos_depois=['a.jpg','b.jpg'],fotos_depois_ids=ids,
                 fotos_depois_b64=[base64.b64encode(x).decode() for x in (a,b)])
        escolha={};registrar_escolha_capa(escolha,hash_foto(b),'depois',b,'b.jpg',ids[1],v['fotos_depois_b64'][1])
        v.update(capa_depois=escolha['referencia'],capa_depois_b64=escolha['b64'])
        self.assertTrue(self.app['sheets_salvar_lancamento_campo'](v,EPSILON))
        recuperada=next(x for x in self.app['situacao_visitas_condominio'](EPSILON)['visitas'] if x['id_visita']==v['id_visita'])
        self.assertEqual(recuperada['capa_depois']['drive_id'],ids[1])
        self.assertNotIn('capa_depois_b64',recuperada)
        self.assertFalse((self.gerados/slugify_nome(EPSILON)/'fotos_campo').exists())
        pdf=self.app['gerar_pdf_relatorio_visita_bem_star'](recuperada,EPSILON)
        cor=cor_capa(pdf);self.assertGreater(cor[2],cor[1]+50)


if __name__=='__main__':
    unittest.main()
