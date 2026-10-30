"""Regressões D1: simulações, cópias temporárias e documentos reais; nunca produção."""
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from utils.rt_contexto import campos, limpar, ler_fotos, sincronizar_fotos, materializar_fotos, ART_PADRAO
from test_fotos_isolamento import _carregar, IDENTIDADE

class Foto(io.BytesIO):
    name = 'image.png'

class TestContextoCompleto(unittest.TestCase):
    def contexto(self, salvo=True):
        return _carregar(IDENTIDADE, {'_relatorio_rt_salvar_rascunho': lambda *a: salvo,
                                     '_relatorio_rt_chave_sessao': lambda **k: 'nova'})

    def test_d1_a_falha_nao_limpa_nem_rotaciona_contexto(self):
        ns=self.contexto(False);ss=ns['st'].session_state
        ss.update(rel_analise_ph_0='0',rel_diagnostico='ALFA',rel_art_numero='TESTE123',rel_origem_dados='a')
        antes=dict(ss)
        self.assertFalse(ns['guardar_tabela_rt']('a'))
        for k,v in antes.items():self.assertEqual(ss[k],v)
        self.assertNotIn('_rel_por_identidade',ss)
        self.assertIn('bloqueada',ss['_rt_contexto_erro'])

    def test_d1_b_roundtrip_de_todos_campos_e_zeros(self):
        ns=self.contexto();ss=ns['st'].session_state
        identidade=lambda c:ns['identidade_dados_rt']('aqua_gestao',{'status':'ok','id_cliente':c},c,'09','2026')
        ns['conciliar_identidade_tabela_rt'](identidade('A'))
        esperado={'rel_diagnostico':'DIAG_ALFA','rel_parecer_complemento_rt':'PARECER_ALFA',
                  'rel_art_status':'Emitida','rel_art_numero':'TESTE123','rel_art_inicio':'01/09/2026',
                  'rel_art_fim':'30/09/2026','rel_analise_ph_0':'0','rel_dos_qtd_0':'0',
                  'rel_obs_0':'OBS_ALFA','rel_nbr_11238':'NBR_ALFA','rel_epi_luvas_ca':'CA_ALFA'}
        ss.update(esperado);ns['conciliar_identidade_tabela_rt'](identidade('B'))
        self.assertNotIn('DIAG_ALFA',campos(ss).values());self.assertEqual(ss['rel_art_status'],ART_PADRAO)
        ss['rel_diagnostico']='DIAG_BETA';ns['conciliar_identidade_tabela_rt'](identidade('A'))
        for k,v in esperado.items():self.assertEqual(ss[k],v,k)

    def test_d1_d_fotos_bytes_recuperaveis_por_identidade(self):
        ss={};foto=Foto(b'ficticia');sincronizar_fotos(ss,'aqua|A|09',[foto])
        salvo=campos(ss);limpar(ss);ss.update(salvo)
        self.assertEqual(len(ler_fotos(ss,'aqua|A|09')),1)
        sincronizar_fotos(ss,'aqua|A|09',[])  # novo uploader vazio não apaga foto restaurada
        with tempfile.TemporaryDirectory() as tmp:
            paths=materializar_fotos(ss,'aqua|A|09',tmp)
            self.assertEqual(paths[0].read_bytes(),b'ficticia')
        with self.assertRaises(ValueError):ler_fotos(ss,'bem|A|09')

    def test_mesmo_id_em_empresas_diferentes_nao_compartilha_contexto(self):
        ns=self.contexto();ss=ns['st'].session_state
        ident=lambda empresa:ns['identidade_dados_rt'](empresa,{'status':'ok','id_cliente':'C1'},'Homonimo','09','2026')
        a,b=ident('aqua_gestao'),ident('bem_star')
        ns['conciliar_identidade_tabela_rt'](a);ss['rel_diagnostico']='SO_AQUA'
        sincronizar_fotos(ss,a['chave'],[Foto(b'imagem ficticia')])
        ns['conciliar_identidade_tabela_rt'](b)
        self.assertNotEqual(a['chave'],b['chave']);self.assertEqual(ss['rel_diagnostico'],'')
        self.assertEqual(ler_fotos(ss,b['chave']),[])
        ss['rel_diagnostico']='SO_BEM';ns['conciliar_identidade_tabela_rt'](a)
        self.assertEqual(ss['rel_diagnostico'],'SO_AQUA');self.assertEqual(len(ler_fotos(ss,a['chave'])),1)

    def test_rascunho_estrangeiro_nao_muda_empresa_nem_campos(self):
        ns=_carregar(['_relatorio_rt_aplicar_rascunho'])
        ss=ns['st'].session_state;ss.update(empresa_ativa='bem_star',rel_diagnostico='BEM_ORIGINAL')
        dados={'empresa_ativa':'aqua_gestao','campos':{'rel_diagnostico':'AQUA','rel_origem_dados':'aqua_gestao|id:C1|09/2026'}}
        self.assertFalse(ns['_relatorio_rt_aplicar_rascunho'](dados))
        self.assertEqual(ss['empresa_ativa'],'bem_star');self.assertEqual(ss['rel_diagnostico'],'BEM_ORIGINAL')

    def test_remover_upload_e_nao_duplicar_mesmo_conteudo(self):
        ss={};a,b=Foto(b'a'),Foto(b'b')
        self.assertEqual(len(sincronizar_fotos(ss,'A',[a,a,b])),2)
        self.assertEqual(len(sincronizar_fotos(ss,'A',[b])),1)
        data=json.loads(ss['rel_fotos_preservadas']);data['fotos'][0]['bytes']='dGFtcGVyZWQ='
        ss['rel_fotos_preservadas']=json.dumps(data)
        with self.assertRaises(ValueError):ler_fotos(ss,'A')

from test_relatorio_rt_campos import TestRelatorioRTCamposNoAppReal as Base, AppTest

@unittest.skipIf(AppTest is None,'Streamlit obrigatório para integração')
class TestDocumentosProtegidos(unittest.TestCase):
    setUp = Base.setUp
    _entrar_e_carregar = Base._entrar_e_carregar
    _cenario_d1 = Base._cenario_d1
    _cadastrar = Base._cadastrar
    _visitas_alfa = Base._visitas_alfa
    _mes = Base._mes
    _carregar = Base._carregar
    _analises = Base._analises
    _gerar_desabilitado = Base._gerar_desabilitado

    def test_d1_b_c_d_e_documentos_sem_dados_de_outro_cliente(self):
        import zipfile
        from PIL import Image
        from pypdf import PdfReader
        at=self._cenario_d1()
        at.text_area(key='rel_diagnostico').input('MARCADOR_DIAG_ALFA').run()
        at.selectbox(key='rel_art_status_widget').select('Emitida').run()
        at.text_input(key='rel_art_numero').input('999999991')
        at.text_input(key='rel_art_inicio').input('01/09/2026')
        at.text_input(key='rel_art_fim').input('30/09/2026').run()
        at.text_area(key='rel_obs_0').input('MARCADOR_OBS_ALFA').run()
        buf=io.BytesIO();Image.new('RGB',(13,17),(241,28,17)).save(buf,format='PNG')
        next(w for w in at.file_uploader if w.label=='Upload de fotos do relatório').upload('alfa.png',buf.getvalue(),'image/png').run()
        self._carregar(at,'Residencial Beta Fictício')
        self.assertNotEqual(at.session_state['rel_diagnostico'],'MARCADOR_DIAG_ALFA')
        self.assertEqual(self._analises(at,1), [('22/09/2026','6,8','2,0')])
        next(b for b in at.button if b.label=='🚀 Gerar relatório mensal').click().run()
        paths=[p for p in (self.pasta/'app'/'generated').glob('**/*Beta*RELATORIO_RT.*') if 'PREVIA' not in p.name]
        self.assertEqual({p.suffix for p in paths},{'.docx','.pdf'})
        import os,shutil
        if os.environ.get('RT_TEST_EVIDENCE'):
            out=Path(os.environ['RT_TEST_EVIDENCE']);out.mkdir(parents=True,exist_ok=True)
            for doc in paths:shutil.copy2(doc,out/doc.name)
        for p in paths:
            if p.suffix=='.docx':
                with zipfile.ZipFile(p) as z:
                    import xml.etree.ElementTree as ET
                    text=' '.join(n.text or '' for n in ET.fromstring(z.read('word/document.xml')).iter() if n.tag.endswith('}t'))
                    self.assertFalse(any(Image.open(io.BytesIO(z.read(n))).size==(13,17) for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/')))
            else:text=' '.join(p.extract_text() or '' for p in PdfReader(p).pages)
            self.assertNotIn('MARCADOR_DIAG_ALFA',text);self.assertNotIn('MARCADOR_OBS_ALFA',text)
            self.assertNotIn('999999991',text)
            self.assertNotIn('15/09/2026',text);self.assertIn('22/09/2026',text)
            self.assertIn('formaliza',text)
        self._carregar(at,'Condomínio Alfa Fictício')
        self.assertEqual(at.session_state['rel_diagnostico'],'MARCADOR_DIAG_ALFA')
        self.assertEqual(at.session_state['rel_art_numero'],'999999991')
        self.assertEqual(len(ler_fotos(at.session_state.filtered_state,at.session_state['rel_origem_dados'])),1)
        next(b for b in at.button if b.label=='🚀 Gerar relatório mensal').click().run()
        alfas=[p for p in (self.pasta/'app'/'generated').glob('**/*Alfa*RELATORIO_RT.*') if 'PREVIA' not in p.name]
        self.assertEqual({p.suffix for p in alfas},{'.docx','.pdf'})
        for p in alfas:
            if p.suffix=='.docx':
                with zipfile.ZipFile(p) as z:
                    self.assertTrue(any(Image.open(io.BytesIO(z.read(n))).size==(13,17) for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/')))
                    text=' '.join(n.text or '' for n in ET.fromstring(z.read('word/document.xml')).iter() if n.tag.endswith('}t'))
            else:
                pdf=PdfReader(p);text=' '.join(pg.extract_text() or '' for pg in pdf.pages)
                self.assertTrue(any(im.image.size==(13,17) for pg in pdf.pages for im in pg.images))
            self.assertIn('MARCADOR_DIAG_ALFA',text);self.assertIn('999999991',text)
        # Direct rename must not make a preview containing the old client's photo.
        at.text_input(key='rel_nome_condominio').input('Residencial Beta Fictício').run()
        self.assertFalse(at.exception,at.exception)
        self.assertNotEqual(at.session_state['rel_diagnostico'],'MARCADOR_DIAG_ALFA')
        previas=list((self.pasta/'app'/'generated').glob('**/*Beta*PREVIA*.docx'))
        self.assertTrue(previas)
        for p in previas:
            with zipfile.ZipFile(p) as z:
                self.assertFalse(any(Image.open(io.BytesIO(z.read(n))).size==(13,17) for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/')))
                text=' '.join(n.text or '' for n in ET.fromstring(z.read('word/document.xml')).iter() if n.tag.endswith('}t'))
                self.assertNotIn('MARCADOR_DIAG_ALFA',text);self.assertNotIn('15/09/2026',text)
        previas_pdf=list((self.pasta/'app'/'generated').glob('**/*Beta*PREVIA*.pdf'))
        self.assertTrue(previas_pdf)
        for p in previas_pdf:
            pdf=PdfReader(p);text=' '.join(pg.extract_text() or '' for pg in pdf.pages)
            self.assertNotIn('MARCADOR_DIAG_ALFA',text);self.assertNotIn('15/09/2026',text)
            self.assertNotIn('999999991',text)
            self.assertFalse(any(im.image.size==(13,17) for pg in pdf.pages for im in pg.images))
        if os.environ.get('RT_TEST_EVIDENCE'):
            for ext,versoes in (('docx',previas),('pdf',previas_pdf)):
                shutil.copy2(max(versoes,key=lambda p:p.stat().st_mtime_ns),out/('Beta_PREVIA_VALIDADA.'+ext))
        # Outage: neither the final button nor direct preview renderer may create files.
        import sys,streamlit as st
        vis=sys.modules['gspread'].PLANILHA.worksheet('🔬 Visitas')
        before={p:p.stat().st_mtime_ns for p in (self.pasta/'app'/'generated').rglob('*PREVIA*.*')}
        with patch.object(vis,'get_all_values',side_effect=ConnectionError('FALHA_FICTICIA')):
            st.cache_data.clear();at.run()
            self.assertTrue(self._gerar_desabilitado(at))
            after={p:p.stat().st_mtime_ns for p in (self.pasta/'app'/'generated').rglob('*PREVIA*.*')}
            self.assertEqual(before,after)

    def test_d1_a_app_bloqueia_troca_com_disco_indisponivel(self):
        import os
        at=self._cenario_d1();original=os.replace
        def falhar(origem,destino,*args,**kwargs):
            if '_rascunho_relatorio_rt' in str(destino) or '_rascunhos_rt_sync' in str(destino):raise OSError('DISCO_FICTICIO_CHEIO')
            return original(origem,destino,*args,**kwargs)
        with patch('utils.rascunho_sync.os.replace',falhar):
            at.text_input(key='rel_analise_cl_0').input('9,87').run()
            self._carregar(at,'Residencial Beta Fictício')
            self.assertEqual(at.session_state['rel_nome_condominio'],'Condomínio Alfa Fictício')
            self.assertEqual(at.session_state['rel_analise_cl_0'],'9,87')
            self.assertTrue(at.error)
        self._carregar(at,'Residencial Beta Fictício')  # armazenamento recuperou
        persistidos=[json.loads(p.read_text()) for p in (self.pasta/'app'/'generated').glob('_rascunho_relatorio_rt*.json')]
        ultima=next(d for d in persistidos if d.get('campos',{}).get('rel_analise_cl_0')=='9,87')
        at2=self._entrar_e_carregar('Residencial Beta Fictício')
        import re
        chave='btn_restaurar_rascunho_rt_'+re.sub(r'\W','_',ultima['_chave'])
        next(b for b in at2.button if b.key==chave).click().run();at2.run()
        self.assertEqual(at2.session_state['rel_analise_cl_0'],'9,87')
        self.assertEqual(at2.session_state['rel_nome_condominio'],'Condomínio Alfa Fictício')


del Base

from test_relatorio_rt_campos import TestTrocaDeCondominioNoAppReal as _OperadorAux

@unittest.skipIf(AppTest is None,'Streamlit obrigatório para integração')
class TestOperadorContexto(unittest.TestCase):
    setUp = _OperadorAux.setUp
    _entrar_operador = _OperadorAux._entrar_operador

    def test_operador_periodo_condominio_foto_e_parecer(self):
        from PIL import Image
        at=self._entrar_operador('rt-ficticio')
        opcoes=[o for o in at.selectbox(key='rt_op_sel_cond').options if o.upper()!='TODOS']
        a,b=opcoes[:2]
        at.selectbox(key='rt_op_sel_cond').select(a).run()
        at.text_area(key='rel_parecer_complemento_rt').input('PARECER_FICTICIO_A').run()
        buf=io.BytesIO();Image.new('RGB',(9,11),'blue').save(buf,format='PNG')
        next(w for w in at.file_uploader if w.label=='Fotos do relatório RT').upload('operador.png',buf.getvalue(),'image/png').run()
        origem=at.session_state['rel_origem_dados']
        at.selectbox(key='rt_op_sel_cond').select(b).run()
        self.assertFalse(at.exception,at.exception)
        self.assertEqual(at.session_state['rel_parecer_complemento_rt'],'')
        self.assertEqual(ler_fotos(at.session_state.filtered_state,at.session_state['rel_origem_dados']),[])
        at.selectbox(key='rt_op_sel_cond').select(a).run()
        self.assertEqual(at.session_state['rel_parecer_complemento_rt'],'PARECER_FICTICIO_A')
        self.assertEqual(len(ler_fotos(at.session_state.filtered_state,origem)),1)
        mes=next(w for w in at.selectbox if w.label=='Mês')
        antes=mes.value;outro='09' if antes!='09' else '10'
        mes.select(outro).run()
        self.assertEqual(at.session_state['rel_parecer_complemento_rt'],'')
        next(w for w in at.selectbox if w.label=='Mês').select(antes).run()
        self.assertEqual(at.session_state['rel_parecer_complemento_rt'],'PARECER_FICTICIO_A')
        self.assertEqual(len(ler_fotos(at.session_state.filtered_state,origem)),1)
        self.assertFalse(at.exception,at.exception)

del _OperadorAux
