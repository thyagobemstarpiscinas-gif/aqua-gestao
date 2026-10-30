"""Planilha fictícia, disco temporário e rede bloqueada."""
import ast
import copy
import hashlib
import json
import os
from pathlib import Path
import socket
import sys
import tempfile
from types import ModuleType, SimpleNamespace
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from utils import ultimos_sheets as remoto
from utils.rt_contexto import registrar_ultimos as local, caminho_ultimos
class WorksheetNotFound(Exception):pass
class Aba:
    def __init__(self):self.rows=[remoto.CABECALHO.copy()];self.falhar=False;self.resposta_perdida=False
    def get_all_values(self):
        if self.falhar:raise ConnectionError('Falha fictícia')
        return copy.deepcopy(self.rows)
    def append_row(self,row,**kwargs):
        if self.falhar:raise ConnectionError('Falha fictícia')
        self.rows.append(row.copy())
        if self.resposta_perdida:raise ConnectionError('Resposta perdida')
    def update(self,**kwargs):self.rows=copy.deepcopy(kwargs['values'])
class Planilha:
    def __init__(self):self.aba=Aba();self.ausente=False;self.criadas=[]
    def worksheet(self,nome):
        assert nome==remoto.ABA
        if self.ausente:raise WorksheetNotFound(nome)
        return self.aba
    def add_worksheet(self,**kwargs):self.criadas.append(kwargs['title']);self.ausente=False;return self.aba
class TestUltimosSheets(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.raiz=Path(self.tmp.name)
        self.sh=Planilha();self.conectar=lambda:self.sh;self.chave='aqua_gestao|id:C1|09/2026'
        self.snapshot={'parametros_ultimos':[{'ph':'7.3','data':'01/09/2026'}],'dosagens_ultimas':[{'produto':'Cloro','quantidade':'0'}]}
        self.enterContext(patch.dict(os.environ,{'CODESPACES':'false'}))
        self.enterContext(patch.object(socket.socket,'connect',side_effect=AssertionError('Rede proibida')))
        pacote,exceptions=ModuleType('gspread'),ModuleType('gspread.exceptions');exceptions.WorksheetNotFound=WorksheetNotFound
        self.enterContext(patch.dict(sys.modules,{'gspread':pacote,'gspread.exceptions':exceptions}))
    def salvar(self,chave=None,snapshot=None):remoto.registrar_ultimos(self.raiz,chave or self.chave,snapshot or self.snapshot,self.conectar)
    def ler(self,chave=None):return remoto.carregar_ultimos(self.raiz,chave or self.chave,self.conectar)
    def test_recuperacao_sem_disco_entre_periodos(self):
        self.salvar();self.assertFalse(caminho_ultimos(self.raiz,self.chave).exists())
        with tempfile.TemporaryDirectory() as novo:d=remoto.carregar_ultimos(novo,'aqua_gestao|id:C1|10/2026',self.conectar)
        self.assertEqual(d['periodo'],'09/2026');self.assertNotIn('data',d['parametros_ultimos'][0])
    def test_isolamento_por_empresa_e_id(self):
        self.salvar()
        for chave in ['bem_star|id:C1|09/2026','aqua_gestao|id:C2|09/2026']:
            with self.assertRaises(FileNotFoundError):self.ler(chave)
    def test_ultima_geracao_inclusive_mes_anterior_sem_sobrescrita(self):
        self.salvar();antes=copy.deepcopy(self.sh.aba.rows[1]);self.salvar('aqua_gestao|id:C1|08/2026')
        self.assertEqual(self.ler()['periodo'],'08/2026');self.assertEqual(self.sh.aba.rows[1],antes);self.assertEqual(len(self.sh.aba.rows),3)
    def test_falha_consulta_nao_usa_cache_local(self):
        local(self.raiz,self.chave,self.snapshot);self.sh.aba.falhar=True
        with self.assertRaises(ConnectionError):self.ler()
    def test_falha_gravacao_preserva_anterior(self):
        self.salvar();antes=copy.deepcopy(self.sh.aba.rows);self.sh.aba.falhar=True
        with self.assertRaises(ConnectionError):self.salvar()
        self.assertEqual(self.sh.aba.rows,antes)
    def test_resposta_perdida_nao_apaga_ultimo_confirmavel(self):
        self.salvar();self.sh.aba.resposta_perdida=True
        with self.assertRaises(ConnectionError):self.salvar('aqua_gestao|id:C1|08/2026')
        self.assertEqual(self.ler()['periodo'],'08/2026')
    def test_criacao_so_apos_ausencia_explicita(self):
        self.sh.ausente=True
        with self.assertRaises(FileNotFoundError):self.ler()
        self.assertEqual(self.sh.criadas,[]);self.salvar();self.assertEqual(self.sh.criadas,[remoto.ABA])
        self.sh.worksheet=lambda nome:(_ for _ in ()).throw(PermissionError('Sem permissão'))
        with self.assertRaises(PermissionError):self.salvar()
        self.assertEqual(len(self.sh.criadas),1)
    def test_cabecalho_invalido_nao_e_modificado(self):
        self.sh.aba.rows=[['Outro conteúdo']];antes=copy.deepcopy(self.sh.aba.rows)
        with self.assertRaises(ValueError):self.salvar()
        self.assertEqual(self.sh.aba.rows,antes)
    def test_corrupcao_e_identidade_divergente_recusadas(self):
        self.salvar();row=self.sh.aba.rows[-1];row[2]='{}'
        with self.assertRaises(ValueError):self.ler()
        d=remoto.rt_contexto.montar_ultimos(self.chave,self.snapshot);d['empresa']='bem_star'
        row[2]=json.dumps(d);row[3]=hashlib.sha256(row[2].encode()).hexdigest()
        with self.assertRaises(ValueError):self.ler()
    def test_payload_grande_nao_altera_aba(self):
        self.salvar();antes=copy.deepcopy(self.sh.aba.rows);grande=copy.deepcopy(self.snapshot);grande['dosagens_ultimas'][0]['produto']='x'*46000
        with self.assertRaises(ValueError):self.salvar(snapshot=grande)
        self.assertEqual(self.sh.aba.rows,antes)
    def test_desenvolvimento_nao_conecta(self):
        with patch.dict(os.environ,{'CODESPACES':'true'}):
            nao=lambda:(_ for _ in ()).throw(AssertionError('Não conectar'))
            remoto.registrar_ultimos(self.raiz,self.chave,self.snapshot,nao)
            self.assertEqual(remoto.carregar_ultimos(self.raiz,self.chave,nao)['periodo'],'09/2026')
    def test_ambos_aplicadores_reais_preservam_na_falha_e_recuperam(self):
        tree=ast.parse((Path(__file__).resolve().parents[1]/'app.py').read_text())
        nodes=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in {'recuperar_ultimos_valores_rt','aplicar_dosagens_ultimas_no_relatorio'}]
        for tipo in ['parametros','dosagens']:
            ss={'rel_origem_dados':'aqua_gestao|id:C1|10/2026','rel_analises_total':12,'rel_analise_data_0':'15/10/2026','rel_analise_ph_0':'ATUAL','rel_dos_produto_0':'ATUAL'}
            ns={'st':SimpleNamespace(session_state=ss),'GENERATED_DIR':self.raiz,'ANALISES_PADRAO':12,'validar_contexto_rt':lambda:[],'conectar_sheets':self.conectar}
            exec(compile(ast.Module(body=nodes,type_ignores=[]),'app.py','exec'),ns)
            self.sh.aba.falhar=True;antes=copy.deepcopy(ss);ns['recuperar_ultimos_valores_rt'](tipo)
            self.assertIn('Não foi possível',ss.pop('_rt_ultimos_aviso'));self.assertEqual(ss,antes)
            self.sh.aba.falhar=False;self.salvar();ns['recuperar_ultimos_valores_rt'](tipo)
            self.assertEqual(ss['rel_analise_data_0'],'15/10/2026')
            self.assertEqual(ss['rel_analise_ph_0'] if tipo=='parametros' else ss['rel_dos_produto_0'],'7.3' if tipo=='parametros' else 'Cloro')
            self.assertIn('09/2026',ss['_rt_ultimos_aviso'])
if __name__=='__main__':unittest.main(verbosity=2)
