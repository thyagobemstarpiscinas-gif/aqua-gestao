"""Relatório mensal com registros fictícios; sem chamadas à planilha ou ao Drive."""
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
from PIL import Image
import fitz
from utils.bem_star_mensal import *
from utils.contexto_bem_star import preparar_contexto_bem_star
from utils.pdf_bem_star_mensal import serie_grafico, gerar_pdf

A = ('id-a', 'Condomínio TESTE A', '09', '2026')
B = ('id-b', 'Condomínio TESTE B', '09', '2026')


def foto(cor='blue'):
    b = io.BytesIO(); Image.new('RGB', (80, 60), cor).save(b, format='PNG'); return b.getvalue()


def visita():
    return {'data':'07/09/2026', 'ph':'9.9', 'fotos':['original.jpg'], 'piscinas':[
        {'nome':'Adulto','tipo':'Adulto','ph':'7.4','cloro_livre':'2.2'},
        {'nome':'Infantil','tipo':'Infantil','ph':'8.3'}]}


class MensalDados(unittest.TestCase):
    def test_expande_todas_sem_herdar_parametro_global_ou_de_outra_piscina(self):
        linhas = expandir([visita()])
        self.assertEqual([(l['piscina'],l['ph'],l['cloro_livre']) for l in linhas], [('Adulto','7.4','2.2'),('Infantil','8.3','')])

    def test_edicao_da_adulta_preserva_infantil_e_referencia_fotografica(self):
        linhas = unir_agenda([visita()], [{'data':'07/09/2026','piscina':'Adulto','ph':'7.2'}])
        self.assertEqual({l['piscina']:l['ph'] for l in linhas}, {'Adulto':'7.2','Infantil':'8.3'})
        self.assertEqual(linhas[0]['fotos'], ['original.jpg'])

    def test_duas_coletas_na_mesma_data_nao_sao_apagadas_pela_agenda(self):
        originais=[{'data':'07/09/2026','piscina':'Adulto','ph':'7.4','id_visita':'V1'},
                   {'data':'07/09/2026','piscina':'Adulto','ph':'7.8','id_visita':'V2'}]
        resultado=unir_agenda(originais,[{'data':'07/09/2026','piscina':'Adulto','ph':'7.2'}])
        self.assertEqual([(l['id_visita'],l['ph']) for l in resultado],[('V1','7.2'),('V2','7.8')])

    def test_coleta_por_piscina_e_mesma_data_sem_colisao(self):
        s = {'bs_mensal_piscinas':['Adulto','Infantil']}
        s[chave_campo('Adulto','07/09/2026','ph')] = '7,4'
        s[chave_campo('Infantil','07/09/2026','ph')] = '8,3'
        linhas = coletar_agenda(s,['07/09/2026'])
        self.assertEqual([l['ph'] for l in linhas], ['7,4','8,3'])

    def test_zero_e_ausencia_sao_distintos_no_grafico(self):
        ls = [{'data':f'{d:02d}/09/2026','cloro_livre':v} for d,v in [(7,'0'),(14,''),(21,'2,1')]]
        self.assertEqual([v for _,v in serie_grafico(ls,'cloro_livre')], [0,None,2.1])
        for v in ['nan','inf','7x','1.2.3']:
            self.assertIsNone(numero(v))

    def test_dose_global_nao_e_atribuida_a_duas_piscinas(self):
        v = visita(); v['dosagens'] = [{'produto':'Produto TESTE'}]
        ls = expandir([v])
        self.assertEqual([l['dosagens'] for l in ls], [[],[]])
        self.assertTrue(ls[0]['dosagens_da_visita'])


class MensalFotos(unittest.TestCase):
    def setUp(self):
        self.fotos=[]; anexar_foto(self.fotos,foto(),'imagem.png',A)

    def test_mesma_imagem_nao_duplica_mesmo_com_outro_nome(self):
        self.assertFalse(anexar_foto(self.fotos,foto(),'outra.png',A))
        self.assertEqual(len(self.fotos),1)

    def test_anexo_de_outro_cliente_bloqueia_geracao(self):
        with tempfile.TemporaryDirectory() as d, self.assertRaises(ValueError):
            materializar_fotos(self.fotos,B,d)

    def test_anexo_de_outro_mes_e_id_distinto_nao_mistura(self):
        for ctx in [(A[0],A[1],'10','2026'),('outro-id',*A[1:])]:
            self.assertNotEqual(contexto_id(A),contexto_id(ctx))

    def test_bytes_adulterados_bloqueiam_sem_usar_cache_local(self):
        self.fotos[0]['bytes'] = b'invalido'
        with tempfile.TemporaryDirectory() as d, self.assertRaises(ValueError):
            materializar_fotos(self.fotos,A,d)

    def test_data_da_imagem_fora_do_mes_bloqueia(self):
        self.fotos[0]['data'] = '01/10/2026'
        with tempfile.TemporaryDirectory() as d, self.assertRaises(ValueError):
            materializar_fotos(self.fotos,A,d)

    def test_troca_contexto_separa_fotos_e_retorno_restaurado(self):
        s={}; preparar_contexto_bem_star(s,A[1],A[2],A[3],A[0]); s['bs_mensal_fotos']=self.fotos
        s['bs_mensal_piscinas']=['Adulto','Infantil']
        preparar_contexto_bem_star(s,B[1],B[2],B[3],B[0])
        self.assertEqual(s['bs_mensal_fotos'],[])
        preparar_contexto_bem_star(s,A[1],A[2],A[3],A[0])
        self.assertEqual(s['bs_mensal_fotos'][0]['sha256'],self.fotos[0]['sha256'])
        self.assertEqual(s['bs_mensal_piscinas'],['Adulto','Infantil'])


class MensalNuvem(unittest.TestCase):
    def estado(self):
        s={'_bs_rel_contexto':A,'csr_obs_rel':'Texto fictício '*7000,'rel_diagnostico':'RT NÃO EXPORTAR'}
        s['bs_mensal_fotos']=[]; anexar_foto(s['bs_mensal_fotos'],foto(),'teste.png',A)
        return s

    def test_salva_dividido_restaura_fotos_sem_drive(self):
        s=self.estado(); aba=Mock(); enviar=Mock(return_value='id-drive-teste')
        salvar_complementos(s,aba,enviar)
        rows=aba.append_rows.call_args.args[0]
        self.assertGreater(len(rows),1)
        self.assertTrue(all(len(r[5]) <= 38000 for r in rows))
        campos=carregar_complementos(A,rows)
        self.assertEqual(campos['csr_obs_rel'],s['csr_obs_rel']); self.assertNotIn('rel_diagnostico',campos)
        self.assertEqual(campos['bs_mensal_fotos'][0]['bytes'],s['bs_mensal_fotos'][0]['bytes'])
        enviar.assert_not_called()
        baixar=Mock(return_value=s['bs_mensal_fotos'][0]['bytes'])
        with tempfile.TemporaryDirectory() as d:
            paths,meta=materializar_fotos(campos['bs_mensal_fotos'],A,d,baixar)
            self.assertEqual(len(paths),1)
            baixar.assert_not_called()
        self.assertIsNone(carregar_complementos(B,rows))

    def test_falha_planilha_preserva_anexos_e_nao_declara_sucesso(self):
        s=self.estado(); aba=Mock(); aba.append_rows.side_effect=RuntimeError('indisponível')
        with self.assertRaises(RuntimeError): salvar_complementos(s,aba)
        self.assertTrue(s['bs_mensal_fotos'][0]['bytes'])

    def test_imagem_salva_adulterada_e_bloqueada(self):
        s=self.estado(); aba=Mock(); salvar_complementos(s,aba)
        rows=aba.append_rows.call_args.args[0]
        payload=''.join(r[5] for r in rows)
        obj=json.loads(payload);obj['campos']['bs_mensal_fotos'][0]['sha256']='0'*64
        payload=json.dumps(obj,ensure_ascii=False);partes=[payload[i:i+38000] for i in range(0,len(payload),38000)]
        alterado=[[rows[0][0],rows[0][1],str(i),str(len(partes)),sha256(payload.encode()).hexdigest(),p] for i,p in enumerate(partes)]
        with self.assertRaises(ValueError): carregar_complementos(A,alterado)

    def test_revisao_parcial_nao_recupera_versao_antiga_silenciosamente(self):
        s=self.estado(); aba=Mock(); salvar_complementos(s,aba,Mock(return_value='fid'))
        rows=aba.append_rows.call_args.args[0]
        with self.assertRaises(ValueError): carregar_complementos(A,rows[:-1])
        rows[0][5]+='alterado'
        with self.assertRaises(ValueError): carregar_complementos(A,rows)


class MensalPDF(unittest.TestCase):
    def test_pdf_com_fotos_tipos_aviso_e_chamada_sem_imposicao(self):
        with tempfile.TemporaryDirectory() as d:
            fs=[]; anexar_foto(fs,foto(),'teste.png',A); fs[0]['piscina']='Infantil'; fs[0]['legenda']='IMAGEM FICTÍCIA PARA TESTE'
            paths,meta=materializar_fotos(fs,A,d)
            ctx={'cliente':A[1],'mes':'09','ano':'2026','lancamentos':[visita()], 'fotos':paths,'fotos_metadados':meta}
            out=Path(d)/'teste.pdf';ok,err=gerar_pdf(out,ctx,cnpj='TESTE',responsavel='Operador TESTE',logo_path=None)
            self.assertTrue(ok,err)
            doc=fitz.open(out); texto='\n'.join(p.get_text() for p in doc)
            for t in ['Infantil','Adulto','332/2025','345/2026','Não é relatório de RT','IMAGEM FICTÍCIA PARA TESTE','não condiciona']:
                self.assertIn(t," ".join(texto.split()))
            self.assertTrue(any(p.get_images() for p in doc)); doc.close()

    def test_valor_invalido_nao_vira_lacuna_silenciosa(self):
        ctx={'cliente':A[1],'mes':'09','ano':'2026','lancamentos':[{'data':'07/09/2026','ph':'7x'}]}
        with tempfile.TemporaryDirectory() as d:
            ok,err=gerar_pdf(Path(d)/'teste.pdf',ctx,cnpj='TESTE',responsavel='TESTE',logo_path=None)
            self.assertFalse(ok); self.assertIn('Valor inválido',err)
