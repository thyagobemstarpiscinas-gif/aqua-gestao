import ast
from email.message import EmailMessage
import io
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from streamlit.testing.v1 import AppTest
from utils.email_assinatura import adicionar_corpo_email_aqua
from utils.email_confirmacao import PEDIDO_CONFIRMACAO, mensagem_com_confirmacao
from utils.envio_documentos import preparar_arquivos, executar_envio


class CanaisEnvio(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.pdf = Path(self.tmp.name) / 'Relatório.pdf'
        self.pdf.write_bytes(b'%PDF-arquivo-ficticio')
        self.arquivos = preparar_arquivos([self.pdf], [])

    def test_whatsapp_nao_chama_smtp_e_nao_exige_destinatario_email(self):
        enviar = Mock()
        resultado = executar_envio('WhatsApp', self.arquivos, enviar, destinatario='', assunto='', mensagem='Olá')
        enviar.assert_not_called()
        self.assertIsNone(resultado['email'])
        self.assertTrue(resultado['whatsapp_preparado'])

    def test_ambos_preserva_anexos_e_nao_repete_falha_email(self):
        upload = io.BytesIO(b'%PDF-segundo'); upload.name = '../Relatório.pdf'
        arquivos = preparar_arquivos([self.pdf], [upload])
        caminhos = []
        def enviar(dest, assunto, mensagem, anexos, **kwargs):
            caminhos.extend(anexos)
            self.assertEqual([p.name for p in anexos], ['Relatório.pdf'] * 2)
            self.assertEqual([p.read_bytes() for p in anexos], [b'%PDF-arquivo-ficticio', b'%PDF-segundo'])
            return False, 'Não foi possível confirmar o envio; confira antes de repetir.'
        mock = Mock(side_effect=enviar)
        resultado = executar_envio('Ambos', arquivos, mock, destinatario='teste@example.invalid', assunto='Teste', mensagem='Olá')
        mock.assert_called_once()
        self.assertFalse(resultado['email'][0])
        self.assertTrue(resultado['whatsapp_preparado'])
        self.assertTrue(all(not p.exists() for p in caminhos))

    def test_arquivo_ausente_vazio_e_extensao_invalida_sao_recusados(self):
        self.pdf.unlink()
        with self.assertRaises(ValueError): preparar_arquivos([self.pdf], [])
        with self.assertRaises(ValueError): preparar_arquivos([], [])
        vazio = io.BytesIO(); vazio.name='vazio.pdf'
        with self.assertRaises(ValueError): preparar_arquivos([], [vazio])
        outro = io.BytesIO(b'oi'); outro.name='arquivo.html'
        with self.assertRaises(ValueError): preparar_arquivos([], [outro])

    def test_pedido_confirmacao_aparece_uma_vez_nas_duas_versoes(self):
        for corpo in ['Mensagem', 'Mensagem\n\n' + PEDIDO_CONFIRMACAO]:
            msg=EmailMessage()
            adicionar_corpo_email_aqua(msg, mensagem_com_confirmacao(corpo))
            self.assertEqual(msg.get_body(('plain',)).get_content().count(PEDIDO_CONFIRMACAO), 1)
            self.assertEqual(msg.get_body(('html',)).get_content().count(PEDIDO_CONFIRMACAO), 1)

    def test_coletor_nao_inclui_documentos_de_outro_condominio(self):
        source=Path(__file__).resolve().parents[1] / 'app.py'
        nodes=[n for n in ast.parse(source.read_text()).body if isinstance(n,ast.FunctionDef) and n.name=='_coletar_documentos_email_aqua']
        outra=Path(self.tmp.name)/'Outro'; outra.mkdir()
        errado=outra/'Outro.pdf'; errado.write_bytes(b'%PDF-outro')
        ns=dict(Path=Path,st=SimpleNamespace(session_state={'ultimos_docs_gerados':{'outro':str(errado)}}),DADOS_JSON_NAME='dados.json',MANIFEST_JSON_NAME='manifest.json')
        exec(compile(ast.Module(body=nodes,type_ignores=[]),str(source),'exec'),ns)
        self.assertEqual(ns['_coletar_documentos_email_aqua'](self.pdf.parent,[errado]),[self.pdf])

    def app(self):
        return AppTest.from_string('''
import streamlit as st
from pathlib import Path
from utils.envio_documentos import exibir_envio_documentos
cliente=st.selectbox('Cliente', ['Alfa', 'Beta'])
def enviar(*args, **kwargs):
    st.session_state['envios_teste'] = st.session_state.get('envios_teste', 0) + 1
    return True, 'Teste aceito'
exibir_envio_documentos(st, nome_condominio=cliente, documentos=[Path(''' + repr(str(self.pdf)) + ''')],
    email_cliente=cliente.lower()+'@example.invalid', mensagem_padrao='Teste', key='teste',
    enviar_email=enviar, email_configurado=(True, ''))
''').run()

    def test_interface_reexecucao_nao_reenvia_email_e_troca_cliente_limpa_resultado(self):
        app=self.app()
        self.assertEqual(len(app.exception),0)
        next(b for b in app.button if b.label=='Enviar por e-mail').click().run()
        self.assertEqual(app.session_state['envios_teste'],1)
        self.assertEqual(len(app.success),1)
        app.run()
        self.assertEqual(app.session_state['envios_teste'],1)
        app.selectbox[0].select('Beta').run()
        self.assertEqual(len(app.success),0)
        self.assertEqual(next(t for t in app.text_input if t.label=='Destinatário do e-mail').value,'beta@example.invalid')

    def test_interface_ambos_prepara_compartilhamento_sem_reenvio(self):
        with patch('utils.envio_documentos.exibir_compartilhamento') as compartilhar:
            app=self.app()
            app.radio[0].set_value('Ambos').run()
            next(b for b in app.button if b.label=='Enviar e-mail e preparar WhatsApp').click().run()
            self.assertEqual(len(app.exception),0)
            self.assertEqual(app.session_state['envios_teste'],1)
            self.assertTrue(compartilhar.called)
            self.assertEqual(compartilhar.call_args.args[1][0]['conteudo'],self.pdf.read_bytes())
            app.run()
            self.assertEqual(app.session_state['envios_teste'],1)

    def test_interface_whatsapp_monta_componente_sem_chamar_email(self):
        app=self.app()
        app.radio[0].set_value('WhatsApp').run()
        next(b for b in app.button if b.label=='Preparar para WhatsApp').click().run()
        self.assertEqual(len(app.exception),0)
        self.assertNotIn('envios_teste',app.session_state)
        self.assertTrue(any('documentos preparados' in i.value for i in app.info))


if __name__=='__main__': unittest.main()
