"""Envio isolado: nenhum e-mail real ou credencial de produção."""
import ast
from email.message import EmailMessage
from email import policy
from email.parser import BytesParser
from pathlib import Path
import smtplib
import ssl
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from utils.email_smtp import enviar_mensagem_smtp
from utils.email_assinatura import EMAIL_ACQUA, LOGO_CID, LOGO_PATH


class TransporteSMTP(unittest.TestCase):
    def setUp(self):
        self.smtp = Mock()
        self.smtp.send_message.return_value = {}
        self.msg = EmailMessage()
        self.msg['From'] = 'origem@example.invalid'
        self.msg['To'] = 'destino@example.invalid'
        self.msg.set_content('Teste fictício')
        self.args = (self.msg, 'smtp.example.invalid', 587, 'origem@example.invalid', 'SEGREDO-TESTE')
        self.normal = patch('utils.email_smtp.smtplib.SMTP', return_value=self.smtp).start()
        self.direto = patch('utils.email_smtp.smtplib.SMTP_SSL', return_value=self.smtp).start()
        self.addCleanup(patch.stopall)

    def test_587_exige_tls_verificado_antes_de_autenticar(self):
        self.assertTrue(enviar_mensagem_smtp(*self.args)[0])
        contexto = self.smtp.starttls.call_args.kwargs['context']
        self.assertTrue(contexto.check_hostname)
        self.assertEqual(contexto.verify_mode, ssl.CERT_REQUIRED)
        nomes = [c[0] for c in self.smtp.method_calls]
        self.assertLess(nomes.index('starttls'), nomes.index('login'))
        self.assertLess(nomes.index('login'), nomes.index('send_message'))
        self.direto.assert_not_called()

    def test_465_usa_tls_desde_inicio_sem_starttls(self):
        args = list(self.args); args[2] = 465
        self.assertTrue(enviar_mensagem_smtp(*args)[0])
        contexto = self.direto.call_args.kwargs['context']
        self.assertTrue(contexto.check_hostname)
        self.assertEqual(contexto.verify_mode, ssl.CERT_REQUIRED)
        self.normal.assert_not_called()
        self.smtp.starttls.assert_not_called()
        self.smtp.send_message.assert_called_once()

    def test_falha_tls_bloqueia_login_e_nao_expoe_resposta(self):
        self.smtp.starttls.side_effect = ssl.SSLError('SEGREDO-TESTE')
        ok, detalhe = enviar_mensagem_smtp(*self.args)
        self.assertFalse(ok)
        self.assertIn('conexão segura', detalhe)
        self.assertNotIn('SEGREDO-TESTE', detalhe)
        self.smtp.login.assert_not_called()
        self.smtp.send_message.assert_not_called()

    def test_queda_na_conexao_identifica_porta_sem_tentar_enviar(self):
        self.normal.side_effect = smtplib.SMTPServerDisconnected('SEGREDO-TESTE')
        ok, detalhe = enviar_mensagem_smtp(*self.args)
        self.assertFalse(ok)
        self.assertIn('porta 587', detalhe)
        self.assertIn('não foi transmitida', detalhe)
        self.assertNotIn('SEGREDO-TESTE', detalhe)
        self.smtp.send_message.assert_not_called()

    def test_autenticacao_recusada_nao_transmite_mensagem(self):
        self.smtp.login.side_effect = smtplib.SMTPAuthenticationError(535, b'SEGREDO-TESTE')
        ok, detalhe = enviar_mensagem_smtp(*self.args)
        self.assertFalse(ok)
        self.assertIn('autenticação', detalhe)
        self.assertNotIn('SEGREDO-TESTE', detalhe)
        self.smtp.send_message.assert_not_called()

    def test_queda_durante_envio_nao_repete_e_informa_incerteza(self):
        self.smtp.send_message.side_effect = smtplib.SMTPServerDisconnected('fechado')
        ok, detalhe = enviar_mensagem_smtp(*self.args)
        self.assertFalse(ok)
        self.assertIn('pode ter sido aceita', detalhe)
        self.smtp.send_message.assert_called_once()
        self.normal.assert_called_once()

    def test_falha_no_quit_nao_invalida_envio_aceito(self):
        self.smtp.quit.side_effect = smtplib.SMTPResponseException(500, b'quit recusado')
        self.assertTrue(enviar_mensagem_smtp(*self.args)[0])
        self.smtp.send_message.assert_called_once()
        self.smtp.close.assert_called_once()

    def test_recusa_parcial_nao_anuncia_sucesso_total_nem_reenvia(self):
        self.smtp.send_message.return_value = {'recusado@example.invalid': (550, b'nao')}
        ok, detalhe = enviar_mensagem_smtp(*self.args, destinatarios=['a@example.invalid', 'recusado@example.invalid'])
        self.assertFalse(ok)
        self.assertIn('parte dos destinatários', detalhe)
        self.assertIn('1 endereço', detalhe)
        self.assertNotIn('recusado@example.invalid', detalhe)
        self.smtp.send_message.assert_called_once()

    def test_todos_recusados_tem_mensagem_clara(self):
        self.smtp.send_message.side_effect = smtplib.SMTPRecipientsRefused({'teste@example.invalid': (550, b'no')})
        self.assertIn('recusou os destinatários', enviar_mensagem_smtp(*self.args)[1])

    def test_porta_invalida_nao_conecta(self):
        args = list(self.args); args[2] = 'invalida'
        self.assertFalse(enviar_mensagem_smtp(*args)[0])
        self.normal.assert_not_called()
        self.direto.assert_not_called()


class IntegracaoApp(unittest.TestCase):
    def test_ambos_fluxos_preservam_pdf_e_encaminham_falha(self):
        source = Path(__file__).resolve().parents[1] / 'app.py'
        nomes = {'enviar_email_aqua_smtp', 'enviar_email_smtp_aqua'}
        nodes = [n for n in ast.parse(source.read_text()).body if isinstance(n, ast.FunctionDef) and n.name in nomes]
        cfg = dict(smtp_host='smtp.example.invalid', smtp_port=465, smtp_user='a@example.invalid',
                   smtp_password='SEGREDO-TESTE', remetente_nome='ACQUA TESTE', reply_to=EMAIL_ACQUA)
        ns = dict(Path=Path, st=SimpleNamespace(secrets={'email': cfg}),
                  _email_aqua_configurado=lambda: (True, ''), _email_secrets_aqua=lambda: cfg,
                  email_smtp_configurado=lambda: True, validar_email=lambda x: True)
        exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), 'exec'), ns)
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / 'teste.pdf'; pdf.write_bytes(b'%PDF-TESTE')
            for nome in nomes:
                with self.subTest(fluxo=nome), patch('utils.email_smtp.enviar_mensagem_smtp', return_value=(False, 'Falha controlada')) as transporte:
                    self.assertEqual(ns[nome]('b@example.invalid', 'Teste', 'Mensagem <teste>\nSegunda linha', [pdf], bcc='c@example.invalid'), (False, 'Falha controlada'))
                    args = transporte.call_args.args
                    self.assertEqual(args[2], 465)
                    anexos = list(args[0].iter_attachments())
                    self.assertEqual(len(anexos), 1)
                    self.assertEqual(anexos[0].get_filename(), 'teste.pdf')
                    self.assertEqual(anexos[0].get_payload(decode=True), b'%PDF-TESTE')
                    # Valida o MIME transmitido, incluindo logo ligada ao HTML e PDF separado.
                    recebida = BytesParser(policy=policy.default).parsebytes(args[0].as_bytes())
                    html = recebida.get_body(preferencelist=('html',)).get_content()
                    self.assertIn('Mensagem &lt;teste&gt;<br>Segunda linha', html)
                    self.assertIn('cid:' + LOGO_CID, html)
                    self.assertIn('mailto:' + EMAIL_ACQUA, html)
                    self.assertIn('Thyago Fernando da Silveira', html)
                    self.assertIn('024025748', html)
                    self.assertNotIn('https://', html)
                    self.assertNotIn('mailto:aqua.gestao.rt@gmail.com', html)
                    self.assertNotIn('bemstarpiscinas.com.br', html)
                    self.assertEqual(recebida['Reply-To'], EMAIL_ACQUA)
                    self.assertIn(EMAIL_ACQUA, recebida.get_body(preferencelist=('plain',)).get_content())
                    logos = [p for p in recebida.walk() if p.get_content_type() == 'image/png']
                    self.assertEqual(len(logos), 1)
                    self.assertEqual(logos[0]['Content-ID'], '<' + LOGO_CID + '>')
                    self.assertEqual(logos[0].get_content_disposition(), 'inline')
                    self.assertEqual(logos[0].get_payload(decode=True), LOGO_PATH.read_bytes())


if __name__ == '__main__':
    unittest.main()
