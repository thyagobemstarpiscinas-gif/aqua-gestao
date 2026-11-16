"""Transporte SMTP com TLS verificado e confirmação de envio por etapa."""
import smtplib
import ssl


def enviar_mensagem_smtp(mensagem, host, porta, usuario, senha, *, destinatarios=None, timeout=30):
    """Envia uma única vez; uma queda após DATA nunca provoca reenvio automático."""
    smtp = None
    etapa = "configuração"
    try:
        porta = int(porta)
        if not 1 <= porta <= 65535:
            raise ValueError("porta inválida")
        contexto = ssl.create_default_context()
        modo = "TLS direto" if porta == 465 else "STARTTLS"
        etapa = "conexão"
        if porta == 465:
            smtp = smtplib.SMTP_SSL(host, porta, timeout=timeout, context=contexto)
        else:
            smtp = smtplib.SMTP(host, porta, timeout=timeout)
            etapa = "conexão segura"
            smtp.ehlo()
            smtp.starttls(context=contexto)
            smtp.ehlo()
        etapa = "autenticação"
        smtp.login(usuario, senha)
        etapa = "transmissão"
        recusados = smtp.send_message(mensagem, from_addr=usuario, to_addrs=destinatarios)
        if recusados:
            return False, (
                "O servidor aceitou o e-mail para parte dos destinatários, mas recusou "
                f"{len(recusados)} endereço(s). Confira o recebimento antes de repetir o envio."
            )
        return True, "E-mail aceito pelo servidor para envio."
    except smtplib.SMTPAuthenticationError:
        return False, (
            "O servidor recusou a autenticação do e-mail. Confira o usuário e a senha de "
            "aplicativo na configuração privada do sistema. A mensagem não foi transmitida."
        )
    except smtplib.SMTPRecipientsRefused:
        return False, "O servidor recusou os destinatários. Confira os endereços de e-mail."
    except Exception as erro:
        # Não expõe resposta bruta, credenciais, conteúdo ou dados de outros destinatários.
        detalhe = f" ({modo}, porta {porta})" if etapa != "configuração" else ""
        if etapa == "transmissão":
            return False, (
                f"Não foi possível confirmar o envio{detalhe}: {type(erro).__name__}. "
                "Confira a caixa de entrada e a pasta Enviados antes de tentar novamente, "
                "pois a mensagem pode ter sido aceita pelo servidor."
            )
        return False, (
            f"Falha na {etapa} do e-mail{detalhe}: {type(erro).__name__}. "
            "Verifique o servidor, a porta e a disponibilidade da conexão. "
            "A mensagem não foi transmitida."
        )
    finally:
        if smtp is not None:
            # Uma falha ao encerrar a conexão não invalida um DATA já confirmado.
            try:
                smtp.quit()
            except Exception:
                pass
            finally:
                try:
                    smtp.close()
                except Exception:
                    pass
