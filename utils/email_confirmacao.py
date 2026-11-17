"""Pedido de confirmação visível no corpo dos e-mails enviados pelo app."""
PEDIDO_CONFIRMACAO = (
    "Por gentileza, confirme o recebimento deste e-mail e a abertura do relatório em anexo."
)


def mensagem_com_confirmacao(mensagem: str) -> str:
    texto = (mensagem or "").strip()
    if PEDIDO_CONFIRMACAO.casefold() not in texto.casefold():
        texto = (texto + "\n\n" + PEDIDO_CONFIRMACAO).strip()
    return texto

