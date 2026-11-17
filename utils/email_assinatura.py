"""Assinatura ACQUA aprovada no Gmail, compartilhada pelos envios do app."""
from html import escape
from pathlib import Path


EMAIL_ACQUA = "acqua.gestao.rt@gmail.com"
PEDIDO_CONFIRMACAO = (
    "Por gentileza, confirme o recebimento deste e-mail e a abertura do relatório em anexo."
)
LOGO_CID = "assinatura-acqua@aqua-gestao"
LOGO_PATH = Path(__file__).resolve().parents[1] / "assets" / "assinatura_acqua.png"


def mensagem_com_confirmacao(mensagem: str) -> str:
    texto = (mensagem or "").strip()
    if PEDIDO_CONFIRMACAO.casefold() not in texto.casefold():
        texto = (texto + "\n\n" + PEDIDO_CONFIRMACAO).strip()
    return texto


def assinatura_email_aqua_gestao() -> str:
    """Mantém o modelo compacto azul e dourado configurado no Gmail."""
    return f"""
    <br><br>
    <div style="font-family:Arial,sans-serif;color:#17354d;line-height:1.45;font-size:18px;font-weight:bold;letter-spacing:1px;">
      <div style="border-top:2px solid #bd9648;padding-top:10px;">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="line-height:1.4;border-collapse:collapse;">
          <tr>
            <td style="vertical-align:middle;padding:0 14px 0 0;">
              <img src="cid:{LOGO_CID}" alt="ACQUA GESTÃO" width="124" height="124" style="display:block;width:124px;height:124px;border:0;">
            </td>
            <td style="vertical-align:middle;border-left:2px solid #bd9648;padding:2px 0 2px 16px;">
              <div style="letter-spacing:0.6px;">ACQUA GESTÃO</div>
              <div style="font-size:10px;color:#8a6929;margin:1px 0 10px;">CONTROLE TÉCNICO DE PISCINAS</div>
              <div style="font-size:13px;">Thyago Fernando da Silveira</div>
              <div style="font-size:11px;color:#4d5964;">Técnico em Química · Responsável Técnico<br>CRQ-MG · 2ª Região · nº&nbsp;024025748</div>
              <div style="margin-top:5px;"><a href="mailto:{EMAIL_ACQUA}" style="font-size:11px;color:#17354d;text-decoration:none;">{EMAIL_ACQUA}</a></div>
            </td>
          </tr>
        </table>
      </div>
    </div>
    """


def montar_email_html_aqua(mensagem_texto: str) -> str:
    """Preserva o texto digitado sem interpretá-lo como HTML."""
    corpo = escape(mensagem_texto or "").replace("\n", "<br>")
    return f"""<html><body style="margin:0;padding:0;background:#ffffff;">
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#1f2937;">
    {corpo}{assinatura_email_aqua_gestao()}
    </div></body></html>"""


def adicionar_corpo_email_aqua(msg, mensagem_texto: str) -> None:
    """Inclui versões texto/HTML e a logo inline, independente de acesso ao Drive."""
    logo = LOGO_PATH.read_bytes()
    mensagem_texto = mensagem_com_confirmacao(mensagem_texto)
    assinatura_texto = (
        "ACQUA GESTÃO\nCONTROLE TÉCNICO DE PISCINAS\n"
        "Thyago Fernando da Silveira\nTécnico em Química · Responsável Técnico\n"
        f"CRQ-MG · 2ª Região · nº 024025748\n{EMAIL_ACQUA}"
    )
    msg.set_content((mensagem_texto or "") + "\n\n--\n" + assinatura_texto)
    msg.add_alternative(montar_email_html_aqua(mensagem_texto), subtype="html")
    msg.get_payload()[-1].add_related(
        logo, maintype="image", subtype="png", cid=f"<{LOGO_CID}>", disposition="inline"
    )
