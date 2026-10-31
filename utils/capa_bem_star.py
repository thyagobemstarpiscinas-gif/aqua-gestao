"""Capa do relatório mensal Bem Star, com as cores da marca sobre branco."""
from html import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph


def desenhar_capa_bem_star(canvas, *, cliente, mes_nome, ano, periodo_ini,
                          periodo_fim, data_emissao, logo_path, cnpj):
    largura, altura = A4
    margem = 22 * mm
    util = largura - 2 * margem
    azul = colors.HexColor("#104888")
    agua = colors.HexColor("#0090D0")
    texto = colors.HexColor("#273745")
    cinza = colors.HexColor("#5C6670")

    def paragrafo(conteudo, topo, *, tamanho=11, entrelinha=15,
                  cor=texto, fonte="Helvetica", limite=None):
        estilo = ParagraphStyle("capa", fontName=fonte, fontSize=tamanho,
                                leading=entrelinha, textColor=cor)
        p = Paragraph(conteudo, estilo)
        _, h = p.wrap(util, altura)
        if limite is not None and h > limite:
            raise ValueError("O nome do cliente excede o espaço disponível na capa.")
        p.drawOn(canvas, margem, topo - h)
        return h

    def rotulo(conteudo, y, cor=cinza):
        canvas.saveState()
        objeto = canvas.beginText(margem, y)
        objeto.setFont("Helvetica-Bold", 8)
        objeto.setCharSpace(1.2)
        objeto.setFillColor(cor)
        objeto.textOut(conteudo)
        canvas.drawText(objeto)
        canvas.restoreState()

    canvas.saveState()
    canvas.setFillColor(colors.white)
    canvas.rect(0, 0, largura, altura, fill=1, stroke=0)

    # A marca azul permanece sobre branco para preservar seu contraste.
    marca_desenhada = False
    if logo_path:
        try:
            canvas.drawImage(ImageReader(str(logo_path)), margem, altura - 48 * mm,
                             width=74 * mm, height=24.5 * mm,
                             preserveAspectRatio=True, anchor="sw", mask="auto")
            marca_desenhada = True
        except Exception:
            pass  # A indisponibilidade da imagem não impede a emissão.
    if not marca_desenhada:
        canvas.setFillColor(azul)
        canvas.setFont("Helvetica-Bold", 23)
        canvas.drawString(margem, altura - 37 * mm, "BEM STAR PISCINAS")

    canvas.setFillColor(cinza)
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(largura - margem, altura - 30 * mm, "RELATÓRIO MENSAL")
    canvas.setFillColor(azul)
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawRightString(largura - margem, altura - 37 * mm,
                           f"{mes_nome.upper()} / {ano}")

    canvas.setStrokeColor(colors.HexColor("#DCE7F1"))
    canvas.setLineWidth(0.6)
    canvas.line(margem, altura - 62 * mm, largura - margem, altura - 62 * mm)
    canvas.setStrokeColor(agua)
    canvas.setLineWidth(2.5)
    canvas.line(margem, altura - 62 * mm, margem + 25 * mm, altura - 62 * mm)

    rotulo("ACOMPANHAMENTO OPERACIONAL", altura - 83 * mm, agua)
    paragrafo("Relatório operacional<br/>de manutenção", altura - 95 * mm,
              tamanho=30, entrelinha=35, cor=azul, fonte="Helvetica-Bold")
    paragrafo("Manutenção e acompanhamento de piscinas", altura - 126 * mm,
              tamanho=11, entrelinha=15, cor=cinza)

    rotulo("CLIENTE / LOCAL ATENDIDO", altura - 153 * mm)
    # Nomes extensos mantêm sua identificação completa, sem sobrepor o período.
    nome_seguro = escape(str(cliente))
    for tamanho in range(20, 11, -1):
        p = Paragraph(nome_seguro, ParagraphStyle("nome", fontName="Helvetica-Bold",
                      fontSize=tamanho, leading=tamanho + 4, textColor=texto))
        _, h_nome = p.wrap(util, altura)
        if h_nome <= 30 * mm:
            break
    paragrafo(nome_seguro, altura - 162 * mm, tamanho=tamanho,
              entrelinha=tamanho + 4, fonte="Helvetica-Bold", limite=30 * mm)

    paragrafo(f"{escape(mes_nome)} / {escape(str(ano))}", altura - 201 * mm,
              tamanho=16, entrelinha=20, cor=azul, fonte="Helvetica-Bold")
    paragrafo(f"Período de referência: {escape(periodo_ini)} a {escape(periodo_fim)}",
              altura - 212 * mm, tamanho=9, entrelinha=12, cor=cinza)

    # Rodapé azul com detalhes discretos que remetem à água.
    faixa = 67 * mm
    canvas.setFillColor(azul)
    canvas.rect(0, 0, largura, faixa, fill=1, stroke=0)
    canvas.setFillColor(agua)
    canvas.rect(0, faixa - 2.5, largura, 2.5, fill=1, stroke=0)
    canvas.setStrokeColor(colors.HexColor("#2876AE"))
    canvas.setLineWidth(1.1)
    for deslocamento in (0, 10, 20):
        onda = canvas.beginPath()
        onda.moveTo(largura - 72 * mm, 42 * mm + deslocamento)
        onda.curveTo(largura - 47 * mm, 20 * mm + deslocamento,
                     largura - 22 * mm, 65 * mm + deslocamento,
                     largura + 10 * mm, 36 * mm + deslocamento)
        canvas.drawPath(onda, stroke=1, fill=0)

    rotulo("REGISTRO MENSAL DOS SERVIÇOS", 54 * mm, colors.HexColor("#A9DFF6"))
    paragrafo("Parâmetros da água e acompanhamento<br/>da rotina de manutenção.",
              48 * mm, tamanho=11, entrelinha=15, cor=colors.white)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica", 9)
    canvas.drawString(margem, 28 * mm, f"Data de emissão: {data_emissao}")
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(margem, 15 * mm, "Bem Star Piscinas Ltda")
    canvas.setFillColor(colors.HexColor("#D4E8F7"))
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(margem, 10 * mm, f"CNPJ {cnpj}")
    canvas.drawRightString(largura - margem, 10 * mm, "Uberlândia/MG")
    canvas.restoreState()
