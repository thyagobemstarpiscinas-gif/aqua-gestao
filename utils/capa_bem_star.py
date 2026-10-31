"""Capa vetorial Bem Star, com a paleta da marca e natureza operacional explícita."""
from html import escape
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph


def desenhar_capa_bem_star(canvas, *, cliente, mes_nome, ano, periodo_ini, periodo_fim,
                           data_emissao, logo_path, cnpj):
    w, h = A4
    azul, agua = colors.HexColor('#104888'), colors.HexColor('#0090D0')
    canvas.saveState()
    canvas.setFillColor(colors.HexColor('#F0F6FC'))
    canvas.rect(0, 0, w, h, fill=1, stroke=0)
    # Cabeçalho claro reservado à marca original.
    canvas.setFillColor(colors.white)
    canvas.rect(0, h-59*mm, w, 59*mm, fill=1, stroke=0)
    try:
        if not logo_path:
            raise ValueError()
        canvas.drawImage(ImageReader(str(logo_path)), 21*mm, h-42*mm,
                         width=76*mm, height=25*mm, preserveAspectRatio=True, mask='auto')
    except Exception:
        canvas.setFillColor(azul)
        canvas.setFont('Helvetica-Bold', 20)
        canvas.drawString(21*mm, h-35*mm, 'BEM STAR PISCINAS')
    canvas.setFillColor(azul)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.drawRightString(w-21*mm, h-27*mm, 'RELATÓRIO MENSAL')
    canvas.setFont('Helvetica', 9)
    canvas.drawRightString(w-21*mm, h-34*mm, f'{mes_nome} / {ano}')
    # Campo central com presença de cor, ondas e linhas de uma piscina vista de cima.
    canvas.setFillColor(azul)
    canvas.rect(0, 65*mm, w, h-124*mm, fill=1, stroke=0)
    canvas.setFillColor(agua)
    canvas.rect(0, h-61*mm, w, 2*mm, fill=1, stroke=0)
    canvas.setStrokeColor(colors.HexColor('#2366A4'))
    canvas.setLineWidth(.55)
    for x in range(135, 221, 13):
        canvas.line(x*mm, 66*mm, x*mm, h-61*mm)
    for y in range(72, 238, 13):
        canvas.line(133*mm, y*mm, w, y*mm)
    canvas.setStrokeColor(colors.HexColor('#268ABF'))
    canvas.setLineWidth(1.3)
    for d in range(4):
        p = canvas.beginPath(); p.moveTo(128*mm, (81+d*7)*mm)
        p.curveTo(160*mm, (110+d*7)*mm, 181*mm, (57+d*7)*mm, 222*mm, (89+d*7)*mm)
        canvas.drawPath(p, stroke=1)

    def text(txt, top, size, leading, color=colors.white, bold=False, width=166*mm):
        para = Paragraph(txt, ParagraphStyle('capa', fontName='Helvetica-Bold' if bold else 'Helvetica',
                                            fontSize=size, leading=leading, textColor=color))
        _, ph = para.wrap(width, h)
        para.drawOn(canvas, 21*mm, top-ph)
        return ph
    text('MANUTENÇÃO &amp; ACOMPANHAMENTO', h-77*mm, 9, 12, colors.HexColor('#9FDFFF'), True)
    text('Cuidado contínuo.<br/>Gestão com clareza.', h-94*mm, 31, 37, bold=True)
    text('Relatório operacional de manutenção', h-127*mm, 13, 18)
    text('CLIENTE / LOCAL ATENDIDO', h-157*mm, 8, 11, colors.HexColor('#9FDFFF'), True)
    for size in range(22, 10, -1):
        p = Paragraph(escape(str(cliente)), ParagraphStyle('nome', fontName='Helvetica-Bold', fontSize=size, leading=size+4))
        if p.wrap(161*mm, h)[1] <= 29*mm:
            break
    text(escape(str(cliente)), h-167*mm, size, size+4, bold=True, width=161*mm)
    text(f'{escape(mes_nome)} / {escape(str(ano))}', 91*mm, 17, 21, bold=True)
    text(f'{escape(periodo_ini)} a {escape(periodo_fim)}', 80*mm, 9, 12, colors.HexColor('#C5E7FF'))
    text('REGISTRO OPERACIONAL · NÃO É RELATÓRIO DE RT', 53*mm, 9, 13, azul, True)
    text('Consolida a rotina de manutenção e as leituras informadas.<br/>Não constitui laudo ou ART, nem substitui o acompanhamento técnico quando aplicável.',
         45*mm, 9, 13, colors.HexColor('#455C74'))
    canvas.setStrokeColor(colors.HexColor('#C5D8EB'))
    canvas.line(21*mm, 23*mm, w-21*mm, 23*mm)
    canvas.setFillColor(azul)
    canvas.setFont('Helvetica', 8)
    canvas.drawString(21*mm, 17*mm, f'Data de emissão: {data_emissao}')
    canvas.setFont('Helvetica', 7.5)
    canvas.drawString(21*mm, 11*mm, f'Bem Star Piscinas Ltda | CNPJ {cnpj}')
    canvas.drawRightString(w-21*mm, 11*mm, 'Uberlândia/MG')
    canvas.restoreState()
