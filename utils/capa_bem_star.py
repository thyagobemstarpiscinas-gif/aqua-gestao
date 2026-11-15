"""Capa do relatório mensal Bem Star, com período e fotografia selecionada."""
from html import escape
from pathlib import Path

from PIL import Image, ImageOps
import reportlab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph


def desenhar_capa_bem_star(canvas, *, cliente, mes_nome, ano, periodo_ini, periodo_fim,
                           data_emissao, logo_path, cnpj, foto_path=None):
    """A fotografia é explícita; sua ausência nunca escolhe outra imagem."""
    w, h = A4
    margem = 44
    azul, agua, cinza, linha = map(colors.HexColor, ('#104888', '#009DB8', '#516575', '#DCE7ED'))
    navy = colors.HexColor('#123D65')
    regular, bold = 'BSMensalVera', 'BSMensalVera-Bold'
    fonts = Path(reportlab.__file__).parent / 'fonts'
    for nome, arquivo in [(regular, 'Vera.ttf'), (bold, 'VeraBd.ttf')]:
        if nome not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(nome, str(fonts / arquivo)))
    canvas.saveState()

    def rect(x, top, width, height, cor):
        canvas.setFillColor(cor)
        canvas.rect(x, h-top-height, width, height, fill=1, stroke=0)

    def text(txt, x, top, size, cor=navy, forte=False, right=False):
        canvas.setFillColor(cor)
        canvas.setFont(bold if forte else regular, size)
        (canvas.drawRightString if right else canvas.drawString)(x, h-top-size*.82, str(txt))

    def fitted(txt, x, top, width, height, size, forte=False):
        for pt in range(size, 6, -1):
            p = Paragraph(escape(str(txt or '—')), ParagraphStyle('capa_mensal',
                fontName=bold if forte else regular, fontSize=pt, leading=pt*1.15,
                textColor=navy, splitLongWords=True))
            _, ph = p.wrap(width, height)
            if ph <= height:
                break
        p.drawOn(canvas, x, h-top-ph)

    periodo = f'{mes_nome} de {ano}'
    rect(0, 0, w, h, colors.white)
    rect(0, 0, w, 5, agua)
    try:
        if not logo_path:
            raise ValueError('Sem logo')
        canvas.drawImage(ImageReader(str(logo_path)), margem, h-103, 182, 60.22,
                         preserveAspectRatio=True, mask='auto')
    except (OSError, ValueError):
        text('BEM STAR PISCINAS', margem, 67, 17, azul, True)
    text('RELATÓRIO MENSAL', w-margem, 59, 8.2, azul, True, True)
    text(periodo, w-margem, 78, 10, right=True)
    rect(margem, 126, w-2*margem, .7, linha)
    text('Relatório Mensal', margem, 160, 30, forte=True)
    text('Operacional', margem, 198, 30, forte=True)
    text('Limpeza, manutenção e acompanhamento da água', margem, 247, 10.5, cinza)
    rect(margem, 282, 34, 3, agua)
    text('CONDOMÍNIO', margem, 307, 8.2, azul, True)
    fitted(cliente, margem, 329, w-2*margem, 56, 34, True)

    foto = None
    if foto_path:
        try:
            with Image.open(foto_path) as im:
                foto = ImageOps.exif_transpose(im).convert('RGB')
        except (OSError, ValueError):
            pass
    if foto is not None:
        px, top, pw, ph = margem, 393, w-2*margem, 250
        iw, ih = foto.size
        scale = max(pw/iw, ph/ih)
        dw, dh = iw*scale, ih*scale
        canvas.saveState()
        clip = canvas.beginPath()
        clip.rect(px, h-top-ph, pw, ph)
        canvas.clipPath(clip, stroke=0, fill=0)
        canvas.drawImage(ImageReader(foto), px+(pw-dw)/2, h-top-dh+min(18, dh-ph), dw, dh)
        canvas.restoreState()
        rect(margem, top+ph, 44, 3, agua)
        text('REGISTRO FOTOGRÁFICO DO PERÍODO', margem, 656, 8.2, cinza)
    else:
        rect(margem, 408, w-2*margem, 220, navy)
        text('CUIDADO CONTÍNUO COM A PISCINA', margem+24, 444, 9.5, colors.HexColor('#91DEEB'), True)
        text('Limpeza e manutenção', margem+24, 478, 22, colors.white, True)
        text('Atividades e medições registradas ao longo do mês', margem+24, 517, 9.5, colors.white)
        canvas.setStrokeColor(agua)
        canvas.setLineWidth(1)
        for d in range(3):
            p = canvas.beginPath()
            p.moveTo(margem+24, h-573-d*10)
            p.curveTo(200, h-548-d*10, 340, h-602-d*10, w-margem-24, h-573-d*10)
            canvas.drawPath(p)
    rect(margem, 696, w-2*margem, 57, colors.HexColor('#F2F7FA'))
    text('PERÍODO DO RELATÓRIO', margem+15, 708, 7.5, azul, True)
    fitted(periodo, margem+15, 726, 220, 22, 11)
    rect(301, 709, .7, 30, linha)
    text('REGISTROS DO MÊS', 320, 708, 7.5, azul, True)
    fitted(f'{periodo_ini} a {periodo_fim}', 320, 726, w-margem-335, 22, 9)
    text('Bem Star Piscinas Ltda', margem, 780, 8.8, forte=True)
    text('Uberlândia / MG', w-margem, 780, 8.8, cinza, right=True)
    text(f'CNPJ {cnpj}', margem, 796, 7.5, cinza)
    text(f'Data de emissão: {data_emissao}', w-margem, 796, 7.5, cinza, right=True)
    rect(0, h-5, w, 5, azul)
    canvas.restoreState()
