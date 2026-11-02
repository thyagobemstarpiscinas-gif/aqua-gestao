"""Capa dos relatórios RT da ACQUA GESTÃO (ReportLab). Exclusiva do RT; o
relatório Bem Star mantém sua própria capa em utils/capa_bem_star.py."""
from __future__ import annotations

from pathlib import Path

from utils.rt_relatorio import MARCA

AZUL_NOITE = "#0B1F3A"
AZUL_MEDIO = "#173A5E"
DOURADO = "#C9A227"
DOURADO_CLARO = "#E6CC7A"


def desenhar_capa(canvas, largura, altura, logo_path, titulo: str, subtitulo: str,
                  linhas: list[tuple[str, str]], rodape: str = ""):
    """Página inteira: fundo azul-noite, logo em destaque e quadro de identificação."""
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.utils import simpleSplit
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import reportlab

    fonte_dir = Path(reportlab.__file__).parent / "fonts"
    if "AcquaCapa" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("AcquaCapa", str(fonte_dir / "Vera.ttf")))
        pdfmetrics.registerFont(TTFont("AcquaCapa-Bold", str(fonte_dir / "VeraBd.ttf")))
    pdfmetrics.registerFontFamily("AcquaCapa", normal="AcquaCapa", bold="AcquaCapa-Bold")

    azul, azul2 = colors.HexColor(AZUL_NOITE), colors.HexColor(AZUL_MEDIO)
    ouro, ouro2 = colors.HexColor(DOURADO), colors.HexColor(DOURADO_CLARO)
    canvas.saveState()
    canvas.setFillColor(azul)
    canvas.rect(0, 0, largura, altura, stroke=0, fill=1)
    # Faixas douradas superior e inferior.
    canvas.setFillColor(ouro)
    canvas.rect(0, altura - 7 * mm, largura, 2.2 * mm, stroke=0, fill=1)
    canvas.rect(0, 9 * mm, largura, 1.2 * mm, stroke=0, fill=1)

    # Logo com proporção preservada e área generosa (≈ 40% da largura da página).
    lado = 92 * mm
    topo_logo = altura - 22 * mm
    if logo_path and Path(logo_path).exists():
        try:
            canvas.drawImage(str(logo_path), (largura - lado) / 2, topo_logo - lado, width=lado, height=lado,
                             preserveAspectRatio=True, anchor="c", mask="auto")
        except Exception:
            pass
    y = topo_logo - lado - 14 * mm
    canvas.setFillColor(ouro)
    canvas.setFont("AcquaCapa-Bold", 34)
    canvas.drawCentredString(largura / 2, y, MARCA)
    y -= 8 * mm
    canvas.setFillColor(ouro2)
    canvas.setFont("AcquaCapa", 11.5)
    canvas.drawCentredString(largura / 2, y, "CONTROLE TÉCNICO DE PISCINAS  •  RESPONSABILIDADE TÉCNICA")
    y -= 7 * mm
    canvas.setStrokeColor(ouro)
    canvas.setLineWidth(0.9)
    canvas.line(largura / 2 - 45 * mm, y, largura / 2 + 45 * mm, y)

    y -= 14 * mm
    canvas.setFillColor(colors.white)
    canvas.setFont("AcquaCapa-Bold", 19)
    for parte in simpleSplit(titulo.upper(), "AcquaCapa-Bold", 19, largura - 40 * mm):
        canvas.drawCentredString(largura / 2, y, parte)
        y -= 8 * mm
    if subtitulo:
        canvas.setFont("AcquaCapa", 11)
        canvas.setFillColor(ouro2)
        canvas.drawCentredString(largura / 2, y, subtitulo)
        y -= 6 * mm

    # Quadro de identificação.
    margem = 24 * mm
    alt_linha = 9.5 * mm
    linhas = [(r, str(v or "—")) for r, v in linhas]
    quebras = [simpleSplit(v, "AcquaCapa-Bold", 10.5, largura - 2 * margem - 58 * mm)[:2] or ["—"] for _, v in linhas]
    alt_total = sum(alt_linha + (len(q) - 1) * 4.5 * mm for q in quebras) + 8 * mm
    base = max(24 * mm, y - 8 * mm - alt_total)
    canvas.setFillColor(azul2)
    canvas.setStrokeColor(ouro)
    canvas.setLineWidth(1)
    canvas.roundRect(margem, base, largura - 2 * margem, alt_total, 4 * mm, stroke=1, fill=1)
    yy = base + alt_total - 4 * mm
    for (rotulo, _), partes in zip(linhas, quebras):
        yy -= 6.5 * mm
        canvas.setFont("AcquaCapa", 8.5)
        canvas.setFillColor(ouro2)
        canvas.drawString(margem + 7 * mm, yy, rotulo.upper())
        canvas.setFont("AcquaCapa-Bold", 10.5)
        canvas.setFillColor(colors.white)
        for k, parte in enumerate(partes):
            canvas.drawString(margem + 58 * mm, yy - k * 4.5 * mm, parte)
        yy -= 3 * mm + (len(partes) - 1) * 4.5 * mm

    if rodape:
        canvas.setFont("AcquaCapa", 8)
        canvas.setFillColor(ouro2)
        canvas.drawCentredString(largura / 2, 14 * mm, rodape)
    canvas.restoreState()
