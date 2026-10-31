"""PDF consolidado (trimestral, semestral ou anual) do RT ACQUA GESTÃO.

Recebe dados já identificados e filtrados; não consulta Sheets/Drive.
"""
from __future__ import annotations

from pathlib import Path

from utils.capa_acqua import desenhar_capa
from utils.rt_relatorio import MARCA_COMPLETA, PARAMETROS, fmt_num, formatar_volume, texto_art


def gerar_pdf_consolidado(ctx: dict, pdf_path: Path, logo_path=None) -> tuple[bool, str | None]:
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_JUSTIFY
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import Image as RLImage
        from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
        from PIL import Image as PILImage
        from xml.sax.saxutils import escape

        pdf_path = Path(pdf_path)
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        azul, azul2, claro = colors.HexColor("#173A5E"), colors.HexColor("#1D78A8"), colors.HexColor("#E8F1F7")
        cinza = colors.HexColor("#4D5661")
        ss = getSampleStyleSheet()
        H1 = ParagraphStyle("H1c", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=13, textColor=azul, spaceBefore=8, spaceAfter=4)
        H2 = ParagraphStyle("H2c", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=10.5, textColor=azul2, spaceBefore=6, spaceAfter=3)
        BODY = ParagraphStyle("Bc", parent=ss["BodyText"], fontName="Helvetica", fontSize=8.8, leading=11.4, alignment=TA_JUSTIFY)
        CELL = ParagraphStyle("Cc", parent=ss["BodyText"], fontName="Helvetica", fontSize=7.6, leading=9.4)
        CELLB = ParagraphStyle("CBc", parent=CELL, fontName="Helvetica-Bold", textColor=colors.white)

        def P(t, st=BODY):
            return Paragraph(escape(str(t if t is not None else "")), st)

        def tabela(dados, larguras):
            linhas = [[P(c, CELLB if i == 0 else CELL) for c in row] for i, row in enumerate(dados)]
            t = Table(linhas, colWidths=larguras, repeatRows=1, hAlign="LEFT")
            cmds = [("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#AEBFCC")),
                    ("BACKGROUND", (0, 0), (-1, 0), azul), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
            cmds += [("BACKGROUND", (0, i), (-1, i), claro) for i in range(2, len(dados), 2)]
            t.setStyle(TableStyle(cmds))
            return t

        art = ctx.get("art") or {}
        art_txt = texto_art(art.get("status"), art.get("numero"), art.get("inicio"), art.get("fim"))
        w, h = A4

        def capa(canvas, doc):
            desenhar_capa(canvas, w, h, logo_path, f"Relatório consolidado {ctx['tipo'].lower()}",
                          "Responsabilidade Técnica — síntese do período", [
                              ("Cliente / estabelecimento", ctx.get("nome_condominio")),
                              ("Período", ctx.get("periodo_rotulo")),
                              ("Responsável técnico", f"{ctx.get('responsavel', '')} | {ctx.get('crq', '')}"),
                              ("Situação da ART", art_txt),
                              ("Data de emissão", ctx.get("data_emissao")),
                          ], rodape=MARCA_COMPLETA)

        def pagina(canvas, doc):
            canvas.saveState()
            canvas.setStrokeColor(azul2)
            canvas.line(14 * mm, h - 16 * mm, w - 14 * mm, h - 16 * mm)
            canvas.setFont("Helvetica-Bold", 8.5)
            canvas.setFillColor(azul)
            canvas.drawString(14 * mm, h - 13 * mm, "ACQUA GESTÃO")
            canvas.drawRightString(w - 14 * mm, h - 13 * mm, f"RELATÓRIO CONSOLIDADO — {ctx.get('periodo_rotulo', '').upper()}")
            canvas.line(14 * mm, 13 * mm, w - 14 * mm, 13 * mm)
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(cinza)
            canvas.drawString(14 * mm, 8.5 * mm, MARCA_COMPLETA)
            canvas.drawRightString(w - 14 * mm, 8.5 * mm, f"Página {doc.page}")
            canvas.restoreState()

        doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                                topMargin=22 * mm, bottomMargin=18 * mm,
                                title=f"Relatório consolidado RT — {ctx.get('periodo_rotulo', '')}",
                                author=ctx.get("responsavel", ""))
        story = [PageBreak()]
        story.append(P("1. IDENTIFICAÇÃO", H1))
        story.append(tabela([["Campo", "Informação"],
                             ["Cliente / estabelecimento", ctx.get("nome_condominio", "")],
                             ["Endereço", ctx.get("endereco", "") or "—"],
                             ["Período consolidado", f"{ctx.get('periodo_rotulo', '')} ({', '.join(ctx.get('meses', []))})"],
                             ["Situação da ART", art_txt],
                             ["Fonte dos dados", "Visitas de campo registradas e identificadas por cliente e período."]],
                            [55 * mm, 127 * mm]))
        pisc = ctx.get("piscinas_cadastro") or {}
        story.append(P("2. PISCINAS DO CADASTRO", H1))
        if pisc.get("piscinas"):
            story.append(tabela([["Identificação", "Tipo", "Volume"]] +
                                [[p["nome"], p["tipo"], formatar_volume(p["volume"])] for p in pisc["piscinas"]],
                                [80 * mm, 52 * mm, 50 * mm]))
        else:
            story.append(P("Nenhuma piscina individual cadastrada para este cliente."))
        if pisc.get("sem_distribuicao"):
            story.append(P(f"Volume total informado sem distribuição por piscina: {formatar_volume(pisc.get('volume_total'))}."))

        cons = ctx.get("consolidado") or {}
        story.append(P("3. SÍNTESE DOS PARÂMETROS POR PISCINA", H1))
        story.append(P("Leitura inicial de cada visita (quando registrada como antes/após, considera-se o primeiro valor). "
                       "Conformidade: percentual de leituras dentro da faixa de referência."))
        if not cons.get("piscinas"):
            story.append(P("Nenhuma leitura encontrada para o período selecionado."))
        for nome, grupo in (cons.get("piscinas") or {}).items():
            story.append(P(f"{nome} — {len(grupo['linhas'])} registro(s)", H2))
            linhas = [["Parâmetro", "Faixa", "Leituras", "Mínimo", "Média", "Máximo", "Conformidade"]]
            for chave, *_ in PARAMETROS:
                p = grupo["parametros"].get(chave)
                if not p:
                    continue
                casas = 2 if p["max"] < 20 else 0
                linhas.append([p["rotulo"], f"{fmt_num(p['faixa'][0], 1)}–{fmt_num(p['faixa'][1], 1)} {p['unidade']}".strip(),
                               str(p["n"]), fmt_num(p["min"], casas), fmt_num(p["media"], casas), fmt_num(p["max"], casas),
                               f"{fmt_num(p['conformidade'], 0)}%"])
            story.append(tabela(linhas, [44 * mm, 34 * mm, 18 * mm, 20 * mm, 20 * mm, 20 * mm, 26 * mm]))
            if grupo["mensal"]:
                cab = ["Mês"] + [r for k, r, *_ in PARAMETROS if k in grupo["parametros"]]
                mensal = [cab]
                for mes in cons.get("meses", []):
                    vals = grupo["mensal"].get(mes)
                    if vals:
                        mensal.append([mes] + [fmt_num(vals.get(k)) for k, *_ in PARAMETROS if k in grupo["parametros"]])
                story.append(Spacer(1, 2 * mm))
                story.append(P("Médias mensais", H2))
                story.append(tabela(mensal, [22 * mm] + [160 * mm / max(len(cab) - 1, 1)] * (len(cab) - 1)))
        if cons.get("descartadas"):
            story.append(P(f"{cons['descartadas']} registro(s) sem data válida ou fora do período não foram considerados."))

        graficos = [(n, Path(g)) for n, g in (ctx.get("graficos") or []) if Path(g).exists()]
        if graficos:
            story.append(PageBreak())
            story.append(P("4. GRÁFICOS HISTÓRICOS POR PISCINA", H1))
            for nome, g in graficos:
                iw, ih = PILImage.open(g).size
                largura = 182 * mm
                story.append(P(nome, H2))
                story.append(RLImage(str(g), width=largura, height=largura * ih / iw))
                story.append(Spacer(1, 4 * mm))

        story.append(P("5. CONSIDERAÇÕES DO RESPONSÁVEL TÉCNICO", H1))
        story.append(P(ctx.get("parecer") or "Síntese elaborada a partir dos registros do período. Recomenda-se a manutenção "
                       "da rotina de monitoramento e a correção imediata de leituras fora da faixa de referência."))
        story.append(Spacer(1, 14 * mm))
        story.append(P("_______________________________________________", CELL))
        story.append(P(f"{ctx.get('responsavel', '')} — {ctx.get('crq', '')}", CELL))
        doc.build(story, onFirstPage=capa, onLaterPages=pagina)
        return pdf_path.exists(), None
    except Exception as e:  # noqa: BLE001 - o chamador exibe a mensagem
        return False, f"{type(e).__name__}: {e}"
