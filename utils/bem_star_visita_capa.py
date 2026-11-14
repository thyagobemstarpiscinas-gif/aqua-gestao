"""Capa e escolha de fotografia da visita operacional Bem Star (não mensal RT)."""
import base64
import hashlib
import io
from datetime import datetime
from html import escape
from pathlib import Path

from PIL import Image, ImageOps

from utils.formatacao import limpar_nome_arquivo


APRESENTACAO = (
    ("Registro<br/>da visita", "Este relatório da Bem Star Piscinas registra os serviços de limpeza e manutenção "
     "e as medições realizadas nesta visita. Seu conteúdo documenta a rotina operacional do condomínio."),
    ("Acompanhamento<br/>técnico", "O acompanhamento da ACQUA GESTÃO complementa a manutenção com interpretação "
     "das análises, identificação de tendências e orientações para o tratamento da água. "
     "O serviço de RT e seu relatório possuem escopo próprio."),
    ("Turbidez e<br/>qualidade da água", "A <b>ABNT NBR 10818:2026, item 4.5</b>, estabelece turbidez inferior a "
     "<b>0,5 NTU</b>, com referência à NBR 10339. A <b>ACQUA GESTÃO</b> oferece essa medição com turbidímetro "
     "e orientação técnica dos resultados. Consulte as opções de acompanhamento para o condomínio."),
)
INTRO_ANALISE = (
    "Medições realizadas com fotômetro Color Q para acompanhamento operacional da água. "
    "A avaliação considera os parâmetros apresentados; análises complementares podem ampliar esse acompanhamento."
)
NOTA_OPERACIONAL = (
    "Documento operacional emitido pela Bem Star Piscinas Ltda., destinado ao registro das atividades e medições "
    "desta visita. Não constitui Relatório de Responsabilidade Técnica (RT).<br/><br/>"
    "Para condomínios residenciais, o artigo 7º da <b>Resolução CFQ nº 332/2025</b>, atualizada pela "
    "<b>Resolução CFQ nº 345/2026</b>, trata da fiscalização do CRQ com caráter de prevenção e orientação, "
    "voltada à identificação de quem responde pela análise e pelo controle da água das piscinas coletivas.<br/><br/>"
    "A <b>ACQUA GESTÃO</b> oferece acompanhamento técnico complementar, com análises, interpretação de resultados "
    "e orientações de tratamento, conforme o serviço contratado."
)


def opcoes_capa_depois(uploads, rascunhos, carregar):
    """Mesmos arquivos do preview: upload atual prevalece sobre rascunho de mesmo nome."""
    nomes_upload = {limpar_nome_arquivo(getattr(f, "name", "")) for f in uploads or []}
    fontes = list(uploads or []) + [p for p in rascunhos or []
        if limpar_nome_arquivo(Path(p).name.removeprefix("rasc_depois_")) not in nomes_upload]
    opcoes, vistos = [], set()
    for fonte in fontes:
        try:
            bruto = bytes(fonte.getvalue()) if hasattr(fonte, "getvalue") else Path(fonte).read_bytes()
            chave = hashlib.sha256(bruto).hexdigest()
            if chave in vistos:
                continue
            imagem = carregar(fonte)
            if imagem is None:
                continue
            vistos.add(chave)
            imagem.thumbnail((640, 640))
            nome = str(getattr(fonte, "name", fonte)).removeprefix("rasc_depois_")
            opcoes.append({"chave": chave, "nome": nome, "imagem": imagem})
        except (OSError, ValueError, TypeError):
            continue
    return opcoes


def renderizar_seletor_capa(st, uploads, rascunhos, carregar, salvar):
    opcoes = opcoes_capa_depois(uploads, rascunhos, carregar)
    ss = st.session_state
    st.markdown("**Foto da capa — relatório Bem Star**")
    if not opcoes:
        if ss.get("op_capa_depois_hash"):
            ss["op_capa_depois_hash"] = ""
            salvar()
        st.caption("Anexe uma foto em Depois do tratamento para incluí-la na capa. Sem foto válida, a capa sai apenas com os dados da visita.")
        return
    chaves = [o["chave"] for o in opcoes]
    preferida = ss.get("op_capa_depois_hash")
    escolhida = preferida if preferida in chaves else chaves[0]
    chave_widget = f"_op_capa_depois_escolha_{ss.get('_op_fotos_nonce', 0)}"
    # Reconciliação antes de instanciar o widget; a chave durável fica no rascunho.
    if ss.get(chave_widget) != escolhida:
        ss[chave_widget] = escolhida

    def ao_trocar():
        ss["op_capa_depois_hash"] = ss[chave_widget]
        salvar()

    rotulos = {o["chave"]: f"Foto {i + 1} — {o['nome']}" for i, o in enumerate(opcoes)}
    escolhida = st.selectbox("Foto da capa", chaves, format_func=rotulos.get,
                            key=chave_widget, on_change=ao_trocar)
    if ss.get("op_capa_depois_hash") != escolhida:
        ss["op_capa_depois_hash"] = escolhida
        salvar()
    st.image(next(o["imagem"] for o in opcoes if o["chave"] == escolhida), width=300)
    st.caption("A foto também permanece em Após o tratamento, no registro fotográfico.")


def registrar_escolha_capa(destino, escolhida, categoria, bruto, arquivo, drive_id, b64):
    """Guarda a referência exata no lançamento, sem depender da posição na lista de IDs."""
    if categoria == "depois" and escolhida and hashlib.sha256(bruto).hexdigest() == escolhida:
        destino.update({"referencia": {"arquivo": arquivo, "drive_id": drive_id or ""}, "b64": b64})


def _abrir_imagem(bruto):
    try:
        with Image.open(io.BytesIO(bruto)) as imagem:
            imagem.load()
            return ImageOps.exif_transpose(imagem).convert("RGB")
    except (OSError, ValueError, TypeError):
        return None


def foto_capa_visita(lancamento, fotos_depois, baixar, pasta_campo):
    """Uma escolha explícita indisponível não é silenciosamente trocada por outra foto."""
    ref = lancamento.get("capa_depois")
    if isinstance(ref, dict) and ref:
        arquivo, fid = ref.get("arquivo", ""), ref.get("drive_id", "")
        # Só referências pertencentes a Depois, nunca Casa de máquinas ou outra visita.
        nome_valido = arquivo and arquivo in (lancamento.get("fotos_depois") or [])
        id_valido = fid and fid in (lancamento.get("fotos_depois_ids") or [])
        if not (nome_valido or id_valido):
            return None
        try:
            imagem = _abrir_imagem(base64.b64decode(lancamento.get("capa_depois_b64") or ""))
        except (ValueError, TypeError):
            imagem = None
        if imagem is not None:
            return imagem
        if id_valido:
            try:
                imagem = _abrir_imagem(baixar(fid))
                if imagem is not None:
                    return imagem
            except Exception:
                pass  # A capa sem foto continua disponível se o Drive falhar.
        if nome_valido and Path(arquivo).name == arquivo:
            try:
                return _abrir_imagem((Path(pasta_campo) / arquivo).read_bytes())
            except OSError:
                pass
        return None
    # Visitas antigas: somente a primeira imagem válida explicitamente declarada em Depois.
    if not any(lancamento.get(f"fotos_depois{s}") for s in ("", "_ids", "_b64")):
        return None
    for foto in fotos_depois or []:
        try:
            imagem = _abrir_imagem(base64.b64decode(foto))
            if imagem is not None:
                return imagem
        except (ValueError, TypeError):
            continue
    return None


def desenhar_capa_visita(canvas, *, cliente, data, operador, logo, foto=None):
    """Capa aprovada: marca em fundo branco, condomínio, fotografia e identificação."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import Paragraph

    W, H = A4
    M = 44
    navy, azul, agua, cinza, linha = map(colors.HexColor,
        ("#123D65", "#104888", "#009DB8", "#516575", "#DCE7ED"))
    regular, bold = "BSansVera", "BSansVera-Bold"
    canvas.saveState()

    def rect(x, top, w, h, cor):
        canvas.setFillColor(cor); canvas.rect(x, H-top-h, w, h, fill=1, stroke=0)

    def text(txt, x, top, size, cor=navy, forte=False, right=False):
        canvas.setFillColor(cor); canvas.setFont(bold if forte else regular, size)
        (canvas.drawRightString if right else canvas.drawString)(x, H-top-size*.82, txt)

    def fitted(txt, x, top, width, height, size, forte=False, cor=navy):
        for pt in range(size, 6, -1):
            p = Paragraph(escape(str(txt or "—")), ParagraphStyle("capa", fontName=bold if forte else regular,
                fontSize=pt, leading=pt*1.15, textColor=cor, splitLongWords=True))
            _, ph = p.wrap(width, height)
            if ph <= height:
                break
        p.drawOn(canvas, x, H-top-ph)

    meses = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
             "setembro", "outubro", "novembro", "dezembro")
    data_longa, data_curta = str(data or "—"), str(data or "—")
    try:
        dt = datetime.strptime(str(data), "%d/%m/%Y")
        data_longa = f"{dt.day} de {meses[dt.month-1]} de {dt.year}"
        data_curta = f"{dt.day:02d} {meses[dt.month-1][:3].upper()} {dt.year}"
    except ValueError:
        pass

    rect(0, 0, W, H, colors.white); rect(0, 0, W, 5, agua)
    if logo:
        try:
            canvas.drawImage(ImageReader(str(logo)), M, H-103, 182, 60.22,
                             preserveAspectRatio=True, mask="auto")
        except (OSError, ValueError):
            text("BEM STAR PISCINAS", M, 67, 17, azul, True)
    else:
        text("BEM STAR PISCINAS", M, 67, 17, azul, True)
    text("VISITA OPERACIONAL", W-M, 59, 8.2, azul, True, True)
    text(data_curta, W-M, 78, 11, right=True)
    rect(M, 126, W-2*M, .7, linha)
    text("Relatório de visita", M, 160, 30, forte=True)
    text("técnica operacional", M, 198, 30, forte=True)
    text("Limpeza, manutenção e controle da água", M, 247, 11, cinza)
    rect(M, 282, 34, 3, agua)
    text("CONDOMÍNIO", M, 307, 8.2, azul, True)
    fitted(cliente, M, 329, W-2*M, 56, 34, True)
    if foto is not None:
        px, py, pw, ph = M, 393, W-2*M, 250
        iw, ih = foto.size; scale = max(pw/iw, ph/ih)
        dw, dh = iw*scale, ih*scale
        canvas.saveState()
        clip = canvas.beginPath(); clip.rect(px, H-py-ph, pw, ph)
        canvas.clipPath(clip, stroke=0, fill=0)
        canvas.drawImage(ImageReader(foto), px+(pw-dw)/2, H-py-dh+min(18, dh-ph), dw, dh)
        canvas.restoreState()
        rect(M, py+ph, 44, 3, agua)
        text("REGISTRO DA VISITA · APÓS O TRATAMENTO", M, 656, 8.2, cinza)
    else:
        rect(M, 408, W-2*M, 220, navy)
        text("CUIDADO COM A PISCINA", M+24, 444, 10, colors.HexColor("#91DEEB"), True)
        text("Limpeza e manutenção", M+24, 478, 22, colors.white, True)
        text("Registro das atividades e medições da visita", M+24, 517, 10, colors.white)
        canvas.setStrokeColor(agua); canvas.setLineWidth(1)
        for d in range(3):
            p=canvas.beginPath(); p.moveTo(M+24, H-573-d*10)
            p.curveTo(200,H-548-d*10,340,H-602-d*10,W-M-24,H-573-d*10)
            canvas.drawPath(p)
    rect(M, 696, W-2*M, 57, colors.HexColor("#F2F7FA"))
    text("DATA DA VISITA", M+15, 708, 7.5, azul, True)
    fitted(data_longa, M+15, 726, 225, 22, 10)
    rect(301, 709, .7, 30, linha)
    text("OPERADOR", 320, 708, 7.5, azul, True)
    fitted(operador, 320, 726, W-M-335, 22, 10)
    text("Bem Star Piscinas Ltda", M, 780, 8.8, forte=True)
    text("Uberlândia / MG", W-M, 780, 8.8, cinza, right=True)
    text("CNPJ 26.799.958/0001-88", M, 796, 7.5, cinza)
    rect(0, H-5, W, 5, azul)
    canvas.restoreState()
