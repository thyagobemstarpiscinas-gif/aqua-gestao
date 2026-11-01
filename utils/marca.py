"""Identidade visual: logomarcas aprovadas da ACQUA GESTÃO e da Bem Star Piscinas.

A ACQUA GESTÃO usa somente `acqua_gestao_logo_rt.png` (grafia com dois C). As
artes antigas com a grafia "AQUA GESTÃO" continuam no repositório apenas como
histórico e nunca devem ser escolhidas por um gerador. Sem a arte aprovada, o
documento sai sem imagem (marca textual), nunca com a arte antiga.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]

LOGO_ACQUA = "acqua_gestao_logo_rt.png"
LOGO_BEM_STAR = "bem_star_logo.png"

# SHA-256 das artes antigas "AQUA GESTÃO" (arquivos PNG e imagens embutidas
# nos modelos DOCX). Usado para trocar a imagem nos modelos e nos testes.
LOGOS_AQUA_ANTIGOS_SHA256 = frozenset({
    "b01f6539831c9c04abb97ed5172e252077b8dbaa5a6262b9fdc31efa3d3d0899",  # aqua_gestao_logo.png / aditivo.docx
    "1e11a9984365a6b440f1e1a5914bc7a88d67727fdd8e78ade46d5bef6c63d8bb",  # aqua_gestao_logo_oficial.png
    "5c6849a035a15e6e1c73a53aeb64fd447ffc7b2d36fb9697ee96246d34424ed6",  # relatorio_mensal.docx
    "06c149960170996323c5775fdfbffcc3592f9a9057293b83c3f39f46acc77a8c",  # template_rt_aqua(_v2).docx
    "82ca32b384beed123468e8f0263756a8bc18a179e707d483b1b9d40de9f502b1",  # template_rt_aqua_v2_relatorio.docx
})


def logo_acqua(base_dir: Path | str | None = None) -> Path | None:
    caminho = Path(base_dir or RAIZ) / LOGO_ACQUA
    return caminho if caminho.is_file() else None


def logo_bem_star(base_dir: Path | str | None = None) -> Path | None:
    caminho = Path(base_dir or RAIZ) / LOGO_BEM_STAR
    return caminho if caminho.is_file() else None


def sha256(dados: bytes) -> str:
    return hashlib.sha256(dados).hexdigest()


def substituir_logo_antigo_docx(doc, logo_path: Path | str | None = None) -> int:
    """Troca, em corpo, cabeçalhos e rodapés, a arte antiga pela aprovada.

    Só imagens cujo conteúdo coincide com uma arte antiga conhecida são trocadas;
    fotos e a marca Bem Star nunca são alteradas. A nova imagem (quadrada) é
    encaixada na caixa original, sem distorção. Retorna o número de trocas.
    """
    from docx.oxml.ns import qn

    logo_path = Path(logo_path) if logo_path else logo_acqua()
    if not logo_path or not logo_path.is_file():
        return 0
    from PIL import Image
    with Image.open(logo_path) as im:
        proporcao = im.width / im.height

    partes = [doc.part]
    for secao in doc.sections:
        for hf in (secao.header, secao.first_page_header, secao.even_page_header,
                   secao.footer, secao.first_page_footer, secao.even_page_footer):
            try:
                if not hf.is_linked_to_previous:
                    partes.append(hf.part)
            except Exception:
                continue

    trocas = 0
    vistos = set()
    for parte in partes:
        if id(parte) in vistos:
            continue
        vistos.add(id(parte))
        antigos = {rid for rid, rel in parte.rels.items()
                   if "image" in rel.reltype and not rel.is_external
                   and sha256(rel.target_part.blob) in LOGOS_AQUA_ANTIGOS_SHA256}
        if not antigos:
            continue
        novo_rid, _img = parte.get_or_add_image(str(logo_path))
        for blip in parte.element.iter(qn("a:blip")):
            if blip.get(qn("r:embed")) not in antigos:
                continue
            blip.set(qn("r:embed"), novo_rid)
            trocas += 1
            # Encaixa a imagem na caixa original mantendo a proporção.
            container = blip
            while container is not None and container.tag not in (qn("wp:inline"), qn("wp:anchor")):
                container = container.getparent()
            if container is None:
                continue
            extent = container.find(qn("wp:extent"))
            if extent is None:
                continue
            cx, cy = int(extent.get("cx")), int(extent.get("cy"))
            if cx / cy > proporcao:
                novo_cx, novo_cy = int(cy * proporcao), cy
            else:
                novo_cx, novo_cy = cx, int(cx / proporcao)
            extent.set("cx", str(novo_cx))
            extent.set("cy", str(novo_cy))
            for ext in container.iter(qn("a:ext")):
                if ext.get("cx") is not None:
                    ext.set("cx", str(novo_cx))
                    ext.set("cy", str(novo_cy))
        # Remove as relações órfãs: a arte antiga deixa de ir no arquivo gerado.
        for rid in antigos:
            if not any(el.get(qn("r:embed")) == rid or el.get(qn("r:id")) == rid
                       for el in parte.element.iter()):
                parte.rels.pop(rid, None)
    return trocas


def logo_acqua_data_uri(lado: int = 160) -> str:
    """Miniatura PNG da logo aprovada como data URI (para prévias HTML)."""
    caminho = logo_acqua()
    if not caminho:
        return ""
    import base64
    import io
    from PIL import Image
    with Image.open(caminho) as im:
        im = im.convert("RGBA")
        im.thumbnail((lado, lado))
        buf = io.BytesIO()
        im.save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


# Grafias antigas do nome-marca em títulos/cabeçalhos. A razão social
# "AQUA GESTÃO CONTROLE TÉCNICO LTDA" (com CNPJ) não é alterada aqui.
GRAFIAS_MARCA = (
    ("AQUA GESTÃO – CONTROLE TÉCNICO DE PISCINAS", "ACQUA GESTÃO – CONTROLE TÉCNICO DE PISCINAS"),
    ("AQUA GESTÃO — CONTROLE TÉCNICO DE PISCINAS", "ACQUA GESTÃO — CONTROLE TÉCNICO DE PISCINAS"),
    ("AQUA GESTÃO - CONTROLE TÉCNICO DE PISCINAS", "ACQUA GESTÃO - CONTROLE TÉCNICO DE PISCINAS"),
)


def corrigir_grafia_marca_docx(doc) -> int:
    """Corrige a grafia do nome-marca em corpo, tabelas, caixas de texto, cabeçalhos e rodapés."""
    from docx.oxml.ns import qn

    partes = [doc.part]
    for secao in doc.sections:
        for hf in (secao.header, secao.first_page_header, secao.even_page_header,
                   secao.footer, secao.first_page_footer, secao.even_page_footer):
            try:
                partes.append(hf.part)
            except Exception:
                continue
    trocas, vistos = 0, set()
    for parte in partes:
        if id(parte) in vistos:
            continue
        vistos.add(id(parte))
        for paragrafo in parte.element.iter(qn("w:p")):
            textos = list(paragrafo.iter(qn("w:t")))
            completo = "".join(t.text or "" for t in textos)
            if "AQUA GEST" not in completo:
                continue
            novo = completo
            for antigo, certo in GRAFIAS_MARCA:
                novo = novo.replace(antigo, certo)
            if novo == completo:
                continue
            # Texto dividido em vários trechos: concentra no primeiro (mantém
            # a formatação do início do parágrafo).
            textos[0].text = novo
            textos[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for t in textos[1:]:
                t.text = ""
            trocas += 1
    return trocas
