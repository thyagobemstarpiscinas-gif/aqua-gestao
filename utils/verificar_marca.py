"""Classifica imagens de documentos gerados: ACQUA aprovada, AQUA antiga ou Bem Star.

Compara impressões digitais reduzidas (48x48, tons de cinza, fundo branco), de
modo que cópias redimensionadas ou recomprimidas da mesma arte sejam
reconhecidas. Usado pelos testes de regressão de marca e pela inspeção dos
documentos fictícios.
"""
from __future__ import annotations

import base64
import io
import re
import zipfile
from pathlib import Path

from PIL import Image

from utils.marca import LOGO_ACQUA, LOGO_BEM_STAR, RAIZ

_LADO = 48
_LIMIAR = 22.0  # diferença média máxima (0–255) para considerar a mesma arte


def _digital(img: Image.Image) -> list[int]:
    img = img.convert("RGBA")
    fundo = Image.new("RGBA", img.size, (255, 255, 255, 255))
    fundo.alpha_composite(img)
    bbox = fundo.convert("L").point(lambda v: 255 if v < 245 else 0).getbbox()
    if bbox:
        fundo = fundo.crop(bbox)
    return list(fundo.convert("L").resize((_LADO, _LADO)).getdata())


def _abrir(dados: bytes) -> Image.Image | None:
    try:
        img = Image.open(io.BytesIO(dados))
        img.load()
        return img
    except Exception:
        return None


def _referencias() -> dict[str, list[list[int]]]:
    refs: dict[str, list[list[int]]] = {"acqua": [], "aqua_antiga": [], "bem_star": []}
    with Image.open(RAIZ / LOGO_ACQUA) as im:
        refs["acqua"].append(_digital(im))
    with Image.open(RAIZ / LOGO_BEM_STAR) as im:
        refs["bem_star"].append(_digital(im))
    for nome in ("aqua_gestao_logo.png", "aqua_gestao_logo_oficial.png"):
        if (RAIZ / nome).is_file():
            with Image.open(RAIZ / nome) as im:
                refs["aqua_antiga"].append(_digital(im))
    for modelo in ("relatorio_mensal.docx", "template_rt_aqua.docx", "template_rt_aqua_v2_relatorio.docx"):
        with zipfile.ZipFile(RAIZ / modelo) as z:
            for n in z.namelist():
                if n.startswith("word/media/") and not n.endswith("/"):
                    img = _abrir(z.read(n))
                    if img:
                        refs["aqua_antiga"].append(_digital(img))
    return refs


_REFS: dict | None = None


def classificar(img: Image.Image) -> str | None:
    """Retorna "acqua", "aqua_antiga", "bem_star" ou None (foto/gráfico/outro)."""
    global _REFS
    if _REFS is None:
        _REFS = _referencias()
    d = _digital(img)
    melhor, distancia = None, _LIMIAR
    for rotulo, lista in _REFS.items():
        for ref in lista:
            dist = sum(abs(a - b) for a, b in zip(d, ref)) / len(d)
            if dist < distancia:
                melhor, distancia = rotulo, dist
    return melhor


def imagens_do_arquivo(caminho: Path) -> list[Image.Image]:
    caminho = Path(caminho)
    sufixo = caminho.suffix.lower()
    saida = []
    if sufixo == ".pdf":
        import fitz
        with fitz.open(caminho) as doc:
            for pagina in doc:
                for entrada in pagina.get_images(full=True):
                    try:
                        pix = fitz.Pixmap(doc, entrada[0])
                        if pix.alpha or pix.n > 3:
                            pix = fitz.Pixmap(fitz.csRGB, pix) if pix.n - pix.alpha > 3 else pix
                        img = _abrir(pix.tobytes("png"))
                        # Aplica a máscara de transparência (SMask) quando houver.
                        smask = entrada[1]
                        if img and smask:
                            mascara = _abrir(fitz.Pixmap(doc, smask).tobytes("png"))
                            if mascara and mascara.size == img.size:
                                img = img.convert("RGB")
                                img.putalpha(mascara.convert("L"))
                        if img:
                            saida.append(img)
                    except Exception:
                        continue
    elif sufixo == ".docx":
        with zipfile.ZipFile(caminho) as z:
            usados = set()
            for n in z.namelist():
                if n.endswith(".rels"):
                    usados.update("word/media/" + m for m in re.findall(r'media/([^"]+)"', z.read(n).decode("utf-8", "ignore")))
            for n in z.namelist():
                if n in usados:
                    img = _abrir(z.read(n))
                    if img:
                        saida.append(img)
    elif sufixo in (".html", ".htm"):
        texto = caminho.read_text(encoding="utf-8", errors="ignore")
        for b64 in re.findall(r"data:image/[a-z+]+;base64,([A-Za-z0-9+/=\s]+)", texto):
            img = _abrir(base64.b64decode(re.sub(r"\s", "", b64)))
            if img:
                saida.append(img)
    return saida


def marcas_do_arquivo(caminho: Path) -> list[str]:
    return [m for m in (classificar(i) for i in imagens_do_arquivo(caminho)) if m]
