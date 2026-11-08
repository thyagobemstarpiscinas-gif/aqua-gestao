"""Regras do relatório RT da ACQUA GESTÃO, sem Streamlit e sem conexões externas.

Reúne: situação da ART, piscinas do cadastro, identificação das análises e das
fotografias por piscina, leitura numérica das análises, gráficos por piscina e
consolidação trimestral/semestral/anual. Nada aqui grava em Sheets ou Drive.
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path

from utils.cadastro_piscinas import ler_metadados, volume as volume_cadastro

MARCA = "ACQUA GESTÃO"
MARCA_COMPLETA = "ACQUA GESTÃO – Controle Técnico de Piscinas"

# ── ART ──────────────────────────────────────────────────────────────────────
ART_NAO_EMITIDA = "Não emitida"
ART_EM_EMISSAO = "Em processo de emissão"
ART_EMITIDA = "Emitida"
ART_OPCOES = (ART_NAO_EMITIDA, ART_EM_EMISSAO, ART_EMITIDA)
ART_PADRAO = ART_EM_EMISSAO
# Valores gravados por versões anteriores (rascunhos, operador RT). Nenhum deles
# é convertido em "Emitida": a emissão só existe quando informada explicitamente.
_ART_LEGADO = {
    "em processo de formalizacao": ART_EM_EMISSAO,
    "em processo de emissao": ART_EM_EMISSAO,
    "em tramitacao": ART_EM_EMISSAO,
    "em tramitacao administrativa": ART_EM_EMISSAO,
    "nao emitida": ART_NAO_EMITIDA,
    "emitida": ART_EMITIDA,
}


def _sem_acento(texto) -> str:
    mapa = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ",
                         "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC")
    return str(texto or "").translate(mapa).strip().lower()


def normalizar_status_art(valor) -> str:
    """Status canônico. Vazio ou desconhecido vira "Em processo de emissão"."""
    return _ART_LEGADO.get(_sem_acento(valor), ART_PADRAO)


def _data_br(texto):
    texto = str(texto or "").strip()
    if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", texto):
        return None
    try:
        return datetime.strptime(texto, "%d/%m/%Y").date()
    except ValueError:
        return None


def validar_art(status, numero="", inicio="", fim="") -> list[str]:
    """Erros comuns à prévia e à geração final."""
    status = normalizar_status_art(status)
    numero, inicio, fim = (str(v or "").strip() for v in (numero, inicio, fim))
    if numero.upper() == "N/A":
        numero = ""
    if inicio.upper() == "N/A":
        inicio = ""
    if fim.upper() == "N/A":
        fim = ""
    erros = []
    if status != ART_EMITIDA:
        return erros
    if not numero:
        erros.append("ART emitida exige o número. Informe-o ou altere a situação para Não emitida / Em processo de emissão.")
    if inicio or fim:
        d_ini, d_fim = _data_br(inicio), _data_br(fim)
        if inicio and not d_ini:
            erros.append("Vigência da ART — início inválido (use dd/mm/aaaa).")
        if fim and not d_fim:
            erros.append("Vigência da ART — fim inválido (use dd/mm/aaaa).")
        if bool(inicio) != bool(fim):
            erros.append("Informe início e fim da vigência da ART, ou deixe ambos em branco.")
        if d_ini and d_fim and d_fim < d_ini:
            erros.append("Vigência da ART: a data final é anterior à inicial.")
    return erros


def texto_art(status, numero="", inicio="", fim="") -> str:
    """Frase do documento. Número e vigência aparecem somente quando informados."""
    status = normalizar_status_art(status)
    numero, inicio, fim = (str(v or "").strip() for v in (numero, inicio, fim))
    if status == ART_EMITIDA:
        if not numero or numero.upper() == "N/A":
            return "ART emitida — número não informado"
        if inicio and fim and inicio.upper() != "N/A" and fim.upper() != "N/A":
            return f"ART nº {numero} | Vigência: {inicio} a {fim}"
        return f"ART nº {numero} | Vigência não informada"
    if status == ART_EM_EMISSAO:
        return "ART em processo de emissão — número e vigência ainda não disponíveis"
    return "ART não emitida até a data de emissão deste relatório"


# ── Piscinas do cadastro ─────────────────────────────────────────────────────
_SLOTS_PADRAO = (("adulto", "Piscina Adulto", "vol_adulto"),
                 ("infantil", "Piscina Infantil", "vol_infantil"),
                 ("family", "Piscina Family", "vol_family"))
SEM_PISCINA = "Piscina não identificada"


def _vol(valor):
    try:
        v = volume_cadastro(valor)
    except ValueError:
        return None
    return v if v > 0 else None


def piscinas_do_cliente(cliente: dict | None) -> dict:
    """Piscinas com identificação, tipo e volume individual.

    Cadastros antigos: tipo "Não informado"; um volume total sem distribuição é
    informado à parte (sem_distribuicao) e nunca atribuído a uma piscina.
    """
    cliente = cliente or {}
    tipos = dict(cliente.get("tipos_piscinas") or {})
    extras = cliente.get("piscinas_extras") or []
    meta = cliente.get("cadastro_piscinas")
    if isinstance(meta, str):
        try:
            meta = ler_metadados(meta)
        except ValueError:
            meta = None
    if isinstance(meta, dict):
        tipos = dict(meta.get("tipos") or tipos)
        extras = meta.get("extras") or extras
    piscinas = []
    for slot, rotulo, campo in _SLOTS_PADRAO:
        v = _vol(cliente.get(campo))
        tipo = str(tipos.get(slot) or "Não informado")
        if v is not None or tipo != "Não informado":
            piscinas.append({"slot": slot, "nome": rotulo, "tipo": tipo, "volume": v})
    for idx, extra in enumerate(extras):
        if not isinstance(extra, dict):
            continue
        nome = str(extra.get("nome") or "").strip()
        if not nome:
            continue
        slot = f"extra{idx + 1}"
        tipo = str(extra.get("tipo") or tipos.get(slot) or "Não informado")
        piscinas.append({"slot": slot, "nome": nome, "tipo": tipo, "volume": _vol(extra.get("vol"))})
    total = _vol(cliente.get("vol_total"))
    soma = sum(p["volume"] or 0 for p in piscinas)
    sem_distribuicao = bool(total) and not soma
    return {"piscinas": piscinas, "volume_total": total if total else (soma or None),
            "sem_distribuicao": sem_distribuicao}


def formatar_volume(v) -> str:
    if v is None or v == "":
        return "Não informado"
    v = float(v)
    txt = f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if txt.endswith(",0"):
        txt = txt[:-2]
    return f"{txt} m³"


# ── Leituras ─────────────────────────────────────────────────────────────────
# Limites aceitáveis da ABNT NBR 10818:2026 (Tabela 3, piscina coletiva,
# tratamento clorado simples), obtidos de utils.normas_piscinas. Cloro total não
# tem limite próprio (avalia-se CC = CT − CRL) e por isso não recebe faixa.
from utils.normas_piscinas import fora_limite as _fora_limite, limites as _limites, texto_limite as _texto_limite


def _lim(campo):
    mn, mx = _limites(campo)
    return (0.0 if mn is None else mn), mx


PARAMETROS = (
    # chave, rótulo, mínimo, máximo, unidade
    ("ph", "pH", *_lim("ph"), ""),
    ("cloro_livre", "Cloro residual livre (CRL)", *_lim("crl"), "mg/L"),
    ("cloro_total", "Cloro total", None, None, "mg/L"),
    ("alcalinidade", "Alcalinidade total", *_lim("alcalinidade"), "mg/L CaCO3"),
    ("dureza", "Dureza cálcica", *_lim("dureza"), "mg/L CaCO3"),
    ("cianurico", "Ácido cianúrico", *_lim("cianurico"), "mg/L"),
)


def tem_limite(chave) -> bool:
    return chave != "cloro_total" and _limites(chave) != (None, None)


def dentro_limite(chave, valor) -> bool:
    return not _fora_limite(chave, valor)


def texto_faixa(chave) -> str:
    return _texto_limite(chave) if tem_limite(chave) else "sem limite próprio"


def leitura(texto):
    """Primeiro número da célula (leitura inicial quando há "antes / após")."""
    m = re.search(r"-?\d+(?:[.,]\d+)?", str(texto or ""))
    if not m:
        return None
    try:
        v = float(m.group(0).replace(",", "."))
    except ValueError:
        return None
    return v if math.isfinite(v) else None


def linha_tem_leitura(item: dict) -> bool:
    return any(leitura(item.get(k)) is not None for k, *_ in PARAMETROS)


def agrupar_por_piscina(analises: list[dict]) -> list[tuple[str, list[dict]]]:
    """Linhas com leitura agrupadas pela piscina informada, na ordem de aparição."""
    grupos: dict[str, list[dict]] = {}
    for item in analises or []:
        if not linha_tem_leitura(item):
            continue
        nome = str(item.get("piscina") or "").strip() or SEM_PISCINA
        grupos.setdefault(nome, []).append(item)
    return list(grupos.items())


def linhas_de_lancamentos(lancamentos: list[dict]) -> list[dict]:
    """Uma linha por piscina de cada visita (antes só a primeira piscina era usada)."""
    linhas = []
    for lc in lancamentos or []:
        piscinas = [p for p in (lc.get("piscinas") or []) if isinstance(p, dict)]
        for p in piscinas or [None]:
            dados = p if p is not None else lc
            nome = str((p or {}).get("nome") or "").strip() if p is not None else str(lc.get("piscina") or "").strip()
            linhas.append({"lancamento": lc, "dados": dados, "piscina": nome})
    return linhas


# ── Fotografias ──────────────────────────────────────────────────────────────
LEGENDA_MAX = 200


def metadados_foto(foto: dict) -> tuple[str, str]:
    return (str(foto.get("legenda") or "").strip()[:LEGENDA_MAX],
            str(foto.get("piscina") or "").strip()[:120])


def aplicar_metadados(fotos: list[dict], metas: dict) -> bool:
    """Atualiza legenda/piscina por SHA-256. Retorna True se algo mudou."""
    mudou = False
    for f in fotos:
        if f.get("sha256") not in metas:
            continue
        legenda, piscina = metas[f["sha256"]]
        legenda = str(legenda or "").strip()[:LEGENDA_MAX]
        piscina = str(piscina or "").strip()[:120]
        if (f.get("legenda") or "") != legenda or (f.get("piscina") or "") != piscina:
            f["legenda"], f["piscina"] = legenda, piscina
            mudou = True
    return mudou


def legenda_foto(idx: int, foto_meta: dict | None) -> str:
    legenda, piscina = metadados_foto(foto_meta or {})
    partes = [f"Foto {idx}"]
    if piscina:
        partes.append(piscina)
    partes.append(legenda or ("Vista geral da piscina no período de referência" if idx == 1
                              else "Registro complementar da área da piscina"))
    return " — ".join(partes)


# ── Datas e períodos ─────────────────────────────────────────────────────────
def data_leitura(texto):
    texto = str(texto or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(texto[:10], fmt).date()
        except ValueError:
            continue
    return None


PERIODOS = {"Trimestral": 3, "Semestral": 6, "Anual": 12}


def meses_do_periodo(tipo: str, indice: int, ano: int) -> list[tuple[int, int]]:
    """Trimestral 1–4, Semestral 1–2, Anual (índice ignorado)."""
    if tipo not in PERIODOS:
        raise ValueError("Tipo de período inválido.")
    ano = int(ano)
    if not 2000 <= ano <= 2100:
        raise ValueError("Ano inválido.")
    tamanho = PERIODOS[tipo]
    indice = 1 if tipo == "Anual" else int(indice)
    if not 1 <= indice <= 12 // tamanho:
        raise ValueError("Período inválido.")
    inicio = (indice - 1) * tamanho + 1
    return [(m, ano) for m in range(inicio, inicio + tamanho)]


def rotulo_periodo(tipo: str, indice: int, ano: int) -> str:
    if tipo == "Trimestral":
        return f"{int(indice)}º trimestre de {ano}"
    if tipo == "Semestral":
        return f"{int(indice)}º semestre de {ano}"
    return f"Ano de {ano}"


def consolidar(linhas: list[dict], meses: list[tuple[int, int]]) -> dict:
    """Estatística por piscina e parâmetro, com médias mensais para histórico.

    linhas: dicts com "data" (dd/mm/aaaa), "piscina" e os parâmetros.
    Linhas fora do período ou sem data válida são descartadas (contadas).
    """
    permitidos = set(meses)
    resultado: dict[str, dict] = {}
    descartadas = 0
    for item in linhas or []:
        d = data_leitura(item.get("data"))
        if not d or (d.month, d.year) not in permitidos:
            descartadas += 1
            continue
        if not linha_tem_leitura(item):
            continue
        nome = str(item.get("piscina") or "").strip() or SEM_PISCINA
        grupo = resultado.setdefault(nome, {"linhas": [], "parametros": {}, "mensal": {}})
        grupo["linhas"].append({**item, "_data": d})
    for nome, grupo in resultado.items():
        grupo["linhas"].sort(key=lambda r: r["_data"])
        for chave, rotulo, mn, mx, unidade in PARAMETROS:
            valores = [(r["_data"], leitura(r.get(chave))) for r in grupo["linhas"]]
            valores = [(d, v) for d, v in valores if v is not None]
            if not valores:
                continue
            nums = [v for _, v in valores]
            dentro = sum(1 for v in nums if dentro_limite(chave, v))
            grupo["parametros"][chave] = {
                "rotulo": rotulo, "unidade": unidade, "faixa": (mn, mx), "n": len(nums),
                "min": min(nums), "max": max(nums), "media": sum(nums) / len(nums),
                "conformidade": (100.0 * dentro / len(nums)) if tem_limite(chave) else None,
            }
            for m, a in meses:
                mes_vals = [v for d, v in valores if (d.month, d.year) == (m, a)]
                if mes_vals:
                    grupo["mensal"].setdefault(f"{m:02d}/{a}", {})[chave] = sum(mes_vals) / len(mes_vals)
    return {"piscinas": resultado, "descartadas": descartadas, "meses": [f"{m:02d}/{a}" for m, a in meses]}


def fmt_num(v, casas=2) -> str:
    if v is None:
        return "—"
    txt = f"{v:.{casas}f}".replace(".", ",")
    return txt


# ── Gráficos (PIL; PNG usado no PDF e no DOCX) ──────────────────────────────
def _fonte(tam: int, bold: bool = False):
    from PIL import ImageFont
    for caminho in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
                    "/System/Library/Fonts/Supplemental/Arial.ttf"):
        try:
            return ImageFont.truetype(caminho, tam)
        except Exception:
            continue
    return ImageFont.load_default()


def gerar_grafico(series: list[tuple[str, list[float | None]]] | None, rotulos_x: list[str], titulo: str,
                  subtitulo: str, destino: Path, parametros=PARAMETROS) -> Path | None:
    """Painéis por parâmetro com o limite aceitável da NBR 10818:2026 em verde."""
    from PIL import Image, ImageDraw
    paineis = [(p, dict(series or {}).get(p[0])) for p in parametros]
    paineis = [(p, vals) for p, vals in paineis if vals and any(v is not None for v in vals)]
    if not paineis or not rotulos_x:
        return None
    colunas = 2 if len(paineis) > 1 else 1
    linhas = math.ceil(len(paineis) / colunas)
    largura, alt_painel, topo = 1600, 340, 150
    altura = topo + linhas * alt_painel + 60
    img = Image.new("RGB", (largura, altura), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    azul, dourado, texto, cinza, grid = (11, 31, 58), (176, 135, 40), (26, 42, 74), (110, 124, 148), (215, 222, 232)
    faixa, linha_cor, fora = (230, 245, 234), (29, 120, 168), (180, 35, 24)
    draw.rectangle((0, 0, largura, 110), fill=azul)
    draw.text((50, 26), titulo, fill=(255, 255, 255), font=_fonte(36, True))
    draw.text((50, 72), subtitulo, fill=(222, 196, 120), font=_fonte(20))
    f_eixo, f_rot = _fonte(16), _fonte(20, True)
    larg_painel = (largura - 60) // colunas
    for i, ((chave, rotulo, mn, mx, unidade), vals) in enumerate(paineis):
        cx, cy = 30 + (i % colunas) * larg_painel, topo + (i // colunas) * alt_painel
        x0, x1, y0, y1 = cx + 80, cx + larg_painel - 30, cy + 50, cy + alt_painel - 70
        nums = [v for v in vals if v is not None]
        com_faixa = mn is not None and mx is not None and tem_limite(chave)
        vmin, vmax = min(nums + ([mn] if com_faixa else [])), max(nums + ([mx] if com_faixa else []))
        pad = max((vmax - vmin) * 0.15, 0.1)
        vmin, vmax = max(0.0, vmin - pad), vmax + pad

        def ym(v):
            return y1 - (v - vmin) / (vmax - vmin) * (y1 - y0)
        draw.text((cx + 10, cy + 8), f"{rotulo}" + (f" ({unidade})" if unidade else ""), fill=texto, font=f_rot)
        if com_faixa:
            draw.text((x1 - 300, cy + 12), f"Limite NBR 10818:2026: {texto_faixa(chave)}", fill=(34, 120, 70), font=f_eixo)
            draw.rectangle((x0, ym(mx), x1, ym(mn)), fill=faixa)
        else:
            draw.text((x1 - 300, cy + 12), "Sem limite próprio (avaliar CC)", fill=cinza, font=f_eixo)
        for g in range(5):
            y = y0 + (y1 - y0) * g / 4
            draw.line((x0, y, x1, y), fill=grid)
            draw.text((cx + 10, y - 9), fmt_num(vmax - (vmax - vmin) * g / 4, 1), fill=cinza, font=f_eixo)
        n = len(rotulos_x)
        xs = [x0 + (x1 - x0) * (k / max(n - 1, 1)) if n > 1 else (x0 + x1) / 2 for k in range(n)]
        passo = max(1, math.ceil(n / 10))
        for k, rot in enumerate(rotulos_x):
            if k % passo == 0 or k == n - 1:
                draw.text((xs[k] - 22, y1 + 10), rot, fill=cinza, font=f_eixo)
        pts = [(xs[k], ym(v), v) for k, v in enumerate(vals) if v is not None]
        if len(pts) > 1:
            draw.line([(x, y) for x, y, _ in pts], fill=linha_cor, width=4)
        for x, y, v in pts:
            cor = fora if com_faixa and not dentro_limite(chave, v) else linha_cor
            draw.ellipse((x - 7, y - 7, x + 7, y + 7), fill=cor, outline=(255, 255, 255), width=2)
    draw.text((50, altura - 45), "Faixa verde: limite aceitável da ABNT NBR 10818:2026 (Tabela 3). Pontos vermelhos: leitura fora do limite.",
              fill=cinza, font=f_eixo)
    draw.line((0, 112, largura, 112), fill=dourado, width=4)
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    img.save(destino, format="PNG", optimize=True)
    return destino


def graficos_por_piscina(analises: list[dict], pasta: Path, nome_condominio: str, periodo: str) -> list[tuple[str, Path]]:
    """Um gráfico de evolução por piscina, a partir das análises do relatório mensal."""
    saida = []
    for idx, (nome, itens) in enumerate(agrupar_por_piscina(analises), start=1):
        ordenados = sorted(enumerate(itens), key=lambda t: (data_leitura(t[1].get("data")) or date.max, t[0]))
        itens = [it for _, it in ordenados]
        rotulos = [str(it.get("data") or f"#{k + 1}")[:5] for k, it in enumerate(itens)]
        series = [(chave, [leitura(it.get(chave)) for it in itens]) for chave, *_ in PARAMETROS]
        p = gerar_grafico(series, rotulos, f"Evolução dos parâmetros — {nome}",
                          f"{nome_condominio} • {periodo}", Path(pasta) / f"grafico_piscina_{idx}.png")
        if p:
            saida.append((nome, p))
    return saida


def graficos_consolidados(consolidado: dict, pasta: Path, nome_condominio: str, periodo: str) -> list[tuple[str, Path]]:
    """Histórico das médias mensais por piscina."""
    saida = []
    meses = consolidado.get("meses") or []
    for idx, (nome, grupo) in enumerate(consolidado.get("piscinas", {}).items(), start=1):
        series = [(chave, [grupo["mensal"].get(m, {}).get(chave) for m in meses]) for chave, *_ in PARAMETROS]
        p = gerar_grafico(series, [m[:2] + "/" + m[-2:] for m in meses], f"Histórico mensal — {nome}",
                          f"{nome_condominio} • {periodo} • médias mensais", Path(pasta) / f"historico_piscina_{idx}.png")
        if p:
            saida.append((nome, p))
    return saida
