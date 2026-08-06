from __future__ import annotations

import io
import re
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from utils.datas_visitas import normalizar_data_visita


RT_NAME = "Thyago Fernando da Silveira"
RT_QUALIFICATION = "Técnico em Química"
RT_CRQ = "CRQ-MG 2ª Região nº 024025748"
RT_ROLE = "Responsável Técnico"
COMPANY_NAME = "Aqua Gestão — Controle Técnico de Piscinas"

MONTH_NAMES_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

RANGES = {
    "ph": (7.2, 7.8),
    "crl": (0.5, 3.0),
    "ct": (0.5, 3.0),
    "alcalinidade": (80.0, 120.0),
    "dureza": (150.0, 300.0),
    "cya": (30.0, 50.0),
}

COLORS = {
    "navy": "#143A63",
    "gold": "#B88A2B",
    "white": "#FFFFFF",
    "light_gray": "#F4F6F8",
    "green_bg": "#EAF7EF",
    "green_text": "#1D6E3A",
    "red_soft": "#FDEBEC",
    "red_text": "#8A1F2B",
    "text": "#1D2939",
    "muted": "#667085",
    "line": "#C9D2DB",
}


def _norm_text(value: Any) -> str:
    text = str(value or "").strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text.casefold()


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.count(",") == 1 and text.count(".") >= 1:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except Exception:
        return None


def _format_number(value: Any, decimals: int = 2) -> str:
    num = _to_float(value)
    if num is None:
        return "—"
    text = f"{num:.{decimals}f}".replace(".", ",")
    text = re.sub(r",00$", "", text)
    return text


def _format_date(value: Any) -> str:
    data = normalizar_data_visita(value)
    try:
        datetime.strptime(data, "%d/%m/%Y")
        return data
    except Exception:
        return ""


def _date_key(value: Any) -> datetime:
    data = _format_date(value)
    if not data:
        return datetime.max
    return datetime.strptime(data, "%d/%m/%Y")


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    norm = _norm_text(value)
    return norm in {"1", "true", "sim", "yes", "rt"}


def normalize_rt_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for item in records or []:
        if not isinstance(item, dict):
            continue

        pools = item.get("piscinas") if isinstance(item.get("piscinas"), list) else []
        pool0 = pools[0] if pools and isinstance(pools[0], dict) else {}

        normalized.append(
            {
                "raw": item,
                "condominio": str(item.get("nome_condominio") or item.get("condominio") or item.get("cliente") or "").strip(),
                "data": _format_date(item.get("data") or item.get("data_visita") or item.get("Data")),
                "observacao": str(item.get("observacao") or item.get("obs") or item.get("anotacoes") or "").strip(),
                "origem": str(item.get("origem") or item.get("origem_lancamento") or item.get("canal") or "").strip(),
                "operador": str(item.get("operador") or item.get("usuario") or item.get("responsavel") or "").strip(),
                "tipo_visita": str(item.get("tipo_visita") or item.get("tipo") or "").strip(),
                "visita_rt_semanal": item.get("visita_rt_semanal"),
                "is_rt": item.get("is_rt"),
                "perfil": str(item.get("perfil") or item.get("perfil_usuario") or "").strip(),
                "ph": item.get("ph") if item.get("ph") not in (None, "") else pool0.get("ph"),
                "crl": item.get("cloro_livre") if item.get("cloro_livre") not in (None, "") else pool0.get("cloro_livre"),
                "ct": item.get("cloro_total") if item.get("cloro_total") not in (None, "") else pool0.get("cloro_total"),
                "alcalinidade": item.get("alcalinidade") if item.get("alcalinidade") not in (None, "") else pool0.get("alcalinidade"),
                "dureza": item.get("dureza") if item.get("dureza") not in (None, "") else pool0.get("dureza"),
                "cya": item.get("cianurico") if item.get("cianurico") not in (None, "") else pool0.get("cianurico"),
                "turbidez": item.get("turbidez") if item.get("turbidez") not in (None, "") else pool0.get("turbidez"),
                "orp": item.get("orp") if item.get("orp") not in (None, "") else pool0.get("orp"),
                "tds": item.get("tds") if item.get("tds") not in (None, "") else pool0.get("tds"),
                "temperatura": item.get("temperatura") if item.get("temperatura") not in (None, "") else pool0.get("temperatura"),
                "foto": item.get("foto") or item.get("foto_path") or item.get("imagem") or "",
                "fotos": item.get("fotos") if isinstance(item.get("fotos"), list) else [],
            }
        )
    return normalized


def select_records_by_condominium(records: list[dict[str, Any]], condominium: str) -> list[dict[str, Any]]:
    if not condominium:
        return []
    target = _norm_text(condominium)
    return [r for r in records if _norm_text(r.get("condominio")) == target]


def select_records_by_month_year(records: list[dict[str, Any]], month: str, year: str) -> list[dict[str, Any]]:
    try:
        m = int(str(month).strip())
        y = int(str(year).strip())
    except Exception:
        return []

    filtered = []
    for row in records:
        data = row.get("data")
        if not data:
            continue
        try:
            dt = datetime.strptime(data, "%d/%m/%Y")
        except Exception:
            continue
        if dt.month == m and dt.year == y:
            filtered.append(row)
    return filtered


def identify_record_origin(record: dict[str, Any]) -> dict[str, Any]:
    text_fields = " | ".join(
        [
            str(record.get("origem", "")),
            str(record.get("tipo_visita", "")),
            str(record.get("perfil", "")),
            str(record.get("operador", "")),
            str((record.get("raw") or {}).get("origem_lancamento", "")),
            str((record.get("raw") or {}).get("acesso_origem", "")),
            str((record.get("raw") or {}).get("usuario_tipo", "")),
        ]
    )
    norm = _norm_text(text_fields)

    has_rt_marker = (
        _truthy(record.get("visita_rt_semanal"))
        or _truthy(record.get("is_rt"))
        or "rt" in _norm_text(record.get("tipo_visita"))
        or "responsavel tecnico" in norm
        or "pin rt" in norm
        or "acesso rt" in norm
        or "origem rt" in norm
    )

    has_operator_marker = (
        "operador" in norm
        or "campo" in norm
        or "pin operador" in norm
        or _norm_text((record.get("raw") or {}).get("is_operador")) in {"1", "true", "sim"}
    )

    return {
        "is_rt": bool(has_rt_marker and not (has_operator_marker and not has_rt_marker)),
        "is_operator": bool(has_operator_marker and not has_rt_marker),
        "reason": "rt_marker" if has_rt_marker else ("operator_marker" if has_operator_marker else "unknown"),
    }


def exclude_operator_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    selected: list[dict[str, Any]] = []
    for row in records:
        origin = identify_record_origin(row)
        if origin["is_rt"]:
            selected.append(row)
    return selected


def sort_records_chronologically(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(records, key=lambda r: _date_key(r.get("data")))


def _status_for_value(field: str, value: Any) -> str:
    num = _to_float(value)
    if num is None:
        return "missing"
    if field not in RANGES:
        return "ok"
    mn, mx = RANGES[field]
    if num < mn or num > mx:
        return "out"
    return "ok"


def prepare_monitoring_table_rows(records: list[dict[str, Any]]) -> list[list[dict[str, str]]]:
    rows: list[list[dict[str, str]]] = []
    for row in records:
        ct = _to_float(row.get("ct"))
        crl = _to_float(row.get("crl"))
        combinado = None
        if ct is not None and crl is not None:
            combinado = max(ct - crl, 0.0)

        entries: list[tuple[str, Any, str]] = [
            ("date", row.get("data"), "date"),
            ("ph", row.get("ph"), "ph"),
            ("crl", row.get("crl"), "crl"),
            ("ct", row.get("ct"), "ct"),
            ("combined", combinado, "combined"),
            ("alcalinidade", row.get("alcalinidade"), "alcalinidade"),
            ("dureza", row.get("dureza"), "dureza"),
            ("cya", row.get("cya"), "cya"),
            ("turbidez", row.get("turbidez"), "turbidez"),
            ("orp", row.get("orp"), "orp"),
            ("tds", row.get("tds"), "tds"),
            ("temperatura", row.get("temperatura"), "temperatura"),
        ]

        out: list[dict[str, str]] = []
        for key, value, range_key in entries:
            if key == "date":
                text = str(value or "—")
                status = "ok"
            else:
                text = _format_number(value)
                status = _status_for_value(range_key, value)
            out.append({"text": text, "status": status})
        rows.append(out)
    return rows


def safe_report_filename(condominium: str, month: str, year: str) -> str:
    base = f"Relatorio_RT_{condominium}_{month}_{year}.pdf"
    base = unicodedata.normalize("NFKD", base)
    base = "".join(ch for ch in base if not unicodedata.combining(ch))
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base)
    base = re.sub(r"_+", "_", base).strip("_")
    return base or "Relatorio_RT.pdf"


def _month_year_text(month: str, year: str) -> str:
    try:
        m = int(month)
        return f"{MONTH_NAMES_PT.get(m, str(m))}/{year}"
    except Exception:
        return f"{month}/{year}"


def _join_pt(items: list[str]) -> str:
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} e {items[1]}"
    return ", ".join(items[:-1]) + f" e {items[-1]}"


def _collect_deviation_data(rows: list[dict[str, Any]]) -> dict[str, list[str]]:
    by_date: dict[str, list[str]] = {}
    labels = {
        "ph": "pH",
        "crl": "CRL",
        "ct": "CT",
        "alcalinidade": "alcalinidade",
        "dureza": "dureza",
        "cya": "CYA",
    }
    for row in rows:
        data = row.get("data") or "—"
        outside: list[str] = []
        for field, label in labels.items():
            if _status_for_value(field, row.get(field)) == "out":
                outside.append(label)
        if outside:
            by_date[data] = outside
    return by_date


def _deviation_sentences(rows: list[dict[str, Any]]) -> list[str]:
    by_date = _collect_deviation_data(rows)
    out: list[str] = []
    for data in sorted(by_date.keys(), key=_date_key):
        fields = by_date[data]
        out.append(f"{data} — {_join_pt(fields)} fora das faixas operacionais.")
    return out


def _recommendations_from_deviations(rows: list[dict[str, Any]]) -> list[str]:
    found = set()
    for row in rows:
        for field in ["ph", "crl", "ct", "alcalinidade", "dureza", "cya"]:
            if _status_for_value(field, row.get(field)) == "out":
                found.add(field)

    recs: list[str] = []
    if "ph" in found:
        recs.append("Ajustar pH para a faixa operacional de 7,2 a 7,8 e repetir leitura após estabilização.")
    if "crl" in found or "ct" in found:
        recs.append("Reavaliar desinfecção (CRL/CT), revisar demanda oxidante e confirmar residual após recirculação.")
    if "alcalinidade" in found:
        recs.append("Corrigir alcalinidade total para melhorar estabilidade do pH e reduzir oscilações operacionais.")
    if "dureza" in found:
        recs.append("Ajustar dureza cálcica para mitigar risco de corrosão ou incrustação.")
    if "cya" in found:
        recs.append("Revisar estabilizante (CYA) e considerar renovação parcial de água quando tecnicamente indicado.")

    if not recs:
        recs.append("Manter rotina de monitoramento técnico com registro das leituras e rastreabilidade operacional.")
    return recs


def _extract_photo_items(records: list[dict[str, Any]]) -> list[dict[str, str]]:
    items: list[dict[str, str]] = []
    for row in records:
        candidates = []
        if row.get("foto"):
            candidates.append(row.get("foto"))
        for fp in row.get("fotos", []):
            candidates.append(fp)
        for c in candidates:
            path = Path(str(c))
            if path.exists() and path.is_file():
                items.append(
                    {
                        "path": str(path),
                        "caption": str((row.get("observacao") or "").strip() or "Registro técnico do período"),
                        "date": str(row.get("data") or ""),
                    }
                )
    return items


def _collect_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
    total = len(rows)
    leituras = 0
    for row in rows:
        if any(
            _to_float(row.get(k)) is not None
            for k in ["ph", "crl", "ct", "alcalinidade", "dureza", "cya", "turbidez", "orp", "tds", "temperatura"]
        ):
            leituras += 1

    deviations = _deviation_sentences(rows)
    if not rows:
        status = "Sem registros técnicos"
        status_flag = "info"
    elif deviations:
        status = "Atenção: há desvios no período"
        status_flag = "attention"
    else:
        status = "Condição estável nas leituras disponíveis"
        status_flag = "ok"

    return {
        "visitas_rt": total,
        "leituras": leituras,
        "status": status,
        "status_flag": status_flag,
        "deviations": deviations,
    }


def _logo_path() -> Path:
    path = Path("assets/branding/aqua_gestao_logo.png")
    if not path.exists() or not path.is_file():
        raise ValueError("Logo oficial não encontrada em assets/branding/aqua_gestao_logo.png. Interrompendo geração do relatório.")
    return path


def _resolve_unicode_fonts() -> tuple[str, str]:
    reg_candidates = [
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
    ]
    bold_candidates = [
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]

    reg = next((p for p in reg_candidates if p.exists()), None)
    bold = next((p for p in bold_candidates if p.exists()), None)
    if not reg or not bold:
        raise ValueError("Fonte Unicode não encontrada no sistema (DejaVu Sans/Liberation Sans).")
    return str(reg), str(bold)


def validate_generation_selection(condominium: str) -> str | None:
    if not str(condominium or "").strip():
        return "Selecione um condomínio para gerar o relatório."
    if _norm_text(condominium) == "todos":
        return "A opção TODOS não pode ser usada para gerar relatório RT em PDF. Selecione um condomínio específico."
    return None


def build_monthly_rt_pdf_bytes(
    records: list[dict[str, Any]],
    condominium: str,
    month: str,
    year: str,
    issue_date: str,
    address: str = "",
    art_number: str = "",
    rt_notes: str = "",
    representative_name: str = "",
) -> bytes:
    error = validate_generation_selection(condominium)
    if error:
        raise ValueError(error)

    normalized = normalize_rt_records(records)
    if not normalized:
        raise ValueError("Não há registros técnicos para processar.")

    by_condo = select_records_by_condominium(normalized, condominium)
    by_period = select_records_by_month_year(by_condo, month, year)
    rt_only = exclude_operator_records(by_period)
    ordered = sort_records_chronologically(rt_only)

    if not ordered:
        raise ValueError("Não há registros técnicos de RT no período selecionado.")

    logo_path = _logo_path()

    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
    from reportlab.platypus import (
        Image as RLImage,
        KeepTogether,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    font_regular_path, font_bold_path = _resolve_unicode_fonts()
    pdfmetrics.registerFont(TTFont("AquaSans", font_regular_path))
    pdfmetrics.registerFont(TTFont("AquaSans-Bold", font_bold_path))

    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states: list[dict[str, Any]] = []
            self.setPageCompression(0)
            self.setTitle("Relatório Mensal de Acompanhamento Técnico")
            self.setAuthor(f"{RT_NAME} — {RT_CRQ}")
            self.setSubject("Acompanhamento técnico mensal da qualidade da água de piscinas")
            self.setKeywords("Aqua Gestão, RT, CRQ, piscina, monitoramento, qualidade da água")

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                _draw_header_footer(self, total)
                super().showPage()
            super().save()

    def _draw_header_footer(c: canvas.Canvas, total_pages: int) -> None:
        w, h = A4
        c.saveState()
        c.setStrokeColor(colors.HexColor(COLORS["line"]))
        c.setLineWidth(0.8)
        c.line(14 * mm, h - 14 * mm, w - 14 * mm, h - 14 * mm)
        c.line(14 * mm, 13 * mm, w - 14 * mm, 13 * mm)

        c.drawImage(str(logo_path), 14 * mm, h - 13 * mm, width=14 * mm, height=10 * mm, preserveAspectRatio=True, mask="auto")

        c.setFont("AquaSans-Bold", 8)
        c.setFillColor(colors.HexColor(COLORS["navy"]))
        c.drawRightString(w - 14 * mm, h - 10.5 * mm, "Relatório Mensal de Acompanhamento Técnico")

        c.setFont("AquaSans", 7)
        c.setFillColor(colors.HexColor(COLORS["muted"]))
        c.drawRightString(w - 14 * mm, h - 13.3 * mm, f"{condominium} | {_month_year_text(month, year)}")

        c.drawString(14 * mm, 9 * mm, COMPANY_NAME)
        c.drawCentredString(w / 2, 9 * mm, f"Emissão: {issue_date} | Documento de acompanhamento técnico do período")
        c.drawRightString(w - 14 * mm, 9 * mm, f"Página {c.getPageNumber()}/{total_pages}")
        c.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=17 * mm,
        bottomMargin=16 * mm,
        title="Relatório Mensal de Acompanhamento Técnico",
        author=RT_NAME,
        subject="Acompanhamento técnico",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleMain", parent=styles["Title"], fontName="AquaSans-Bold", fontSize=18, leading=22, alignment=TA_CENTER, textColor=colors.HexColor(COLORS["navy"]), spaceAfter=2))
    styles.add(ParagraphStyle(name="Subtitle", parent=styles["Normal"], fontName="AquaSans", fontSize=10.5, leading=13.5, alignment=TA_CENTER, textColor=colors.HexColor(COLORS["gold"]), spaceAfter=8))
    styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontName="AquaSans-Bold", fontSize=11, leading=13, textColor=colors.HexColor(COLORS["navy"]), spaceBefore=5, spaceAfter=3))
    styles.add(ParagraphStyle(name="Body", parent=styles["BodyText"], fontName="AquaSans", fontSize=8.3, leading=10.3, alignment=TA_LEFT, textColor=colors.HexColor(COLORS["text"])))
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontName="AquaSans", fontSize=7.7, leading=9.4, alignment=TA_LEFT, textColor=colors.HexColor(COLORS["muted"])))
    styles.add(ParagraphStyle(name="Cell", parent=styles["BodyText"], fontName="AquaSans", fontSize=7.2, leading=9.0, alignment=TA_CENTER, textColor=colors.HexColor(COLORS["text"])))
    styles.add(ParagraphStyle(name="CellHead", parent=styles["BodyText"], fontName="AquaSans-Bold", fontSize=7.1, leading=8.7, alignment=TA_CENTER, textColor=colors.white))

    def P(text: str, style: str = "Body", escape: bool = True) -> Paragraph:
        esc = str(text or "")
        if escape:
            esc = esc.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(esc, styles[style])

    summary = _collect_summary(ordered)
    table_rows = prepare_monitoring_table_rows(ordered)
    deviation_lines = summary["deviations"]
    recommendations = _recommendations_from_deviations(ordered)

    art_line = str(art_number or "").strip() or "Não informada"

    story = []

    story.append(RLImage(str(logo_path), width=44 * mm, height=30 * mm, hAlign="CENTER"))
    story.append(Spacer(1, 3 * mm))
    story.append(P("RELATÓRIO MENSAL DE ACOMPANHAMENTO TÉCNICO", "TitleMain"))
    story.append(P("Controle da Qualidade da Água de Piscinas", "Subtitle"))

    ident_data = [
        ["Condomínio", condominium],
        ["Endereço", address or "Não informado"],
        ["Mês/ano de referência", _month_year_text(month, year)],
        ["Data de emissão", issue_date],
        ["Responsável Técnico", RT_NAME],
        ["Qualificação", RT_QUALIFICATION],
        ["Registro", RT_CRQ],
        ["ART", art_line],
    ]

    ident_tbl = Table([[P(a, "Small"), P(b, "Body")] for a, b in ident_data], colWidths=[52 * mm, 124 * mm], hAlign="LEFT")
    ident_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(COLORS["light_gray"])),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(COLORS["line"])),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(ident_tbl)
    story.append(Spacer(1, 3 * mm))

    resumo_data = [
        ["Resumo Executivo", ""],
        ["Visitas técnicas do RT", str(summary["visitas_rt"])],
        ["Leituras avaliadas", str(summary["leituras"])],
        ["Situação geral", summary["status"]],
        ["Principais desvios", " ; ".join(deviation_lines[:3]) if deviation_lines else "Sem desvios relevantes nas leituras disponíveis."],
    ]
    resumo_tbl = Table([[P(a, "Small"), P(b, "Body")] for a, b in resumo_data], colWidths=[52 * mm, 124 * mm], hAlign="LEFT")
    resumo_style = [
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (1, 0), colors.HexColor(COLORS["navy"])),
        ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
        ("FONTNAME", (0, 0), (1, 0), "AquaSans-Bold"),
        ("ALIGN", (0, 0), (1, 0), "CENTER"),
        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor(COLORS["light_gray"])),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(COLORS["line"])),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]
    if summary["status_flag"] == "attention":
        resumo_style.append(("BACKGROUND", (1, 3), (1, 3), colors.HexColor("#FFF4E5")))
        resumo_style.append(("TEXTCOLOR", (1, 3), (1, 3), colors.HexColor("#8A5800")))
    elif summary["status_flag"] == "ok":
        resumo_style.append(("BACKGROUND", (1, 3), (1, 3), colors.HexColor(COLORS["green_bg"])))
        resumo_style.append(("TEXTCOLOR", (1, 3), (1, 3), colors.HexColor(COLORS["green_text"])))
    resumo_tbl.setStyle(TableStyle(resumo_style))
    story.append(resumo_tbl)

    story.append(Spacer(1, 4 * mm))
    story.append(P("2. Monitoramento Técnico", "H2"))

    header = ["Data", "pH", "CRL", "CT", "CC", "Alc.", "Dureza", "CYA", "Turb. (NTU)", "ORP (mV)", "TDS (ppm)", "Temp. (°C)"]

    body = []
    for row in table_rows:
        row_cells = []
        for idx, cell in enumerate(row):
            txt = cell["text"]
            if idx == 0 and txt not in ("", "—"):
                txt = txt.replace("/", "&#8209;/")
            row_cells.append(P(txt, "Cell"))
        body.append(row_cells)

    monitor_data = [[P(h, "CellHead") for h in header]] + body
    col_widths = [20 * mm, 10.8 * mm, 10.8 * mm, 10.8 * mm, 10.8 * mm, 11.5 * mm, 12.5 * mm, 10.8 * mm, 13 * mm, 12.3 * mm, 12.3 * mm, 12.3 * mm]
    monitor_tbl = Table(monitor_data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")

    style_cmds: list[tuple[Any, ...]] = [
        ("GRID", (0, 0), (-1, -1), 0.30, colors.HexColor(COLORS["line"])),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLORS["navy"])),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.6),
        ("TOPPADDING", (0, 0), (-1, -1), 3.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.6),
    ]

    for r_idx, row in enumerate(table_rows, start=1):
        if r_idx % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, r_idx), (-1, r_idx), colors.HexColor(COLORS["light_gray"])))
        for c_idx, cell in enumerate(row):
            if cell["status"] == "out":
                style_cmds.append(("BACKGROUND", (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(COLORS["red_soft"])))
                style_cmds.append(("TEXTCOLOR", (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(COLORS["red_text"])))

    monitor_tbl.setStyle(TableStyle(style_cmds))
    story.append(monitor_tbl)
    story.append(Spacer(1, 1.8 * mm))
    story.append(P("CC: cloro combinado calculado como CT − CRL.", "Small"))
    story.append(Spacer(1, 2.4 * mm))

    synthesis = "Leituras sem não conformidades críticas no período." if not deviation_lines else "Foram identificados desvios que exigem ação corretiva com rastreabilidade."
    if len(table_rows) <= 2:
        trends = "As leituras do período indicam variação pontual, sem série histórica suficiente para caracterização de tendência."
    else:
        trends = "As leituras do período indicam comportamento técnico consistente para acompanhamento mensal." if not deviation_lines else "As leituras do período indicam necessidade de reavaliação após as ações corretivas."

    observations = rt_notes.strip() or "Sem observações adicionais registradas pelo RT no período."

    story.append(P(f"Síntese dos resultados: {synthesis}", "Body"))
    story.append(P(f"Desvios relevantes: {' ; '.join(deviation_lines) if deviation_lines else 'Não foram observados desvios relevantes.'}", "Body"))
    story.append(P(f"Tendências observadas: {trends}", "Body"))
    story.append(P(f"Observações registradas pelo RT: {observations}", "Body"))

    story.append(Spacer(1, 3 * mm))
    story.append(P("3. Parecer e Recomendações", "H2"))

    def _bullet(title: str, text: str) -> Paragraph:
        return P(f"<b>{title}:</b> {text}", "Body", escape=False)

    conclusion = "No período avaliado, as evidências técnicas registradas sustentam acompanhamento continuado com foco em estabilidade físico-química e rastreabilidade operacional."
    if deviation_lines:
        conclusion = "No período avaliado, foram identificadas não conformidades pontuais e recomenda-se executar plano corretivo com reavaliação técnica sequencial."

    story.append(_bullet("1. Avaliação técnica do período", synthesis))
    story.append(_bullet("2. Não conformidades identificadas", " ; ".join(deviation_lines) if deviation_lines else "Não foram identificadas não conformidades nas leituras registradas."))
    story.append(_bullet("3. Recomendações técnicas", _join_pt(recommendations)))
    story.append(_bullet("4. Prazos sugeridos", "Imediato para desvios críticos e até a próxima rotina para ajustes preventivos."))
    story.append(_bullet("5. Responsável pela ação", "Operação local com acompanhamento do Responsável Técnico."))
    story.append(_bullet("6. Conclusão técnica", conclusion))

    refs = (
        "Referências: Lei Federal nº 2.800/1956; Decreto nº 85.877/1981; "
        "Resolução CFQ nº 332/2025; ABNT NBR 10339:2018 (versão corrigida de 2019). "
        "A Portaria GM/MS nº 888/2021 é referência sanitária complementar para água de consumo humano."
    )
    story.append(Spacer(1, 2.2 * mm))
    story.append(P(refs, "Small"))
    story.append(Spacer(1, 2.2 * mm))

    signature_table = Table(
        [
            [P("Responsável Técnico", "Small"), P("Representante do estabelecimento", "Small")],
            [
                P(
                    f"{RT_NAME}<br/>{RT_QUALIFICATION}<br/>{RT_CRQ}<br/>{RT_ROLE}<br/><br/>Data de emissão: {issue_date}",
                    "Body",
                    escape=False,
                ),
                P("Assinatura: ___________________________<br/>Nome: ________________________________", "Body", escape=False),
            ],
        ],
        colWidths=[90 * mm, 86 * mm],
        hAlign="LEFT",
    )
    signature_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(COLORS["line"])),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(COLORS["light_gray"])),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(signature_table)
    story.append(Spacer(1, 1.6 * mm))
    story.append(P("Aviso: análises microbiológicas dependem de coleta e laboratório competente, quando aplicável.", "Small"))

    photos = _extract_photo_items(ordered)
    if photos:
        story.append(Spacer(1, 4 * mm))
        story.append(P("Registro Fotográfico", "H2"))
        for idx in range(0, len(photos), 2):
            chunk = photos[idx : idx + 2]
            blocks = []
            for item in chunk:
                p = Path(item["path"])
                if not p.exists():
                    continue
                try:
                    from PIL import Image

                    img = Image.open(p)
                    iw, ih = img.size
                    max_w = 82 * mm
                    max_h = 95 * mm
                    ratio = min(max_w / float(iw), max_h / float(ih))
                    w = max(20.0, float(iw) * ratio)
                    h = max(20.0, float(ih) * ratio)
                    blocks.append(
                        KeepTogether(
                            [
                                RLImage(str(p), width=w, height=h, hAlign="CENTER"),
                                Spacer(1, 1.2 * mm),
                                P(f"{item.get('date', '')} — {item.get('caption', '')}", "Small"),
                                Spacer(1, 3 * mm),
                            ]
                        )
                    )
                except Exception:
                    continue

            for block in blocks:
                story.append(block)

    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()


def generate_monthly_rt_pdf(
    records: list[dict[str, Any]],
    condominium: str,
    month: str,
    year: str,
    issue_date: str,
    output_path: str | Path,
    address: str = "",
    art_number: str = "",
    rt_notes: str = "",
    representative_name: str = "",
) -> Path:
    pdf_bytes = build_monthly_rt_pdf_bytes(
        records=records,
        condominium=condominium,
        month=month,
        year=year,
        issue_date=issue_date,
        address=address,
        art_number=art_number,
        rt_notes=rt_notes,
        representative_name=representative_name,
    )
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(pdf_bytes)
    return out
