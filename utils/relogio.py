"""Relógio explícito para documentos operacionais Bem Star no Cloud."""
from datetime import datetime
from zoneinfo import ZoneInfo


def hoje_brasilia(agora: datetime | None = None) -> str:
    fuso = ZoneInfo("America/Sao_Paulo")
    instante = agora if agora is not None else datetime.now(fuso)
    if instante.tzinfo is None:
        raise ValueError("O instante de referência deve incluir fuso horário.")
    return instante.astimezone(fuso).strftime("%d/%m/%Y")
