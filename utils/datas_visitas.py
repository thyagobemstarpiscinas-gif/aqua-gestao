import re
from datetime import datetime


def normalizar_data_visita(valor) -> str:
    """Converte datas como 17/04/26, 170426, 2026-04-17 para dd/mm/aaaa."""
    texto = str(valor or "").strip()
    if not texto:
        return ""

    formatos = [
        "%d/%m/%Y", "%d/%m/%y",
        "%d-%m-%Y", "%d-%m-%y",
        "%Y-%m-%d",
        "%d%m%Y", "%d%m%y",
    ]
    for fmt in formatos:
        try:
            dt = datetime.strptime(texto, fmt)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass

    digitos = re.sub(r"\D", "", texto)
    if len(digitos) == 6:
        try:
            dt = datetime.strptime(digitos, "%d%m%y")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
    if len(digitos) == 8:
        for fmt in ("%d%m%Y", "%Y%m%d"):
            try:
                dt = datetime.strptime(digitos, fmt)
                return dt.strftime("%d/%m/%Y")
            except Exception:
                pass

    return texto


def lancamento_pertence_mes_ano(data_lancamento: str, mes: str, ano: str) -> bool:
    """Confere se uma visita pertence ao mês/ano do relatório."""
    data_norm = normalizar_data_visita(data_lancamento)
    try:
        dt = datetime.strptime(data_norm, "%d/%m/%Y")
        mes_int = int(str(mes).zfill(2))
        ano_int = int(str(ano))
    except Exception:
        return False
    return dt.month == mes_int and dt.year == ano_int
