from datetime import date, datetime


def hoje_br() -> str:
    return date.today().strftime("%d/%m/%Y")


def formatar_data_hora_arquivo(ts: float) -> str:
    dt = datetime.fromtimestamp(ts)
    return dt.strftime("%d/%m/%Y %H:%M")


def parse_data_br(texto: str):
    try:
        return datetime.strptime((texto or "").strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def formatar_data_br(dt: date) -> str:
    return dt.strftime("%d/%m/%Y")


def adicionar_um_ano(dt: date) -> date:
    try:
        return dt.replace(year=dt.year + 1)
    except ValueError:
        return dt.replace(month=2, day=28, year=dt.year + 1)
