import re
import calendar
from datetime import datetime, date


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


def filtrar_lancamentos_rt_tercas(lancamentos: list[dict]) -> list[dict]:
    """Filtra lançamentos mantendo apenas aqueles com data válida em terça-feira.

    Regras:
    - aceita lista de lançamentos; retorna [] para None ou lista vazia
    - lê a data nas chaves, na prioridade: 'data', 'data_visita', 'Data'
    - normaliza a data usando `normalizar_data_visita`
    - mantém somente datas válidas cujo weekday() seja 1 (terça)
    - preserva a ordem original
    - ignora itens que não sejam dicionários
    - ignora datas vazias ou inválidas
    - não modifica os dicionários recebidos
    """
    if not lancamentos:
        return []

    resultado: list[dict] = []
    for item in lancamentos:
        if not isinstance(item, dict):
            continue

        # Prioridade de chaves
        if "data" in item:
            valor = item.get("data")
        elif "data_visita" in item:
            valor = item.get("data_visita")
        else:
            valor = item.get("Data") if "Data" in item else None

        data_norm = normalizar_data_visita(valor)
        if not data_norm:
            continue

        try:
            dt = datetime.strptime(data_norm, "%d/%m/%Y")
        except Exception:
            continue

        if dt.weekday() == 1:
            resultado.append(item)

    return resultado


def _eh_verdadeiro_equivalente(valor) -> bool:
    if isinstance(valor, bool):
        return valor is True
    if valor is None:
        return False
    texto = str(valor).strip().casefold()
    return texto in {"true", "1", "sim"}


def filtrar_lancamentos_visitas_rt(lancamentos: list[dict]) -> list[dict]:
    """Filtra apenas lançamentos de RT para alimentar tabela de parâmetros do relatório mensal."""
    if not lancamentos:
        return []

    resultado: list[dict] = []
    for item in lancamentos:
        if not isinstance(item, dict):
            continue

        visita_rt_semanal = item.get("visita_rt_semanal")
        tipo_visita = item.get("tipo_visita")

        if _eh_verdadeiro_equivalente(visita_rt_semanal):
            resultado.append(item)
            continue

        if isinstance(tipo_visita, str) and tipo_visita.strip().casefold() == "rt semanal":
            resultado.append(item)
            continue

    return resultado


def gerar_datas_tercas_mes(mes: str, ano: str) -> list[str]:
    """Gera todas as terças-feiras do mês/ano fornecidos no formato dd/mm/aaaa.

    - `mes`: string numérica de 1 a 12 (aceita '7' ou '07')
    - `ano`: string com 4 dígitos
    Retorna lista ordenada cronologicamente. Para entradas inválidas retorna [].
    """
    try:
        if mes is None or ano is None:
            return []
        mes_int = int(str(mes).zfill(2))
        ano_str = str(ano)
        if len(ano_str) != 4 or not ano_str.isdigit():
            return []
        ano_int = int(ano_str)
        if mes_int < 1 or mes_int > 12:
            return []
    except Exception:
        return []

    _, dias_no_mes = calendar.monthrange(ano_int, mes_int)
    resultado: list[str] = []
    for d in range(1, dias_no_mes + 1):
        try:
            dt = date(ano_int, mes_int, d)
        except Exception:
            continue
        if dt.weekday() == 1:  # terça-feira
            resultado.append(dt.strftime("%d/%m/%Y"))
    return resultado
