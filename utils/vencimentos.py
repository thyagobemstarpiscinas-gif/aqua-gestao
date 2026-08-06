from datetime import date, timedelta

from utils.datas import adicionar_um_ano, parse_data_br


def calcular_renovacao_anual(data_fim_texto: str):
    fim_atual = parse_data_br(data_fim_texto)
    if not fim_atual:
        return None, None

    novo_inicio = fim_atual + timedelta(days=1)
    novo_fim = adicionar_um_ano(novo_inicio) - timedelta(days=1)
    return novo_inicio, novo_fim


def status_vencimento(data_fim_texto: str, alerta_dias: int = 30):
    fim = parse_data_br(data_fim_texto)
    if not fim:
        return {
            "codigo": "indefinido",
            "rotulo": "Sem vigência válida",
            "mensagem": "Data final ausente ou inválida.",
            "dias": None,
            "css": "status-indefinido",
        }

    hoje = date.today()
    dias = (fim - hoje).days

    if dias < 0:
        return {
            "codigo": "vencido",
            "rotulo": "Vencido",
            "mensagem": f"Contrato vencido há {abs(dias)} dia(s).",
            "dias": dias,
            "css": "status-vencido",
        }

    if dias <= alerta_dias:
        return {
            "codigo": "vencendo",
            "rotulo": "Vence em breve",
            "mensagem": f"Contrato vence em {dias} dia(s).",
            "dias": dias,
            "css": "status-vencendo",
        }

    return {
        "codigo": "vigente",
        "rotulo": "Vigente",
        "mensagem": f"Contrato vigente. Restam {dias} dia(s) para o vencimento.",
        "dias": dias,
        "css": "status-vigente",
    }


def texto_dias_restantes(status: dict) -> str:
    dias = status.get("dias")
    if dias is None:
        return "Dias restantes: não disponível"
    if dias < 0:
        return f"Atrasado há {abs(dias)} dia(s)"
    return f"Restam {dias} dia(s)"
