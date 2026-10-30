"""Últimos valores finais duráveis; nenhuma conexão é aberta ao importar.

Produção usa uma aba própria, sem modificar abas ou rascunhos existentes.
Cada relatório final acrescenta um registro; autosaves e prévias não chamam
este módulo. Leituras em produção nunca recaem silenciosamente no disco.
"""
import hashlib
import json
import os
from datetime import datetime

from utils import rt_contexto

ABA = "_RT_Ultimos_Valores"
CABECALHO = ["Identidade", "Gerado em", "Dados JSON", "SHA256"]
LIMITE_JSON = 45000


def identidade(chave):
    empresa, cliente, _ = rt_contexto.identidade_ultimos(chave)
    return hashlib.sha256(json.dumps([empresa, cliente], ensure_ascii=False).encode()).hexdigest()


def _local():
    return os.environ.get("CODESPACES", "").strip().lower() == "true"


def _aba(conectar, criar=False):
    planilha = conectar()
    if planilha is None:
        raise ConnectionError("Planilha indisponível; últimos valores não confirmados na nuvem.")
    try:
        aba = planilha.worksheet(ABA)
    except Exception as erro:
        # Apenas ausência explícita permite criação. Falhas de rede/permissão
        # não são tratadas como uma aba inexistente.
        from gspread.exceptions import WorksheetNotFound
        if not isinstance(erro, WorksheetNotFound):
            raise
        if not criar:
            raise FileNotFoundError("Ainda não há últimos valores salvos na nuvem.") from erro
        try:
            aba = planilha.add_worksheet(title=ABA, rows=100, cols=4)
        except Exception:
            # Outra sessão pode ter criado a aba; não alterar seu conteúdo.
            aba = planilha.worksheet(ABA)
        else:
            aba.update(range_name="A1:D1", values=[CABECALHO], value_input_option="RAW")
    return aba


def _linhas(aba):
    linhas = aba.get_all_values()
    if not linhas or linhas[0] != CABECALHO:
        raise ValueError("Cabeçalho de últimos valores inválido; nenhuma alteração realizada.")
    return linhas[1:]


def _registro(row, chave):
    if len(row) != 4:
        raise ValueError("Registro remoto incompleto.")
    bruto = row[2]
    if hashlib.sha256(bruto.encode()).hexdigest() != row[3]:
        raise ValueError("Integridade dos últimos valores não confirmada.")
    dados = rt_contexto.validar_ultimos(json.loads(bruto), chave)
    horario = datetime.fromisoformat(dados["gerado_em"])
    if horario.tzinfo is None or row[1] != dados["gerado_em"]:
        raise ValueError("Horário remoto inválido.")
    return dados, horario


def carregar_ultimos(raiz, chave, conectar):
    if _local():
        return rt_contexto.carregar_ultimos(raiz, chave)
    alvo = identidade(chave)
    candidatos = []
    for pos, row in enumerate(_linhas(_aba(conectar))):
        if row and row[0] == alvo:
            dados, horario = _registro(row, chave)
            candidatos.append((horario, pos, dados))
    if not candidatos:
        raise FileNotFoundError("Nenhum relatório final confirmado para esta empresa e cliente.")
    return max(candidatos, key=lambda item: (item[0], item[1]))[2]


def registrar_ultimos(raiz, chave, snapshot, conectar):
    if _local():
        return rt_contexto.registrar_ultimos(raiz, chave, snapshot)
    dados = rt_contexto.montar_ultimos(chave, snapshot)
    bruto = json.dumps(dados, ensure_ascii=False, sort_keys=True, allow_nan=False)
    if len(bruto) > LIMITE_JSON:
        raise ValueError("Últimos valores excedem o tamanho permitido; registro anterior preservado.")
    row = [identidade(chave), dados["gerado_em"], bruto, hashlib.sha256(bruto.encode()).hexdigest()]
    aba = _aba(conectar, criar=True)
    _linhas(aba)  # validar a estrutura antes de escrever
    # Inclusão, sem sobrescrever valores de outra sessão ou o último válido.
    aba.append_row(row, value_input_option="RAW")
    if row not in _linhas(aba):
        raise IOError("Gravação dos últimos valores sem confirmação; consulte novamente antes de reutilizar.")
