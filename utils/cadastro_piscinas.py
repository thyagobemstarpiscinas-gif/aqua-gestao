"""Metadados de piscinas na coluna P de Clientes; A:O permanece compatível."""
import json
import math

TIPOS = ("Não informado", "Adulto", "Infantil", "Family", "SPA", "Raia", "Recreativa", "Outra")
SLOTS = ("adulto", "infantil", "family", "extra1", "extra2")


def volume(valor):
    texto = str(valor if valor is not None else "").strip()
    if not texto:
        return 0.0
    try:
        numero = float(texto.replace(",", "."))
    except (TypeError, ValueError):
        raise ValueError("Informe um volume numérico em m³.") from None
    if not math.isfinite(numero) or numero < 0:
        raise ValueError("O volume deve ser finito e não negativo.")
    return numero


def ler_metadados(texto):
    if not str(texto or "").strip():
        return None
    try:
        dados = json.loads(texto)
    except (ValueError, TypeError):
        raise ValueError("Coluna P contém dados de piscinas inválidos; conteúdo preservado.") from None
    if not isinstance(dados, dict) or dados.get("schema") != "acqua.piscinas.v1":
        raise ValueError("Coluna P ocupada por conteúdo não reconhecido; gravação bloqueada.")
    if not isinstance(dados.get("tipos"), dict) or not isinstance(dados.get("extras"), list):
        raise ValueError("Cadastro de piscinas remoto inválido.")
    if any(k not in SLOTS or v not in TIPOS for k, v in dados["tipos"].items()):
        raise ValueError("Tipo de piscina remoto inválido.")
    for extra in dados["extras"]:
        if not isinstance(extra, dict) or not str(extra.get("nome") or "").strip():
            raise ValueError("Identificação de piscina extra inválida.")
        volume(extra.get("vol"))
    return dados


def serializar(tipos, extras):
    dados = {"schema": "acqua.piscinas.v1", "tipos": dict(tipos), "extras": extras}
    texto = json.dumps(dados, ensure_ascii=False, allow_nan=False)
    ler_metadados(texto)
    if len(texto) > 40000:
        raise ValueError("Cadastro de piscinas excede o limite de armazenamento.")
    return texto
