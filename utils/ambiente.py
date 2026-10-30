"""Seleciona destinos Google sem abrir conexão ou ler credenciais."""
from collections.abc import Mapping
from functools import lru_cache
from pathlib import Path
import re
import subprocess


@lru_cache(maxsize=1)
def registrar_commit_homologacao():
    """Registra somente o SHA local; uma falha de diagnóstico não interrompe o app."""
    try:
        resultado = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=Path(__file__).resolve().parents[1],
            capture_output=True, text=True, timeout=3, check=True,
        )
        sha = resultado.stdout.strip()
        if not re.fullmatch(r"[0-9a-f]{40}", sha):
            raise ValueError("SHA inválido")
    except (OSError, subprocess.SubprocessError, ValueError):
        sha = "indisponivel"
    print(f"AQUA_HOMOLOGACAO_COMMIT={sha}", flush=True)
    return sha


def destinos_google(secrets, sheets_producao, drive_producao):
    try:
        cfg = secrets.get("aqua_app", {})
    except FileNotFoundError:
        cfg = {}
    if not isinstance(cfg, Mapping):
        raise ValueError("Configuração aqua_app inválida.")
    ambiente = cfg.get("ambiente", "producao")
    if ambiente == "producao":
        if cfg.get("sheets_id") or cfg.get("drive_fotos_folder_id"):
            raise ValueError("Destinos alternativos exigem ambiente de homologacao explícito.")
        return ambiente, sheets_producao, drive_producao
    if ambiente != "homologacao":
        raise ValueError("Ambiente inválido; use producao ou homologacao.")
    destinos = []
    for campo, original in (("sheets_id", sheets_producao),
                            ("drive_fotos_folder_id", drive_producao)):
        valor = cfg.get(campo)
        if not isinstance(valor, str) or not re.fullmatch(r"[A-Za-z0-9_-]{10,}", valor):
            raise ValueError("Homologação exige IDs próprios de planilha e pasta Drive.")
        if valor == original:
            raise ValueError("Homologação não pode usar destinos de produção.")
        destinos.append(valor)
    registrar_commit_homologacao()
    return ambiente, *destinos
