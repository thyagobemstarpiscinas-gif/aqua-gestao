# v6: bootstrap seguro do Codespaces — BUG-CLOUD
"""
Gera .streamlit/secrets.toml a partir do secret GCP_SERVICE_ACCOUNT_JSON.

Uso:
    python scripts/codespaces_bootstrap.py

Regras:
- Não imprime chave privada.
- Não commita secrets.
- Não cria dependências novas.
"""

from __future__ import annotations

import json
import os
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
STREAMLIT_DIR = ROOT / ".streamlit"
SECRETS_PATH = STREAMLIT_DIR / "secrets.toml"


def _toml_escape(value: object) -> str:
    texto = str(value)
    texto = texto.replace("\\", "\\\\")
    texto = texto.replace("\n", "\\n")
    texto = texto.replace("\r", "\\r")
    texto = texto.replace('"', '\\"')
    return texto


def main() -> int:
    raw = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        print("ERRO: secret GCP_SERVICE_ACCOUNT_JSON não encontrado no Codespaces.")
        print("Crie em GitHub → Settings → Secrets and variables → Codespaces.")
        return 1

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        print(f"ERRO: GCP_SERVICE_ACCOUNT_JSON não é JSON válido: {exc}")
        return 1

    email = str(data.get("client_email", "")).strip()
    private_key = str(data.get("private_key", ""))

    if not email:
        print("ERRO: client_email ausente no JSON.")
        return 1

    if "BEGIN PRIVATE KEY" not in private_key or "END PRIVATE KEY" not in private_key:
        print("ERRO: private_key ausente ou inválida no JSON.")
        return 1

    STREAMLIT_DIR.mkdir(exist_ok=True)

    lines = ["[gcp_service_account]"]
    for key, value in data.items():
        lines.append(f'{key} = "{_toml_escape(value)}"')

    SECRETS_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("OK: .streamlit/secrets.toml gerado com segurança.")
    print(f"OK: service account: {email}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
