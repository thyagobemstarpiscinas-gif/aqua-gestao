# v6: healthcheck do ambiente Codespaces e Google Sheets — BUG-CLOUD
"""
Valida o ambiente do Aqua Gestão no Codespaces.

Uso:
    python scripts/healthcheck.py

Checa:
- app.py compila.
- .streamlit/secrets.toml existe.
- credencial Google é válida.
- conexão com Google Sheets abre.
- abas essenciais existem.
- abas auxiliares ausentes viram aviso, não bloqueio.
"""

from __future__ import annotations

import py_compile
import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_PATH = ROOT / "app.py"
SECRETS_PATH = ROOT / ".streamlit" / "secrets.toml"

SHEETS_ID = "1uvZ6qfYCYFl_feGGgvIIXMQlUWvx0MTzTuC8TwfPBlM"

ABAS_OBRIGATORIAS = {
    "👥 Clientes",
    "👷 Operadores",
}

ABAS_AUXILIARES = {
    "_Rascunhos",
    "_RascunhosRT",
}


def ok(msg: str) -> None:
    print(f"OK: {msg}")


def aviso(msg: str) -> None:
    print(f"AVISO: {msg}")


def erro(msg: str) -> None:
    print(f"ERRO: {msg}")


def validar_compile() -> bool:
    try:
        py_compile.compile(str(APP_PATH), doraise=True)
        ok("app.py compila sem erro.")
        return True
    except Exception as exc:
        erro(f"app.py não compila: {exc}")
        return False


def carregar_credencial() -> dict | None:
    if not SECRETS_PATH.exists():
        erro(".streamlit/secrets.toml não encontrado. Rode scripts/codespaces_bootstrap.py.")
        return None

    try:
        data = tomllib.loads(SECRETS_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        erro(f"secrets.toml inválido: {exc}")
        return None

    gcp = data.get("gcp_service_account")
    if not isinstance(gcp, dict):
        erro("seção [gcp_service_account] não encontrada.")
        return None

    email = str(gcp.get("client_email", "")).strip()
    private_key = str(gcp.get("private_key", ""))

    if not email:
        erro("client_email ausente.")
        return None

    if "BEGIN PRIVATE KEY" not in private_key or "END PRIVATE KEY" not in private_key:
        erro("private_key ausente ou inválida.")
        return None

    ok(f"credencial carregada para {email}.")
    return gcp


def validar_sheets(gcp: dict) -> bool:
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception as exc:
        erro(f"dependências Google ausentes: {exc}")
        return False

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        creds = Credentials.from_service_account_info(gcp, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SHEETS_ID)
        ok(f"planilha aberta: {sh.title}")

        abas = {ws.title for ws in sh.worksheets()}

        obrigatorias_faltantes = sorted(ABAS_OBRIGATORIAS - abas)
        if obrigatorias_faltantes:
            erro("abas obrigatórias ausentes: " + ", ".join(obrigatorias_faltantes))
            return False

        if "🔬 Visitas" not in abas and "📋 Lançamentos" not in abas:
            erro("aba de visitas ausente: esperado 🔬 Visitas ou 📋 Lançamentos")
            return False

        auxiliares_faltantes = sorted(ABAS_AUXILIARES - abas)
        if auxiliares_faltantes:
            aviso("abas auxiliares ausentes: " + ", ".join(auxiliares_faltantes))

        ok("abas obrigatórias encontradas.")

        operadores = sh.worksheet("👷 Operadores").get_all_values()
        ok(f"aba 👷 Operadores lida com {len(operadores)} linha(s).")

        clientes = sh.worksheet("👥 Clientes").get_all_values()
        ok(f"aba 👥 Clientes lida com {len(clientes)} linha(s).")

        return True

    except Exception as exc:
        erro(f"falha na conexão Google Sheets: {exc}")
        return False


def main() -> int:
    checks = []

    checks.append(validar_compile())

    gcp = carregar_credencial()
    checks.append(gcp is not None)

    if gcp:
        checks.append(validar_sheets(gcp))

    if all(checks):
        ok("healthcheck concluído com sucesso.")
        return 0

    erro("healthcheck encontrou falhas.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
