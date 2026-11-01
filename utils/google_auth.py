"""Credenciais Google do app, separadas por ambiente.

Produção (comportamento inalterado): conta de serviço em
``[gcp_service_account]`` (ou arquivo local legado), escopos ``spreadsheets`` e
``drive``.

Homologação com Gmail pessoal: OAuth 2.0 de usuário em ``[google_oauth]``
(client_id, client_secret, refresh_token). Motivo: contas de serviço não têm
cota de armazenamento e não conseguem gravar fotos em "Meu Drive" de uma conta
Gmail (exigiria Drive compartilhado, só no Google Workspace).

Permissão mínima: o token de homologação só pode ter o escopo
``https://www.googleapis.com/auth/drive.file``, que dá acesso apenas aos
arquivos criados pelo próprio app (ou abertos com ele). Um escopo amplo
(``drive``/``spreadsheets``) daria acesso a TODAS as planilhas e arquivos da
conta, independentemente do ID configurado; por isso é recusado. A checagem de
IDs de produção em ``utils.ambiente`` continua valendo como segunda barreira.

Este módulo nunca registra nem devolve segredos em mensagens de erro.
"""
from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

ESCOPOS_PRODUCAO = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
)
ESCOPO_OAUTH_HOMOLOGACAO = "https://www.googleapis.com/auth/drive.file"
ESCOPOS_OAUTH_PERMITIDOS = frozenset({ESCOPO_OAUTH_HOMOLOGACAO})
TOKEN_URI = "https://oauth2.googleapis.com/token"
CAMPOS_OAUTH = ("client_id", "client_secret", "refresh_token")


class ConfiguracaoGoogleInvalida(ValueError):
    """Configuração recusada. A mensagem nunca contém valores de segredos."""


def _secao(secrets, nome):
    try:
        valor = secrets.get(nome)
    except FileNotFoundError:
        return None
    except Exception:
        return None
    return valor


def tipo_credencial(secrets, ambiente: str) -> str:
    """'oauth_usuario' ou 'conta_servico', validando a combinação com o ambiente."""
    oauth = _secao(secrets, "google_oauth")
    if oauth is not None:
        if ambiente != "homologacao":
            raise ConfiguracaoGoogleInvalida(
                "[google_oauth] só é aceito em homologação; produção usa a conta de serviço.")
        if not isinstance(oauth, Mapping):
            raise ConfiguracaoGoogleInvalida("[google_oauth] inválido.")
        faltando = [c for c in CAMPOS_OAUTH if not str(oauth.get(c) or "").strip()]
        if faltando:
            raise ConfiguracaoGoogleInvalida("[google_oauth] incompleto: falta " + ", ".join(faltando) + ".")
        escopos = oauth.get("scopes")
        if escopos is not None:
            escopos = [escopos] if isinstance(escopos, str) else list(escopos)
            if set(escopos) - ESCOPOS_OAUTH_PERMITIDOS:
                raise ConfiguracaoGoogleInvalida(
                    "[google_oauth] pede escopo amplo; em homologação só drive.file é permitido.")
        return "oauth_usuario"
    return "conta_servico"


def validar_escopos_concedidos(escopos) -> None:
    """Recusa token cujo consentimento tenha escopos além de drive.file."""
    if not escopos:
        return  # o servidor não informou; o pedido já foi limitado a drive.file
    concedidos = set(escopos.split()) if isinstance(escopos, str) else set(escopos)
    extras = concedidos - ESCOPOS_OAUTH_PERMITIDOS - {"openid"}
    if extras:
        raise ConfiguracaoGoogleInvalida(
            "O token OAuth de homologação tem escopos amplos (" + ", ".join(sorted(extras)) +
            "). Revogue o acesso e gere um novo token somente com drive.file.")


def credenciais_oauth(oauth: Mapping):
    """Credenciais de usuário (refresh token) limitadas a drive.file."""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    creds = Credentials(
        token=None,
        refresh_token=str(oauth["refresh_token"]).strip(),
        client_id=str(oauth["client_id"]).strip(),
        client_secret=str(oauth["client_secret"]).strip(),
        token_uri=str(oauth.get("token_uri") or TOKEN_URI),
        scopes=[ESCOPO_OAUTH_HOMOLOGACAO],
    )
    try:
        creds.refresh(Request())
    except Exception as erro:  # sem detalhes: a resposta pode ecoar dados do cliente
        raise ConfiguracaoGoogleInvalida(
            f"Falha ao renovar o token OAuth de homologação ({type(erro).__name__}). "
            "O token pode ter expirado ou sido revogado; gere outro com scripts/oauth_homologacao_setup.py."
        ) from None
    validar_escopos_concedidos(getattr(creds, "granted_scopes", None))
    return creds


def credenciais_google(secrets, ambiente: str, base_dir: Path | None = None):
    """Credenciais para Sheets e Drive conforme o ambiente. None se não houver."""
    tipo = tipo_credencial(secrets, ambiente)
    if tipo == "oauth_usuario":
        return credenciais_oauth(_secao(secrets, "google_oauth"))

    from google.oauth2.service_account import Credentials

    conta = _secao(secrets, "gcp_service_account")
    if conta is not None:
        return Credentials.from_service_account_info(dict(conta), scopes=list(ESCOPOS_PRODUCAO))
    if ambiente == "homologacao":
        return None  # homologação nunca usa o arquivo local de produção
    caminho = Path(base_dir or Path(__file__).resolve().parents[1]) / "aqua-gestao-rt-87316ebf5331.json"
    if not caminho.exists():
        return None
    return Credentials.from_service_account_file(str(caminho), scopes=list(ESCOPOS_PRODUCAO))
