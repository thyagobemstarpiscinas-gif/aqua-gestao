"""Configuração única do OAuth de HOMOLOGAÇÃO (executar no seu computador).

O que faz, com o escopo mínimo drive.file:
1. Abre o navegador para você consentir com a conta Gmail de TESTE.
2. Cria, com o próprio app, uma planilha e uma pasta exclusivas para dados
   fictícios. Com drive.file, o app só enxerga o que ele mesmo criou: nenhuma
   outra planilha ou arquivo da conta fica acessível.
3. Mostra no terminal o trecho para colar em Streamlit Cloud → app de
   homologação → Settings → Secrets.

Nada é gravado no repositório. O refresh token aparece somente no terminal:
copie direto para o painel de Secrets e limpe o histórico/rolagem do terminal.

Uso (fora do repositório, em ambiente local):
    pip install google-auth-oauthlib gspread google-api-python-client
    python scripts/oauth_homologacao_setup.py --client-secret CAMINHO/client_secret.json

O arquivo client_secret.json é do tipo "App para computador" (Desktop) e deve
ficar FORA da pasta do projeto. Não envie esse arquivo nem o token por chat.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ESCOPO = "https://www.googleapis.com/auth/drive.file"
TITULO_PLANILHA = "ACQUA HOMOLOGACAO - DADOS FICTICIOS"
TITULO_PASTA = "ACQUA HOMOLOGACAO - FOTOS FICTICIAS"
ABAS = ("👥 Clientes",)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--client-secret", required=True, help="client_secret.json (Desktop) fora do repositório")
    args = parser.parse_args()

    caminho = Path(args.client_secret).expanduser().resolve()
    raiz = Path(__file__).resolve().parents[1]
    if raiz in caminho.parents:
        print("ERRO: mantenha o client_secret.json fora da pasta do projeto.")
        return 1

    from google_auth_oauthlib.flow import InstalledAppFlow
    import gspread
    from googleapiclient.discovery import build

    flow = InstalledAppFlow.from_client_secrets_file(str(caminho), scopes=[ESCOPO])
    # access_type=offline + prompt=consent garantem a emissão do refresh token.
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent",
                                  include_granted_scopes="false")
    concedidos = set(creds.granted_scopes or [ESCOPO])
    if concedidos - {ESCOPO, "openid"}:
        print("ERRO: o consentimento incluiu escopos além de drive.file:", ", ".join(sorted(concedidos)))
        print("Revogue em https://myaccount.google.com/permissions e repita.")
        return 1
    if not creds.refresh_token:
        print("ERRO: o Google não devolveu refresh token. Revogue o acesso do app e repita.")
        return 1

    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    pasta = drive.files().create(
        body={"name": TITULO_PASTA, "mimeType": "application/vnd.google-apps.folder"}, fields="id").execute()

    gc = gspread.authorize(creds)
    planilha = gc.create(TITULO_PLANILHA)
    primeira = planilha.sheet1
    primeira.update_title(ABAS[0])
    for aba in ABAS[1:]:
        planilha.add_worksheet(title=aba, rows=100, cols=12)

    print("\nRecursos de TESTE criados pelo app (somente eles ficam acessíveis com drive.file):")
    print(f"  Planilha: {TITULO_PLANILHA}")
    print(f"  Pasta:    {TITULO_PASTA}")
    print("\nCole no Secrets do app de HOMOLOGAÇÃO (nunca no repositório, nunca no app de produção):\n")
    print("[aqua_app]")
    print('ambiente = "homologacao"')
    print(f'sheets_id = "{planilha.id}"')
    print(f'drive_fotos_folder_id = "{pasta["id"]}"')
    print("\n[google_oauth]")
    print(f'client_id = "{creds.client_id}"')
    print(f'client_secret = "{creds.client_secret}"')
    print(f'refresh_token = "{creds.refresh_token}"')
    print(f'scopes = ["{ESCOPO}"]')
    print("\nRemova [gcp_service_account] desse app de homologação, se existir.")
    print("Depois de copiar, limpe o terminal (o token dá acesso aos arquivos acima).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
