# Aqua Gestão — Fluxo 100% Nuvem com GitHub Codespaces

## Abrir ambiente

1. Acesse o repositório no GitHub.
2. Clique em Code → Codespaces.
3. Abra o Codespace existente ou crie um novo.

## Secret usado no Codespaces

O ambiente usa este secret:

GCP_SERVICE_ACCOUNT_JSON

Ele deve conter o JSON inteiro da service account do Google.

## Preparar secrets locais do Streamlit

No terminal do Codespaces:

python scripts/codespaces_bootstrap.py

Esse comando gera:

.streamlit/secrets.toml

Esse arquivo nunca deve ser commitado.

## Validar ambiente

python scripts/healthcheck.py

Também valide sintaxe:

python -m py_compile app.py

## Rodar app no Codespaces

python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501

Depois abra a porta 8501 na aba PORTAS.

## Teste rápido

Modo operador:

PIN 5010

O app deve carregar os condomínios vinculados.

## Fluxo de edição

git pull origin main
python -m py_compile app.py
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501

Depois de testar:

git status
git add app.py
git commit -m "Descrição objetiva da alteração"
git push origin main

## Deploy

O Streamlit Cloud atualiza automaticamente após:

git push origin main

## Nunca commitar

.streamlit/secrets.toml
*.json
