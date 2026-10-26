#!/bin/sh
# Teste de navegador (Chromium/Playwright) do relatório RT do escritório.
# Roda uma CÓPIA do app numa pasta temporária, com planilha falsa em memória
# (tests/e2e/fakes), PIN fictício e proxy inválido: nenhum acesso a Google
# Sheets, Drive ou rede. Requer: python com streamlit e node com playwright.
# Uso: sh tests/e2e/rodar_rt_upload.sh [porta]
set -e
RAIZ=$(cd "$(dirname "$0")/../.." && pwd)
PORTA=${1:-8599}
PY=${PYTHON:-python3}
TMP=$(mktemp -d)
trap 'kill $PID 2>/dev/null; rm -rf "$TMP"' EXIT
(cd "$RAIZ" && git ls-files -z | grep -zv '^tests/' | xargs -0 tar -cf - | tar -xf - -C "$TMP")
mkdir -p "$TMP/.streamlit"
printf '[admin]\npin = "pin-ficticio-e2e"\n\n[gcp_service_account]\ntipo = "ficticio"\n' > "$TMP/.streamlit/secrets.toml"
cd "$TMP"
PYTHONPATH="$RAIZ/tests/e2e/fakes" CODESPACES=true \
  HTTPS_PROXY=http://127.0.0.1:9 HTTP_PROXY=http://127.0.0.1:9 https_proxy=http://127.0.0.1:9 http_proxy=http://127.0.0.1:9 \
  NO_PROXY=localhost,127.0.0.1 no_proxy=localhost,127.0.0.1 \
  "$PY" -m streamlit run app.py --server.port "$PORTA" --server.address 127.0.0.1 --server.headless true \
  --browser.gatherUsageStats false > "$TMP/servidor.log" 2>&1 &
PID=$!
i=0; until curl -s -o /dev/null "http://127.0.0.1:$PORTA/"; do i=$((i+1)); [ $i -gt 60 ] && { cat "$TMP/servidor.log"; exit 2; }; sleep 1; done
node "$RAIZ/tests/e2e/rt_upload_campos.js" "$PORTA" "pin-ficticio-e2e"
