# Protocolo de Desenvolvimento — Aqua Gestão App

## Regra principal
Nunca editar direto na main para melhorias ou novos formulários.

## Fluxo obrigatório
1. git checkout main
2. git pull
3. git checkout -b fix/nome-da-melhoria
4. editar app.py
5. python -m py_compile app.py
6. streamlit run app.py --server.port 8501
7. testar Aqua Gestão, Bem Star e modo operador
8. git status
9. git diff -- app.py
10. git add app.py
11. git commit -m "Mensagem objetiva"
12. git checkout main
13. git merge fix/nome-da-melhoria
14. git push

## Segurança
- Nunca commitar secrets.
- Nunca alterar lógica de dosagem sem autorização.
- Nunca alterar nomes das abas do Google Sheets sem autorização.
- Nunca usar st.session_state.clear().
- Nunca usar apenas CSS para esconder módulo exclusivo.
- Todo módulo exclusivo precisa de guard Python.

## Guards obrigatórios por empresa
empresa_ativa = st.session_state.get("empresa_ativa", "aqua_gestao")

if empresa_ativa == "aqua_gestao":
    # módulos exclusivos Aqua Gestão
    ...

if empresa_ativa == "bem_star":
    # módulos exclusivos Bem Star
    ...

## Antes de mexer em formulário
- git status
- python -m py_compile app.py
- localizar o bloco com grep antes de editar

## Depois de mexer
- python -m py_compile app.py
- streamlit run app.py --server.port 8501
- testar login Aqua Gestão
- testar login Bem Star
- testar modo operador
- testar se não houve vazamento de módulos entre empresas

## Recuperação rápida
Se algo quebrar antes do commit:
git restore app.py
python -m py_compile app.py