# Protocolo de Desenvolvimento — Aqua Gestão App

## Regra principal
Nunca editar direto na main para melhorias ou novos formulários.

Fluxo obrigatório:
1. git checkout main
2. git pull
3. git checkout -b fix/nome-da-melhoria
4. editar app.py
5. python -m py_compile app.py
6. streamlit run app.py --server.port 8501
7. testar Aqua Gestão
8. testar Bem Star
9. git status
10. git diff -- app.py
11. git add app.py
12. git commit -m "mensagem objetiva"
13. git checkout main
14. git merge fix/nome-da-melhoria
15. git push

## Segurança
- Nunca commitar secrets.
- Nunca alterar lógica de dosagem sem autorização.
- Nunca alterar nomes das abas do Google Sheets sem autorização.
- Nunca usar st.session_state.clear().
- Nunca usar apenas CSS para esconder módulo exclusivo.
- Todo módulo exclusivo precisa de guard Python.

## Guards obrigatórios por empresa

```python
empresa_ativa = st.session_state.get("empresa_ativa", "aqua_gestao")

if empresa_ativa == "aqua_gestao":
    # módulos exclusivos Aqua Gestão
    ...

if empresa_ativa == "bem_star":
    # módulos exclusivos Bem Star
    ...
