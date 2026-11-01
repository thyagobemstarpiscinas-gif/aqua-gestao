# ACQUA GESTÃO — Homologação com Gmail pessoal (OAuth 2.0, escopo `drive.file`)

Situação: conta Gmail pessoal, sem Google Workspace e sem Drive compartilhado.
Objetivo: homologar com uma planilha e uma pasta exclusivas para dados
fictícios, sem tocar a produção. **Nada aqui foi executado contra o Google**:
nenhuma credencial, consentimento ou recurso foi criado. Os passos abaixo só
devem ser feitos depois da sua aprovação.

## 1. Auditoria da autenticação atual

| Item | Produção (inalterada) | Antes, em homologação |
|---|---|---|
| Identidade | Conta de serviço `[gcp_service_account]` | A mesma lógica; sem Secrets caía no arquivo local `aqua-gestao-rt-87316ebf5331.json` (credencial de produção) |
| Escopos | `spreadsheets` + `drive` (amplos) | Iguais |
| Isolamento | Só o que foi compartilhado com a conta de serviço | Dependia de uma conta de serviço de teste separada e da checagem de IDs (`utils/ambiente.py`) |
| Fotos | `files.create` em pasta de "Meu Drive" do usuário | Igual |

Achados:
- **Fotos com conta de serviço em Gmail pessoal**: contas de serviço não têm
  cota de armazenamento; o Google responde `403 storageQuotaExceeded` ao criar
  arquivos em "Meu Drive" de um usuário (contas criadas após 15/04/2025 não
  recebem cota). A saída indicada pelo Google é Drive compartilhado (exige
  Workspace) ou OAuth em nome do usuário. Por isso a homologação de fotos com
  Gmail pessoal precisa de OAuth. *Recomendação separada:* verificar, sem
  alterar a produção, se os uploads de produção estão de fato chegando ao Drive.
- Homologação sem `[gcp_service_account]` podia usar o arquivo local de produção
  (corrigido: agora retorna "sem credencial").
- Consulta de pastas do Drive quebrava com apóstrofo no nome do condomínio
  (corrigido).
- Rascunho salvo explicitamente na planilha nunca era gravado numa planilha
  nova (aba `_Rascunhos` não era criada) e a leitura podia ficar presa a um
  cache de "aba ausente" (corrigido; afeta também a produção quando a aba não
  existe).

## 2. Por que `drive.file` e não os escopos amplos

OAuth **não** restringe acesso por ID. O ID configurado só diz ao app qual
arquivo abrir; quem limita o alcance é o escopo:

| Escopo | Alcance real do token | Classificação Google |
|---|---|---|
| `spreadsheets` | TODAS as planilhas da conta | Sensível |
| `drive` | TODOS os arquivos do Drive da conta | Restrito |
| `drive.file` | Somente arquivos **criados pelo app** (ou abertos com ele) | Não sensível, recomendado |

Com `drive.file`, a planilha e a pasta de teste são **criadas pelo próprio app**
(script de configuração). Assim o token não enxerga nenhuma outra planilha,
foto ou documento da conta — nem a planilha de produção, mesmo que alguém
configure o ID dela. A API do Sheets aceita `drive.file` para planilhas criadas
pelo app.

Barreiras implementadas (`utils/google_auth.py`):
1. `[google_oauth]` só é aceito com `ambiente = "homologacao"`; em produção o app
   para na inicialização.
2. O token é pedido somente com `drive.file`; se o consentimento tiver qualquer
   escopo extra (`drive`, `spreadsheets`…), o app recusa o token.
3. IDs de produção continuam bloqueados em homologação (`utils/ambiente.py`).
4. Homologação nunca usa o arquivo local de credencial de produção.
5. Mensagens de erro nunca incluem client secret ou tokens.

## 3. Fluxo proposto (para aprovação)

```
Você (uma vez, no seu computador)                     Streamlit Cloud (app de homologação)
───────────────────────────────────                   ────────────────────────────────────
1. Google Cloud: projeto NOVO só de homologação
2. Ativar APIs: Google Drive API e Google Sheets API
3. Tela de consentimento: Externo; escopo drive.file
   Usuário de teste: a conta Gmail de teste
4. Criar cliente OAuth "App para computador"
   (baixar client_secret.json para FORA do projeto)
5. python scripts/oauth_homologacao_setup.py \
     --client-secret ~/Downloads/client_secret.json
   → navegador: login com a conta de TESTE e consentir
   → o script cria a planilha e a pasta de teste
   → imprime o bloco de Secrets
6. Colar o bloco em Settings → Secrets  ───────────►  app lê [google_oauth] e [aqua_app]
                                                     renova o access token (1 h) com o
                                                     refresh token; acessa só a planilha
                                                     e a pasta criadas no passo 5
7. python scripts/healthcheck.py (opcional, local)
```

Secrets do app de homologação (valores impressos pelo script; não registrar em
lugar nenhum além do painel):

```toml
[aqua_app]
ambiente = "homologacao"
sheets_id = "<impresso pelo script>"
drive_fotos_folder_id = "<impresso pelo script>"

[google_oauth]
client_id = "<impresso pelo script>"
client_secret = "<impresso pelo script>"
refresh_token = "<impresso pelo script>"
scopes = ["https://www.googleapis.com/auth/drive.file"]
```

Também configurar no app de homologação: PINs de teste em `[rt]`, `[admin]` e
`[operador]` (nunca os de produção) e acesso restrito ao app (privado).
Não definir `CODESPACES=true`. O app de produção não recebe nada disso.

## 4. Riscos e mitigação

| Risco | Mitigação |
|---|---|
| Refresh token vazado dá acesso aos arquivos que o app criou | Escopo `drive.file` (nada além disso); token só no painel de Secrets; app de homologação privado; revogar em myaccount.google.com/permissions |
| Uso da conta Gmail pessoal | Recomendado criar uma conta Gmail **só de teste**; com a pessoal, o token ainda não enxerga outros arquivos, mas a identidade é a sua |
| Tela "app não verificado" no consentimento | Esperado para app próprio; confirmar que o nome do app e o escopo exibido são somente "arquivos usados com este app" |
| Status "Testing" (Externo): refresh token expira em **7 dias** | Rodar o script de novo por semana, ou publicar o app ("In production"). `drive.file` é não sensível; confirmar no console se a publicação exige verificação. Relatos indicam ser preciso gerar novo token após publicar |
| Limite de 100 refresh tokens por cliente/conta | Novos tokens invalidam os mais antigos sem aviso; usar só o último |
| Fotos ocupam a cota da conta Gmail (15 GB gratuitos) | Usar imagens fictícias pequenas; apagar a pasta de teste ao final |
| Arquivos criados pelo app ficam acessíveis a qualquer código com o token | Cliente OAuth e projeto Google Cloud exclusivos de homologação; nunca reutilizar em produção |
| Mesmo cliente OAuth usado por outra pessoa | Só a conta de teste como usuário de teste |

Itens a confirmar na documentação oficial do Google (o acesso a
developers.google.com estava bloqueado nesta sessão; as informações vieram de
buscas): expiração de 7 dias em "Testing", exigências para publicar com
`drive.file`, e a regra de cota das contas de serviço.

## 5. Encerramento da homologação

1. Revogar o acesso do app em myaccount.google.com/permissions (conta de teste).
2. Excluir o cliente OAuth / projeto Google Cloud de homologação.
3. Apagar a planilha e a pasta de teste.
4. Remover `[google_oauth]` dos Secrets do app de homologação.

## 6. Testes locais (sem rede)

`tests/test_google_oauth_homologacao.py` — 17 testes com Google falso que aplica
a regra do `drive.file`:
- segurança: produção inalterada (conta de serviço e mesmos escopos); OAuth
  recusado em produção; escopos amplos recusados na configuração e no token;
  erros sem segredos; homologação não usa o arquivo local de produção; IDs de
  produção bloqueados; repositório sem tokens/chaves;
- gravação e leitura na planilha de teste; planilha de produção inacessível;
- fotografias: estrutura condomínio/mês, nome com apóstrofo, download, recusa
  fora da pasta de teste;
- rascunhos: gravação na planilha de teste, recuperação sem cópia local,
  regravação sem duplicar linha;
- script de configuração: pede só `drive.file`, recusa escopo amplo e
  `client_secret.json` dentro do projeto.
