# Rascunhos RT no PostgreSQL de homologação

Branch `feat/rt-postgres-rascunhos`, a partir de `cb8c65f`. Esse commit contém a
versão do app validada no PR #14 (`9f7053c`) mais o contrato `rt_persist`, aprovado
nos 15 cenários.

## Onde o app salva e recupera rascunhos (antes desta etapa)

- `app.py` → `_relatorio_rt_salvar_rascunho`: coleta os campos `rel_*` e grava
  primeiro a cópia local íntegra (`utils/rascunho_sync.GestorRascunhos`). Depois
  envia à aba `_Rascunhos_RT`, com uma linha por sessão de edição.
- `_relatorio_rt_listar_rascunhos` / `_relatorio_rt_renderizar_painel_rascunho`:
  painel de versões. A restauração (`_relatorio_rt_aplicar_pendente`) começa uma
  chave de sessão nova.
- Identidade: `rel_origem_dados = "empresa|id:<cliente>|MM/AAAA"`. Fotos ficam
  dentro do rascunho (`rel_fotos_preservadas`, base64 + SHA-256 + identidade).

## Seleção do armazenamento

| Configuração | Armazenamento remoto |
|---|---|
| nenhuma (padrão, produção) | Google Sheets `_Rascunhos_RT` (sem mudança) |
| `CODESPACES=true` | somente disco local (sem mudança) |
| `ACQUA_RT_RASCUNHO_BACKEND=postgres` + `ACQUA_RT_PG_CONFIG=<toml>` | PostgreSQL `rt_persist` |

O PostgreSQL só é ativado quando o TOML (fora do repositório) declara
`ambiente = "homologacao"`. Qualquer outro valor, arquivo ausente ou conexão
desconhecida deixa o rascunho pendente no disco local, com aviso. Não há
fallback para a planilha. Com PostgreSQL ativo, a aba `_Rascunhos_RT` não é
lida nem gravada.

```toml
ambiente = "homologacao"
conexao = "executor_homologacao"            # database/rt/supabase_runner.py
manifesto = 'C:\Users\...\AppData\Local\AcquaGestao\rt-homologacao\authorization.jit.toml'

[empresas.aqua_gestao]
company_id = "10000000-0000-0000-0000-000000000001"   # EMPRESA A FICTICIA
principal = "rt_test_a"

[empresas.bem_star]
company_id = "20000000-0000-0000-0000-000000000002"   # EMPRESA B FICTICIA
principal = "rt_test_b"
```

A conexão reutiliza o executor validado: TLS verify-full, marcador do banco,
identidade SQL real e token JIT via DPAPI, sem senha e sem fallback. Por isso
o app com PostgreSQL roda **localmente no Windows** do responsável, não no
Streamlit Cloud.

## Modelo no contrato

- Documento = empresa + cliente + período. `company_id` vem da configuração.
  `condominium_id = uuid5(NAMESPACE, "empresa|id:<cliente>")`, determinístico.
  O período é o 1º dia do mês.
- O corpo é o rascunho do app mais `identity` (company/condominium/period). O banco
  recusa uma identidade divergente do envelope.
- Cada gravação envia a **revisão-base** que a sessão viu: 0 se nunca restaurou,
  a revisão restaurada, ou a última que ela mesma gravou. Se outra sessão já
  avançou o documento, o contrato registra **conflito** (`save_requests`) e nada é
  sobrescrito. A cópia local da sessão fica em `conflito`, e o painel oferece as
  duas versões para restauração explícita.
- Reenvio idempotente: `operation = uuid5(chave, hash do corpo, base)`, com a base
  guardada na cópia local.
- Isolamento: o escopo `access_scopes` é por LOGIN + empresa + condomínio +
  período. Sem escopo, a gravação e a leitura são negadas (42501). Revisões
  restauradas são conferidas contra identidade, empresa e origem.
- Revisões são imutáveis: o painel não oferece "Descartar" para versões do
  PostgreSQL. O descarte continua valendo para cópias locais.
- Limite do contrato: corpo ≤ 8 MiB, com as fotos incluídas.

## Provisionamento de escopos (administrador, dados fictícios)

`database/rt/provision_app_scopes.py`: somente clientes fictícios `C00N` e LOGINs
`rt_test_*`. Cria apenas o condomínio derivado e o escopo que faltam, nunca
empresa, papel ou privilégio, e nunca altera nem apaga. O padrão é simulação;
grava somente com `--aplicar`.

Observação: escopos novos de `rt_test_b` invalidam a pré-verificação de
`--run-selected`, que já foi concluída (15/15) e não deve ser repetida.

## Testes locais (sem rede, sem banco real)

- `tests/test_rt_postgres.py`: adaptador sobre o duplo do contrato
  (`tests/e2e/fakes/rt_persist_falso.py`) e seleção do backend extraída do
  `app.py` real.
- `tests/test_rt_postgres_apptest.py`: app real (AppTest) com a sequência salvar →
  reabrir em outra sessão → gerar relatório mensal (dados, foto, marca ACQUA
  GESTÃO) → duas sessões editando o mesmo rascunho (conflito, nada sobrescrito)
  → outro condomínio sem escopo (negado).
- `database/rt/test_provision_app_scopes.py`: cadastro de escopos com conexão simulada.

O duplo reproduz as regras do SQL, mas não substitui o PostgreSQL. A
confirmação no banco real é a etapa de homologação.

## Homologação real (2026-11-04, app local + PostgreSQL de homologação)

App da branch rodando localmente (`127.0.0.1:8502`) com os secrets de homologação
(planilha fictícia via OAuth `drive.file`) e o PostgreSQL via executor (JIT/DPAPI).
Escopo fictício criado só para Alfa (C001) 09/2026, usando `rt_test_a`. Dados
fictícios; a V00005 (Beta, C002) não foi tocada.

| Verificação | Resultado no PostgreSQL real |
|---|---|
| Visita de 29/09/2026 importada para Alfa 09/2026 | 2 linhas (pH 7,4/CRL 2,5 e pH 7,5/CRL 2,2) |
| Salvar rascunho (sessão 1) | revisões 1–4, base 0→3, autor `rt_test_a`, identidade conferida |
| Foto no rascunho | revisões 3–5 com a foto fictícia; restaurada na sessão 2 |
| Reabrir em outra sessão | sessão 2 listou e restaurou a revisão 4 (dados, linhas e foto) |
| Relatório mensal (sessão 2) | DOCX e PDF com ACQUA GESTÃO, Alfa, 09/2026, 29/09/2026, diagnóstico e foto 240×180; sem Bem Star |
| Sobrescrita: sessão 2 sem restaurar | autosave base 0 com o banco na revisão 4 → conflito registrado, nada alterado |
| Sobrescrita: duas sessões na revisão 4 | sessão 2 gravou a revisão 5; sessão 1 → conflito (base 4, revisão observada 5), revisão 5 intacta, aviso "Nada foi sobrescrito" |
| Isolamento por período | Alfa 11/2026 negado (42501); cópias locais preservadas |
| Isolamento por condomínio | Beta 09/2026 (V00005) negado (42501) para `rt_test_a` |
| Isolamento por empresa | `rt_test_b` (Bem Star) negado (42501) ao ler Alfa |
| Limpeza | LOGINs NOLOGIN, senha nula, zero sessões; regra excluída; token revogado (401) |

Achados corrigidos durante a homologação:
1. Uma pendência de outro período ou condomínio (sem escopo) bloqueava o salvamento
   do período atual na mesma chave. Agora ela continua preservada, e a edição
   atual segue em chave nova. Há teste de regressão em
   `tests/test_rt_postgres_apptest.py`.
2. Concordância das mensagens ("à planilha" / "ao PostgreSQL…"). Os textos da
   planilha continuam idênticos.

3. Seleção do armazenamento compatível com os testes de rascunho da planilha:
   `_rascunho_usa_sheets()` continua sendo o seletor principal, e o PostgreSQL é
   detectado por `_rascunho_pg_ativo()`. Os testes que extraem funções do `app.py`
   receberam só os nomes novos na lista `FUNCOES_RT`; nenhuma asserção mudou.

### Legenda da foto (2026-11-06)

Na rodada de 2026-11-04 a legenda não foi gravada porque o campo fica dentro do
expansor "📷 Legendas e piscinas das fotografias", que abre fechado. O texto foi
digitado sem abrir o expansor e nunca entrou no campo. O código não tinha defeito.
Na nova rodada, com o expansor aberto e o valor conferido antes de salvar, o mesmo
texto `LEGENDA_SETEMBRO_PG_123` foi confirmado em três lugares:

| Lugar | Resultado |
|---|---|
| PostgreSQL de homologação | revisões 6 (autosave) e 7 (manual), base 5→6, autor `rt_test_a`, foto com `legenda` |
| Rascunho reaberto (sessão 2) | revisão 7 restaurada; campo "Legenda — foto 1" com o texto |
| PDF mensal de Alfa 09/2026 | página 6: "Foto 1 — LEGENDA_SETEMBRO_PG_123", foto 240×180; DOCX também |

As revisões 1 a 5 continuam intactas. Mais um conflito foi registrado (base 0,
revisão observada 5): o autosave ao abrir Alfa antes de restaurar, sem
sobrescrita. Limpeza verificada: LOGINs NOLOGIN, regra excluída, token revogado.

## Suíte local integral (2026-11-05)

`tests/rodar_isolado.py`, venv espelhando o Python 3.14 do sistema, mais PyMuPDF e
`PYTHONUTF8=1`: 423 testes. 417 aprovados de primeira; 6 testes de app excederam
o limite de 300 s por falta de memória (execução de cerca de 10,5 h) e passaram
ao rodar de novo, um a um. Resultado: 423 de 423.

## Homologação no Streamlit Cloud (2026-11-07)

App de homologação criado só para este PR, sem alterar o pr14, o `homologacao-rt`
nem a produção:

- URL: https://aqua-gestao-homologacao-pr16.streamlit.app
- Branch `feat/rt-postgres-rascunhos`, commit implantado
  `4b32fcb4cd01d49e046ae4677f35ecf0ee0bb738` (o log do app mostra
  `AQUA_HOMOLOGACAO_COMMIT`), Python 3.12.
- Secrets iguais aos do pr14 (planilha fictícia), sem chaves `ACQUA_RT_*`: os
  rascunhos ficam na planilha (backend padrão).

| Verificação | Resultado |
|---|---|
| Fluxo do operador (PIN de operador) | Alfa, visita de 07/11/2026 (sábado) com foto, pH 7,4 / CRL 2,0 e 7,5 / 1,8; gravada como V00006 |
| Fluxo do RT, sessão 1 (PIN administrativo) | Alfa 10/2026; foto com legenda `LEGENDA_PR16_CLOUD_OUT` e diagnóstico `DIAGNOSTICO_PR16_CLOUD`; "Rascunho completo preservado na planilha" |
| Reabertura em outra sessão (outra aba, novo login) | rascunho das 18:44:36 restaurado com diagnóstico, linha de análise, legenda e foto |
| PDF mensal (sessão 2) | 6 páginas; ACQUA GESTÃO em todas, sem Bem Star; diagnóstico na p. 5; "Foto 1 — LEGENDA_PR16_CLOUD_OUT" e a foto fictícia 640×480 na p. 6 |

V00005 e V00006 foram preservadas. Ficaram na planilha de homologação a V00006 e
os rascunhos de Alfa das 18:43 às 18:47.

### Correção do registro: visitas de operador no mensal RT

O relato da homologação dizia que uma visita do operador numa terça-feira
entraria no mensal RT. Isso descrevia uma falha, não a regra. A regra é: o mensal
RT usa só visitas técnicas RT, e visitas de operador ficam fora em qualquer dia.

A conferência do código confirmou a falha. No relatório mensal do escritório, o
filtro aplicava só "terça-feira + uma visita por semana". A área RT do PIN já
usava `exclude_operator_records`, mas o escritório não. Na própria homologação, a
linha de análise de 06/10/2026 do PDF acima veio da V00002, gravada pelo
"OPERADOR FICTÍCIO" e sem marca de RT. Com a regra correta, ela não entra. As
verificações de rascunho, reabertura, foto, legenda, diagnóstico e marca
continuam válidas. A linha de análise daquele PDF não vale como evidência de
visita técnica.

Correção (`app.py`, importação do relatório mensal):

- `exclude_operator_records` é aplicado antes do filtro de terça-feira e do limite
  semanal, para que uma visita de operador não ocupe a vaga da semana.
- Entram só visitas com marca de RT: PIN RT (`visita_rt_semanal` / `tipo_visita =
  "RT semanal"`) ou RT manual.
- Ficam fora as visitas de operador e as linhas antigas sem Payload JSON nem marca
  (decisão do responsável). Um aviso informa quantas ficaram fora, e nenhum
  registro é apagado.

Testes:

- `tests/test_rt_exclui_operador.py` (novo): visitas de operador em terças (06/10
  e 27/10, esta antes de uma visita RT na mesma semana) e uma linha antiga sem
  marca ficam fora; entram só as visitas RT de 13/10 e 27/10. Sem a correção, o
  teste falha porque a visita de operador de 06/10 é importada e a do dia 27/10
  toma a vaga da visita RT.
- As fixtures que usavam visitas sem marca como fonte do mensal RT passaram a ter
  marca de RT (coluna Operador "Responsável Técnico Fictício" ou payload
  `RT semanal`). Nenhuma asserção mudou.
- Testes locais relacionados (2026-11-07), todos aprovados, rodados módulo a módulo
  por falta de memória numa rodada única: `test_rt_exclui_operador`,
  `test_relatorio_rt_campos`, `test_rt_relatorio_apptest`, `test_rt_postgres`,
  `test_rt_postgres_apptest`, `test_rascunho_rt`, `test_rt_protecao_integrada`,
  `test_homologacao_pr14*` (5 módulos), `test_rt_ultimos_*`, `test_rt_cloud_fluxo`,
  `test_fotos_isolamento`, `test_operador_formulario`,
  `test_rascunho_legado_restauracao` e `test_ultimos_sheets`. São testes locais e
  não substituem o reteste no app de homologação.

Consequência na homologação: as visitas fictícias V00001–V00004 não têm marca de
RT e não entram mais no mensal RT. Para repetir o teste do mensal no app, uma
visita técnica precisa ser registrada pelo PIN RT numa terça-feira.
