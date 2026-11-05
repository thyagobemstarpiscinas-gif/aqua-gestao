# Executor Supabase — instalacao controlada, sem integracao ao aplicativo

Implementacao derivada de 6f731373655c91c9283a8742ceecf56fed4b3e70.
A instalacao inicial abaixo e um roteiro separado: nao reaplicar migracoes
nem provisionar LOGINs ao executar as inspecoes de Session Pooler.
A migracao original e o executor local original permanecem inalterados.
Python 3.11+; dependencias em requirements-test.txt.

## Autorizacao e destino

A execucao remota exige confirmacao do proprietario. Nao enviar senhas, tokens,
pgpass ou strings de conexao pelo chat, Git ou logs. Usar computador controlado.
Conferir no painel: acqua-gestao-homologacao, Free, sa-east-1, Healthy,
Data API desativada. Obter Project Ref e endpoint direto no proprio painel.
O executor suporta endpoint direto db.<ref>.supabase.co:5432 ou, mediante
manifesto explicito connection_mode="session_pooler", Session Pooler compartilhado
em sa-east-1, porta 5432, banco postgres. Transaction Pooler/porta 6543 sao recusados.
Nao relaxar validacao de hostname ou TLS. Configuracoes anteriores sem
connection_mode continuam exigindo conexao direta.

Configurar copia de supabase.example.toml FORA do repositorio, com modo 0600.
Esse arquivo autoriza exatamente o Project Ref, hostname, porta e banco escolhidos.
O nome amigavel nao e prova da identidade remota: conferir o ref no painel.
Gerar marcador aleatorio de 32 hexadecimais para vincular configuracao ao banco.
Ele e um intertravamento operacional, nao substitui credenciais nem TLS.

Credenciais somente em passfile libpq privado modo 0600, sem curingas:
hostname:5432:postgres:postgres:senha-administrativa
hostname:5432:postgres:rt_test_a:senha-exclusiva
hostname:5432:postgres:rt_test_b:senha-exclusiva
hostname:5432:postgres:rt_test_none:senha-exclusiva
Preencher em editor seguro; respeitar escape libpq de dois-pontos e barra invertida.
Nao usar credenciais de producao. Nao colocar senhas em argumentos de terminal.
O manifesto rejeita password, DSN, service, sslmode e outras opcoes desconhecidas.
O executor sempre usa verify-full e CA explicita, e remove defaults PG* do
ambiente de seu proprio processo para nao herdar PGOPTIONS/PGSERVICE ocultos.
Proteger os arquivos com ACL equivalente em Windows; o ensaio validado usa Linux.

## Session Pooler: inspecao previa, sem provisionamento

Criar um manifesto separado em pasta privada fora do repo, usando o modelo
abaixo. Substituir os placeholders localmente, sem enviar valores pelo chat/Git:

```toml
environment = "homologacao"
connection_mode = "session_pooler"
project_name = "acqua-gestao-homologacao"
project_ref = "PREENCHER_REF_REAL"
host = "aws-PREENCHER_INDICE-sa-east-1.pooler.supabase.com"
port = 5432
database = "postgres"
marker = "PREENCHER_32_HEXADECIMAIS_EXISTENTES"
admin_user = "postgres"
ca_file = "/caminho/absoluto/ca-supabase.crt"
passfile = "/caminho/absoluto/pgpass-session-pooler"
```

Esta secao atualiza o suporte de transporte do executor. O roteiro separado
PROVISION_TEST_LOGINS.md descreve provisionamento por conexao direta; a
inspecao de Session Pooler aqui documentada nao autoriza sua execucao pelo pooler.

Conferir no painel Connect > Session pooler o host EXATO
aws-<indice>-sa-east-1.pooler.supabase.com, porta 5432 e usuario
postgres.<project_ref>. O indice nao pode ser deduzido pela regiao.
O manifesto fixa esse host e o ref autorizado. admin_user permanece postgres.
O executor deriva <LOGIN>.<project_ref> somente para autenticacao/roteamento;
session_user e current_user devem retornar o LOGIN PostgreSQL sem sufixo.
O comentario do banco deve coincidir com o marcador ja guardado no manifesto.
O host compartilhado, sozinho, nao identifica o projeto.

Usar passfile separado com host/banco/usuario exatos, sem curingas. O campo usuario
passa a ser postgres.<project_ref> ou rt_test_a.<project_ref>, por exemplo. Nao
alterar senhas no Supabase para mudar de transporte. Preservar arquivos diretos.
CA explicita confiavel e TLS verify-full continuam obrigatorios. Tanto o TLS do
cliente quanto o TLS observado no backend por pg_stat_ssl precisam ser verdadeiros;
uma falha interrompe o executor, sem fallback para require ou verify-ca.

Antes de criar LOGINs, executar SOMENTE:

    python database/rt/supabase_runner.py --config /privado/authorization.session-pooler.toml --inspect-session

Esse modo exige apenas o administrador existente. Abre duas conexoes reais,
verifica identidade/marcador/TLS e somente depois aplica por SQL de sessao
statement_timeout=20s, lock_timeout=10s e default_transaction_read_only=on.
Nao envia esses ajustes em options de startup: na homologacao, o pooler aceitou
essas opcoes sem aplica-las ao PostgreSQL. A simples presenca no cliente nao
constitui validacao. Valores efetivos divergentes interrompem a conexao.
Cada transacao de inspecao abre com BEGIN READ ONLY, confirma os dois timeouts,
transaction_read_only=on e o padrao de leitura antes das consultas. Depois de
COMMIT/ROLLBACK, a sondagem repete a abertura explicita e as verificacoes.
Usa SELECTs, COMMIT, ROLLBACK e SAVEPOINT/ROLLBACK TO/RELEASE; sem SET ROLE,
DDL, DML, tabelas temporarias, advisory locks ou escrita de fixtures. Os SETs
acima afetam somente a sessao; nao alteram papeis, banco ou configuracao global.
Conexoes de outros modos tambem aplicam e verificam os timeouts por SQL, mas o
executor so habilita leitura obrigatoria nos modos de inspecao. Confere
backends distintos e estaveis entre transacoes, statement_timeout=20s,
lock_timeout=10s e isolamento read committed. Nenhuma senha ou marcador completo
e impresso. Esperado: SESSION_INSPECTION_OK e resumo sem segredos.

Em falha de TLS, SESSION_INSPECTION_STOP pode mostrar somente indicadores
booleanos de identidade/marcador, TLS do cliente e TLS do backend, sem valores
privados. pg_stat_ssl descreve a conexao que chega ao PostgreSQL; atraves do
pooler, ela corresponde ao trecho pooler -> PostgreSQL. Ela nao substitui a
validacao verify-full do cliente -> pooler. Se o backend nao confirmar TLS,
parar. Nao remover a verificacao para liberar provisionamento.

A documentacao oficial informa que Enforce SSL on incoming connections governa
os dois trechos do Shared Pooler e que sua alteracao reinicia o banco. Verificar
o estado atual no painel e revisar a correcao separadamente com o proprietario;
este executor nao altera esse ajuste nem reinicia o banco.

Essa sondagem comprova somente as propriedades observadas nessas sessoes.
Nao comprova rollback de dados, concorrencia do contrato, custom LOGINs ainda
ausentes, backup ou todos os cenarios transacionais. A suite de 15 testes continua
pendente ate confirmacao do proprietario. Nao provisionar nem executar fixtures
apenas porque o TCP funciona ou a sondagem passou. --inspect completo ainda
exige os tres LOGINs existentes e verifica suas ACLs e identidades reais.
Em Session Pooler, --inspect completo repete tambem a sondagem de sessoes.

Referencias oficiais:
- https://supabase.com/docs/guides/database/connecting-to-postgres
- https://supabase.com/docs/guides/database/psql
- https://supabase.com/docs/guides/database/postgres/timeouts
- https://supabase.com/docs/guides/platform/ssl-enforcement
- https://www.postgresql.org/docs/17/monitoring-stats.html#MONITORING-PG-STAT-SSL-VIEW

## Preparacao administrativa — somente apos confirmacao

1. Conferir commit, arquivo 001_contract.sql e ausencia de rt_persist.
   Rodar inspect_supabase.sql em modo administrativo: somente leitura.
   Conferir capacidade de criar roles/esquema e disponibilidade de sha256.
2. Aplicar 001_contract.sql completo via SQL Editor ou psql com ON_ERROR_STOP.
   Nao reaplicar quando houver esquema/role existente; nao apagar para destravar.
   O executor NAO le nem aplica migracoes.
3. Inspecionar proprietarios/ACLs com inspect_supabase.sql. Os papeis anon,
   authenticated e service_role nao devem ter USAGE/CREATE no esquema, acesso
   a tabelas/colunas nem EXECUTE nas funcoes. Conferir PUBLIC e defaults globais.
   Se houver concessoes extras, parar e revisar uma correcao de ACL separada;
   o executor nao revoga nem concede privilegios automaticamente.
4. Criar manualmente tres LOGINs reais exclusivos deste ensaio:
   rt_test_a, rt_test_b e rt_test_none. Atributos LOGIN INHERIT NOSUPERUSER
   NOCREATEDB NOCREATEROLE NOREPLICATION NOBYPASSRLS.
   Conceder rt_persist_client aos tres. Nao conceder postgres, outros papeis
   administrativos, acesso direto a tabelas ou CREATE no esquema.
   Definir senhas fortes por mecanismo seguro, por exemplo comando interativo
   psql \\password, sem SQL de senha em historico ou logs.
   Todos possuem a API do contrato; rt_test_none nao recebe escopo de dados.
   Os escopos ficticios de A/B sao criados pelo ensaio, nao nesta preparacao.
5. Apos confirmar novamente o destino, registrar administrativamente o comentario
   do banco: acqua-gestao-homologacao:<marcador-de-32-hex>.
   Exemplo de forma SQL (substituir marcador localmente antes de executar):
   COMMENT ON DATABASE postgres IS 'acqua-gestao-homologacao:MARCADOR';
   Nao substituir comentario existente sem revisar seu uso.
6. Obter CA adequada ao endpoint via fonte oficial Supabase. Configurar
   ca_file absoluto; nao aceitar sslmode=require como alternativa.
7. Abrir cada LOGIN separadamente e conferir session_user=current_user=LOGIN,
   current_database()=postgres e TLS em pg_stat_ssl. Nao usar SET ROLE nem
   SET SESSION AUTHORIZATION. O administrador nao e tratado como superusuario.

Instalacao SQL manual sugerida, depois de criar um servico libpq local protegido:
    psql "service=acqua_rt_homolog_admin" -X -v ON_ERROR_STOP=1 -f database/rt/001_contract.sql
Esse servico e usado apenas pelo operador/psql; o runner usa o manifesto explicito.
Tambem exige verify-full/CA no servico psql. Nao confundir SSL ativo no servidor
com validacao de certificado pelo cliente: ambos sao necessarios.

## Inspecao e execucao

    python -m pip install -r database/rt/requirements-test.txt
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --inspect

Esperado: INSPECTION_OK. A inspecao nao altera fixtures ou esquema.
Confere TLS/identidade em cada conexao, versao 1, durabilidade, ACLs PUBLIC/Data API
e ausencia de privilegios diretos dos LOGINs. E uma barreira especifica do contrato,
nao uma auditoria completa de todos os papeis/funcoes do projeto Supabase.

Depois, com autorizacao para inserir somente dados ficticios:
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --run-fixtures

Exige tabelas RT vazias (exceto schema_migrations) e nao faz limpeza.
Produz os mesmos 15 nomes de cenarios do executor original; somente bootstrap,
conexao e subprocesso de recuperacao sao adaptados. Concorrencia usa conexoes
reais independentes. O teste de imutabilidade tenta DELETE e espera rejeicao;
nenhuma exclusao de fixtures e autorizada/implementada.
Falhas nao imprimem traceback, DSN, mensagens libpq, payload ou senha: mostram
nome do teste e classe/SQLSTATE. Nao ativar logging SQL sensivel no cliente.

As fixtures permanecem no banco, inclusive apos falha parcial. Reexecutar a suite
completa sera recusado. Uma limpeza ou reinstalacao exige plano e confirmacao
separados. A recuperacao pode ser repetida sem reinstalar:
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --recover
Esperado: RECOVERY_OK, sem arquivo local de rascunho.
O replay pode confirmar uma operacao existente; nao e comando de limpeza.
Nao ha reinicio automatico de projeto Supabase. Teste de backup/restauracao e
reinicio do servico sao ensaios separados, ainda nao executados neste provedor.

## Validacao isolada e limites

O workflow rt-supabase-executor.yml usa PostgreSQL 16 descartavel, TLS com
certificado DNS:localhost e LOGINs reais de senhas exclusivamente ficticias.
O script local_supabase_ci.py aplica o esquema APENAS ao destino local fixo
localhost:5432/rt_contract_test. Esse bootstrap e separado e nunca aceita URL.
O runner --local tambem exige marcador/TLS/identidade e o destino local exato.
Nunca usar --local com tunnel para um banco remoto.

Os testes adicionais verificam allowlist, recusa de opcoes desconhecidas,
sanitizacao de erro e rejeicao de certificado/CA inadequados. O CI tambem tenta
repetir fixtures, exige recusa e confirma que a revisao anterior continua legivel.
Nenhum segredo do repositorio e usado e nenhuma conexao Supabase e tentada.
A suite original de 15 cenarios e herdada, sem alterar test_contract.py.

Limites: resultado local nao comprova disponibilidade IPv6, CA, privilegios ou
politicas do Supabase. Plano Free nao estabelece backup/PITR adequado a producao.
O runner usa LOGINs de teste por identidade; nao integra autenticacao Streamlit.
Nenhuma etapa 3 a 9, exportacao Sheets ou migracao de dados reais e implementada.


## Temporary Access (JIT) para os tres papeis de teste

Modo opt-in, sem alterar o contrato RT, as memberships ou os 15 cenarios.
Somente `rt_test_a`, `rt_test_b` e `rt_test_none` usam JIT. `postgres` continua
pelo passfile administrativo privado existente, sem token e sem `jit=true`.
O modo anterior continua sendo o padrao (`test_auth_mode="passfile"`).

O modo JIT exige o projeto `xfumoyxmhpxmqidmxjhq`, host exato
`aws-1-sa-east-1.pooler.supabase.com`, banco `postgres`, porta 5432 e
`connection_mode="session_pooler"`. Nao aceita destino local, conexao direta,
Transaction Pooler ou outro indice/regiao/ref. As verificacoes existentes de
marcador, identidade SQL real, CA, TLS do cliente/backend, timeouts por SQL,
READ ONLY nas inspecoes e permissoes permanecem ativas.

Preparar futuramente um manifesto privado separado, fora de ambos os worktrees,
copiando os campos ja validados do manifesto do Session Pooler. Preservar o
passfile anterior, que neste modo e usado somente pelo administrador. Adicionar:

```toml
test_auth_mode = "jit_dpapi"
jit_token_file = 'C:\Users\SEU_USUARIO\AppData\Local\AcquaGestao\secrets\rt-tests.dpapi'
```

O arquivo apontado deve existir fora do repositorio e ter owner igual ao usuario
Windows que executa o processo. A DACL deve conceder acesso somente a esse usuario,
SYSTEM e/ou Administradores locais. Links/junctions, reparse points, DACL aberta,
ACLs nao suportadas e arquivos maiores que 64 KiB sao recusados. Nenhuma ACL e
alterada pelo executor. O formato aceito e DPAPI Windows sem entropia adicional:
blob binario ou hexadecimal de `ConvertFrom-SecureString` com DPAPI (sem `-Key`),
em arquivo UTF-8 ou UTF-16. O texto protegido deve ser um PAT `sbp_...` completo,
sem espacos, quebras de linha ou wrappers. JWT e arquivo de token em texto puro
nao sao aceitos. Os testes usam apenas um valor sintetico nao emitido pelo Supabase.

Cada conexao de teste descriptografa pelo Windows no processo atual e passa o PAT
apenas ao argumento de memoria `password` do driver. Envia `options="-c jit=true"`
e o usuario de roteamento `<papel>.<ref>`; exige `session_user=current_user=<papel>`.
O passfile da conexao JIT e explicitamente desabilitado (`os.devnull`). Ausencia,
formato invalido, ACL inadequada, erro DPAPI ou recusa do servidor interrompem a
conexao; nao ha fallback, criacao de senha, geracao/renovacao de token ou concessao.
Um token com sintaxe valida ainda pode estar expirado/revogado: o servidor decide
a autenticacao. O executor nao consulta nem modifica a Management API.

O PAT nao e armazenado em globais, argv, ambiente, SQL, logs ou arquivos em texto
puro. O buffer nativo DPAPI e zerado antes de `LocalFree`; as referencias temporarias
sao liberadas apos conectar. Python/libpq podem reter copias imutaveis em memoria;
nao ha promessa de apagamento forense. Nao habilitar debugger, tracing, dumps,
logs do driver ou gravacao de memoria durante uma futura execucao real.

A recuperacao em novo processo recebe somente o caminho do mesmo manifesto. O
filho relera o modo JIT e descriptografara o arquivo sob o mesmo usuario Windows.
O ambiente do filho exclui defaults `PG*` e variaveis conhecidas de tokens. Nenhum
token e transmitido por argv/ambiente. A assercao original `RECOVERY_OK`, a revisao
unica e a idempotencia do teste de recuperacao continuam inalteradas.

### Comandos preparados; nao executados remotamente nesta etapa

Antes de qualquer comando remoto futuro: obter autorizacao, preparar nova concessao
exclusiva para os tres papeis com prazo/IP restritos, confirmar validade/IP e habilitar
LOGIN como acao autorizada separada, mantendo senhas nulas. O token e a concessao
do piloto revogados nao servem para esta etapa. Este executor nao habilita/desabilita
LOGIN e nao revoga concessoes; a limpeza administrativa e uma etapa separada.

No diretorio `database/rt` do worktree JIT, com as dependencias de teste instaladas:

```powershell
python -B -m unittest -v test_supabase_guard test_supabase_jit
```

Para uma futura rodada autorizada, na raiz do worktree JIT:

```powershell
$jitManifest = Join-Path $env:LOCALAPPDATA 'AcquaGestao\rt-homologacao\authorization.jit.toml'
python database/rt/supabase_runner.py --config $jitManifest --inspect
# Somente apos INSPECTION_OK e confirmacao de tabelas vazias; executar uma vez:
python database/rt/supabase_runner.py --config $jitManifest --run-fixtures
```

`--run-fixtures` inclui a recuperacao em novo processo entre os 15 testes; nao
repetir a suite nem iniciar uma recuperacao extra para contornar falhas. Na falha,
preservar dados e resultados sanitizados. Nesta adaptacao foram executados somente
testes offline; nenhum SQL remoto, fixture, senha, token real ou concessao foi usado.

Referencias oficiais:
- https://supabase.com/docs/guides/platform/temporary-access
- https://learn.microsoft.com/en-us/windows/win32/api/dpapi/nf-dpapi-cryptunprotectdata

### Reexecucao seletiva (`--run-selected`) — somente os tres cenarios recusados

Na rodada de 2026-11-04 (INSPECTION_OK, 15 cenarios, 44,8 s) 12 passaram e 3 terminaram
em ERROR com `Refused` na abertura da conexao: `test_setting_cannot_forge_authorization`,
`test_simultaneous_same_operation_is_one_revision` e `test_stale_retry_cannot_regress_head`.
Os logs do Pooler registraram "A solicitacao JIT falhou para o usuario rt_test_a: erro de
API inesperado" (21:21:50-52); a recusa ocorreu na validacao JIT do servico, antes do SQL.

    python database/rt/supabase_runner.py --config $jitManifest --run-selected

- Executa `--inspect` completo e depois somente `SELECTED` (exatamente 3 cenarios).
- Os metodos sao herdados de `ContractTests` sem alteracao; assercoes identicas.
- Pre-verificacao somente leitura exige as fixtures da rodada anterior (empresas,
  condominios, escopos, marcador de recuperacao) e recusa escopos de `rt_test_none`.
  Nao insere fixtures-base, nao apaga, nao limpa.
- Cada cenario cria seu proprio condominio ficticio novo (`setUp` original, UUID aleatorio).
- Antes de executar, registra md5 de cada linha RT existente; ao final exige que todas
  continuem presentes (dados da rodada anterior preservados), senao falha.
- Espaca aberturas JIT em 2 s, sem retry: uma recusa continua sendo falha.
- Nao executa a recuperacao em novo processo nem os demais 12 cenarios.
- Concessao, LOGIN temporario, limpeza e revogacao continuam etapas administrativas separadas.

### Resultado registrado da homologacao (2026-11-04)

- Rodada 1 (`--run-fixtures`, 15 cenarios, 44,8 s): 12 ok, 3 ERROR por recusa JIT do
  servico na abertura de conexao (sem SQLSTATE; sem execucao de assercao).
- Rodada 2 (`--run-selected`, 3 cenarios, 28,6 s): 3 ok; dados da rodada 1 preservados.
- Consolidado: 15 de 15 cenarios unicos aprovados contra a homologacao; 0 falhas; 0 sem execucao.
- Ao final de cada rodada: LOGINs `rt_test_*` NOLOGIN, senha nula, zero sessoes;
  regra temporaria excluida e token revogado (API 401).
- Ambiente local das rodadas: Windows, Python 3.14, `psycopg[binary]` 3.2.13 em venv
  temporario (3.2.6 nao publica wheel para 3.14). O projeto, o devcontainer e o CI usam
  Python 3.12 com `psycopg[binary]==3.2.6`; o pin de `requirements-test.txt` nao foi alterado.
