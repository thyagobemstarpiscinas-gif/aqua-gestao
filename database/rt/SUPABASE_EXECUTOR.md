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
