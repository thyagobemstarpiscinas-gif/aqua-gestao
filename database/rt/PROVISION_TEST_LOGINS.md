# Provisionamento revisavel dos LOGINs RT — Supabase de homologacao

Referencia: d74f38d6c154427cd0c976780bfd8d097f82be08.
Estado informado pelo proprietario: PostgreSQL 17.11, esquema rt_persist versao 1
instalado, Data API desativada, sem dados reais. Esse estado nao foi verificado
remotamente nesta preparacao. Nenhum SQL deste pacote foi aplicado ao Supabase.

Nao seguir novamente a secao de instalacao inicial de SUPABASE_EXECUTOR.md.
Nao executar 001_contract.sql. Nao alterar Streamlit, Google, PINs ou OAuth.

## Arquivos e sequencia

1. provision_test_roles.sql: valida destino e esquema; cria tres papeis NOLOGIN,
   sem senha, com membership restrito. Uma transacao: erro interrompe tudo.
2. Configuracao privada das senhas, fora de qualquer arquivo SQL versionado.
3. enable_test_logins.sql: repete validacoes, confere atributos/membership e habilita
   LOGIN; nao configura ou le senhas. Reexecucao com LOGIN ja habilitado e recusada.
4. verify_test_logins.sql: consultas somente de leitura em cada conexao real.
5. preview_test_scopes.sql: apenas SELECT com plano dos escopos ficticios.
6. Executor existente: --inspect, depois --run-fixtures quando autorizado.

Os papeis sao rt_test_a (empresa A), rt_test_b (empresa B) e rt_test_none
(sem escopo). Todos herdam a API rt_persist_client; somente a verificacao de
access_scopes determina quais dados podem ler/gravar. Nao sao usuarios Supabase Auth.

## Antes de qualquer aplicacao futura

- Obter autorizacao para provisionamento. Esta entrega serve para revisao.
- Conferir projeto acqua-gestao-homologacao, Project Ref e sa-east-1 no painel.
  Nao abrir ou selecionar o projeto Bem Star Piscinas.
- Conexao DIRETA db.<ref>.supabase.co:5432/postgres, TLS verify-full e CA oficial.
  O executor atual nao aceita pooler. Se nao houver acesso direto/IPv6, parar.
- Usar computador privado e psql local, sem gravacao de terminal, compartilhamento
  de tela, shell tracing (set -x), transcricao automatica ou gravacao de teclado.
- Configurar servicos libpq e passfile FORA do repositorio, arquivos privados
  modo 0600/ACL equivalente. Nao enviar dados de conexao ou senhas pelo chat.
- Confirmar session_user=current_user=postgres na conexao administrativa.
  Nao assumir SUPERUSER. Falha de permissao em CREATE ROLE/GRANT implica parar,
  nao conceder privilegios amplos para contornar.
- Rodar inspect_supabase.sql para reconferir ACLs, inclusive PUBLIC, anon,
  authenticated e service_role. Este pacote nao corrige nem amplia essas ACLs.
- Conferir marcador ja associado ao banco:
  acqua-gestao-homologacao:<32 caracteres hexadecimais>.
  O marcador precisa coincidir com o manifesto privado do executor.
  Se ausente ou divergente, parar e revisar seu cadastro como acao separada.
  Nenhum script aqui muda o comentario do banco.

## Criar os papeis, ainda impedidos de fazer LOGIN

Abrir psql sem historico, em terminal privado:
    PSQL_HISTORY=/dev/null psql "service=acqua_rt_homolog_admin" -X -v ON_ERROR_STOP=1

O servico local deve fixar host/ref correto, database=postgres, user=postgres,
sslmode=verify-full, sslrootcert e passfile. Nenhuma senha na linha de comando.
Usar uma sessao nova, sem PGOPTIONS/PGSERVICE herdados de outro projeto.
Validar \conninfo e a identidade SQL; \conninfo nao substitui conferir o servico
para garantir verify-full. pg_stat_ssl confirma apenas criptografia no servidor.

Na MESMA sessao, configurar o marcador esperado (nao e segredo de autenticacao):
    SET rt_homolog.expected_marker = 'SUBSTITUIR_PELOS_32_HEX_DO_MANIFESTO';
    \i database/rt/provision_test_roles.sql

Placeholder ou marcador incorreto causa falha antes de CREATE ROLE.
Tambem recusa qualquer papel rt_test_* preexistente ou dado em tabelas RT.
Nao modifica/reutiliza papeis existentes. Nao limpa dados para destravar.
Se o psql encerrar com transacao aberta por erro, conexao encerrada reverte a
transacao; em sessao interativa ainda aberta, executar ROLLBACK antes de investigar.

Atributos: NOLOGIN INHERIT NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION
NOBYPASSRLS; limite de quatro conexoes por papel (o ensaio concorrente usa duas).
Membership explicito PG17: INHERIT TRUE, SET FALSE, ADMIN FALSE.
Sem grants diretos em tabelas, colunas, funcoes adicionais ou CREATE no esquema.
CONNECT ao banco normalmente vem de PUBLIC no Supabase: verificar, nao alterar
privilegios globais. Se CONNECT nao existir, revisar concessao somente de CONNECT
aos tres papeis como acao adicional, sem conceder TEMP/CREATE.

Limite de escopo: estes scripts restringem a API RT. Privilegios herdados de PUBLIC
em outros esquemas/bancos nao sao removidos por eles. Nao declarar isolamento
global de todo o projeto apenas porque o esquema RT esta protegido.

## Senhas fortes sem registro

1. Gerar tres senhas DIFERENTES no gerenciador de senhas local confiavel, usando
   gerador criptografico, 32 caracteres aleatorios ou mais. Preferir alfabeto
   alfanumerico aleatorio, para evitar problemas de escape em pgpass.
2. Salvar somente no cofre privado. Nao gerar via SELECT, print, echo, arquivo de
   codigo, argumento de processo, variavel exibida ou pipeline com tee.
3. Na sessao psql PRIVADA, confirmar SCRAM:
       SET password_encryption = 'scram-sha-256';
       SHOW password_encryption;
4. Configurar cada senha pelos prompts sem eco:
       \password rt_test_a
       \password rt_test_b
       \password rt_test_none
   Colar a senha somente no prompt de senha; nunca em ALTER ROLE ... PASSWORD.
   Nao imprimir a tela ou copiar conteudo do prompt para suporte/chat.
5. O psql evita enviar senha em texto claro como comando SQL. O verificador SCRAM
   tambem e material sensivel: conferir previamente a politica de auditoria/logs
   do projeto para nao registrar material de credencial. Se essa garantia nao
   puder ser estabelecida, interromper e definir mecanismo administrativo seguro.
   Nao ler pg_authid para diagnosticar ou copiar hashes.
6. Preencher passfile privado do executor em editor seguro. Sem curingas de
   hostname/banco/usuario; uma entrada para cada LOGIN e para o administrador.
   Sem backup/sincronizacao desse arquivo em pasta publica ou repositorio.
   Limpar a area de transferencia apos uso, conforme recurso do gerenciador.
7. Somente depois dos tres prompts concluidos:
       \i database/rt/enable_test_logins.sql

Se uma senha nao foi configurada, nao habilitar os LOGINs. O script de habilitacao
nao consegue atestar a senha sem ler catalogo protegido: a comprovacao efetiva
e autenticar cada LOGIN separadamente. Senha ausente/incorreta deve falhar fechada.
Se o processo parar depois de habilitar apenas parte dos passos privados, nao
improvisar reexecucao: registrar estado sem segredos e revisar retomada.
Os papeis criados NOLOGIN permanecem bloqueados caso os prompts falhem.

## Verificar cada conexao com identidade real

Criar servicos locais separados: acqua_rt_test_a, acqua_rt_test_b,
acqua_rt_test_none, com o mesmo destino/verify-full/CA e usuarios correspondentes.
Nunca usar SET ROLE ou SET SESSION AUTHORIZATION para simular LOGIN.

Para cada um, abrir nova conexao e executar:
    psql "service=acqua_rt_test_a" -X -v ON_ERROR_STOP=1 -f database/rt/verify_test_logins.sql
Repetir mudando somente o servico para B e none.

Esperado:
- session_user=current_user=LOGIN solicitado; current_database()=postgres.
- SSL ativo, certificado/hostname validados pelo cliente.
- can_connect=true, schema_usage=true, schema_create=false.
- atributos administrativos/replication/BYPASSRLS=false.
- todas as linhas any_table_access e any_column_access=false.
- EXECUTE=true somente save, history e conflicts; funcoes internas=false.
- unica membership de saida: rt_persist_client; sem ADMIN/SET, com INHERIT.

Na propria conexao restrita, provar a negacao sem ler dados:
    BEGIN;
    SELECT * FROM rt_persist.revisions LIMIT 0;
Resultado esperado: permission denied, SQLSTATE 42501; executar ROLLBACK.
Com ON_ERROR_STOP em modo nao interativo, a sessao pode encerrar no erro:
encerramento reverte a transacao. Nao executar esse teste com administrador.
Nao usar tentativas de UPDATE/DELETE como mecanismo de limpeza.

## Escopos e compatibilidade com os 15 testes

Rodar preview_test_scopes.sql para revisar os quatro escopos iniciais:
A tem tres condominios ficticios (inclui marcador de recuperacao), B tem um;
periodo 2026-10-01. none tem zero. O mesmo UUID de condominio em A/B testa
separacao pela identidade composta da empresa, nao compartilhamento de dados.

IMPORTANTE: nao inserir previamente companies/condominiums/access_scopes.
O setUpClass de supabase_runner.py exige as tabelas RT vazias e cria exatamente
esses escopos em transacao. Cada cenario adiciona ainda um condominio UUID novo
para A. Uma carga manual de escopos antes do runner bloquearia a suite.
A previa SELECT prepara a revisao do plano; a criacao efetiva e feita pelo runner.

Depois do provisionamento e somente com autorizacao para os testes ficticios:
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --inspect
Esperado: INSPECTION_OK.
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --run-fixtures
Esperado: 15 testes aprovados; nenhum segredo no log.
    python database/rt/supabase_runner.py --config /caminho/privado/authorization.toml --recover
Esperado: RECOVERY_OK.

Se houver dados/fixtures de ensaio anterior, parar. Nao reaplicar a migracao nem
usar DELETE/DROP/TRUNCATE para permitir nova execucao sem revisao/autorizacao.
Ate o provisionamento e os ensaios reais, nao declarar sucesso no Supabase.
Finalizado o ensaio, revisar desabilitacao dos LOGINs (NOLOGIN) e tratamento das
sessoes abertas em acao separada: NOLOGIN sozinho nao encerra conexoes existentes.
