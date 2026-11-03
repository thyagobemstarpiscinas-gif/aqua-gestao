# Executor Supabase — instalacao controlada, sem integracao ao aplicativo

Implementacao derivada de 6f731373655c91c9283a8742ceecf56fed4b3e70.
Nenhum comando deste roteiro foi aplicado ao Supabase.
A migracao original e o executor local original permanecem inalterados.
Python 3.11+; dependencias em requirements-test.txt.

## Autorizacao e destino

A execucao remota exige confirmacao do proprietario. Nao enviar senhas, tokens,
pgpass ou strings de conexao pelo chat, Git ou logs. Usar computador controlado.
Conferir no painel: acqua-gestao-homologacao, Free, sa-east-1, Healthy,
Data API desativada. Obter Project Ref e endpoint direto no proprio painel.
O executor suporta somente endpoint direto db.<ref>.supabase.co:5432, banco
postgres. Se nao houver IPv6/acesso direto, interromper; poolers nao foram
autorizados por este executor. Nao relaxar validacao de hostname ou TLS.

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
