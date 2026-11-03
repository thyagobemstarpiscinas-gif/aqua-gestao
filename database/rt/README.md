# Persistência RT — contrato e PostgreSQL isolado (Etapas 1 e 2)

Base: 9f7053c6da538fca38a4e1aa75ce9d05918423f3.
Sem integração com Streamlit, Sheets, Drive ou autenticação do aplicativo.
Sem migração de dados existentes, alterações no PR #14 ou implantação.
Esta pasta implementa uma prova transacional; não autoriza Etapas 3 a 9.

## Contrato

Identidade: (company_id UUID, condominium_id UUID, period DATE no primeiro dia
do mês). IDs são provisionados e estáveis; nomes não autorizam acesso.
Um documento RT por identidade. O formato do conteúdo é JSONB schema_version=1
com objeto campos. Nenhum campo legado é removido. Fotos permanecem referências;
nenhum upload ou download é feito aqui. Limite inicial: 8 MiB por conteúdo,
a validar antes da futura migração. O envelope relacional prevalece para identidade.
Um campo identity, se enviado, precisa coincidir exatamente com o envelope.
Metadados legados de exibição não substituem esse envelope. O futuro adaptador
deverá validar a associação dos campos legados ao cadastro antes de salvar.

Operação SQL:
SELECT rt_persist.save(company_uuid, condominium_uuid, period_date,
                      operation_uuid, base_revision, payload_jsonb,
                      restored_from_revision_or_null);

- Base 0: primeira revisão. Base atual: grava revisão atual + 1.
- Base antiga: preserva payload em save_requests com outcome=conflict.
- Base futura: rejeita como contrato inválido, sem gravar.
- Mesmo operation UUID, identidade, autor, base e conteúdo: mesmo resultado.
- Reutilização da operação com outro conteúdo/base/restauração/autor: rejeita.
- Uma operação é escopada por documento. Outra identidade é outra operação lógica.
- O lock FOR UPDATE no documento serializa edição, retentativa e criação.
- Revisão, head, resultado idempotente e outbox confirmam numa transação.
- Conflito retorna normalmente: COMMIT obrigatório para preservar o candidato.
- O resultado só pode ser exibido como salvo após COMMIT bem-sucedido.
- Resultado incerto no COMMIT: repetir o mesmo operation UUID; não gerar outro.
- Funções suportam READ COMMITTED. Em isolamento mais forte, tratar SQLSTATE
  40001/40P01 repetindo a transação inteira com o mesmo identificador.
- Revisões e solicitações têm triggers contra UPDATE/DELETE. Runtime não tem
  acesso direto a tabelas, TRUNCATE ou funções administrativas.
- history() recupera todas as revisões; conflicts() recupera candidatos divergentes.
  Paginação será necessária antes de uso em grande volume.
- Restaurar cria revisão nova com restored_from da mesma identidade. A escolha
  e resolução de conflitos na interface pertencem a etapas futuras.

Estados de salvamento: unconfirmed (cliente), saved (COMMIT confirmado),
conflict (candidato confirmado no banco), unavailable/unknown (cliente não tem
confirmação). Nenhum retorno local equivale a saved.
Estados de sincronização reservados: pending, sending, retry, confirmed.
Nesta etapa somente pending é produzido. Não há consumidor nem exportador.
O resultado idempotente registra o estado no momento do save; não é uma consulta
ao estado futuro da exportação.

## Segurança e permissões mínimas

Aplicar migração usando administrador somente em banco vazio de homologação.
001_contract.sql é uma migração única e transacional, com versão registrada;
reaplicação falha sem alterar o esquema existente.
O administrador cria roles e funções SECURITY DEFINER, com search_path fixo.
Não usar esse administrador como conexão de runtime.

rt_persist_client é papel NOLOGIN: só USAGE no esquema e EXECUTE nas três funções
públicas do contrato. Nenhum SELECT/INSERT/UPDATE/DELETE/TRUNCATE nas tabelas.
Provisionar um LOGIN PostgreSQL restrito, INHERIT e sem privilégios administrativos,
conceder rt_persist_client e registrar escopos em access_scopes.
A autorização usa session_user, não current_setting, parâmetro do navegador,
SET ROLE ou nome enviado no JSON. Cada escopo permite uma combinação exata de
empresa, condomínio e período. Cadastro e concessão de escopos são administrativos.

Não expor rt_persist na Data API/Supabase REST; não usar anon, authenticated ou
service_role como conexão deste contrato. Sem necessidade de Supabase Auth.
Escolher conexão direta ou pooler que preserve session_user como o LOGIN real.
Validar SELECT session_user na conexão antes de homologar; se o pooler substituir
a identidade, o mecanismo deve falhar fechado e o endpoint é inadequado.
Nunca conceder CREATE no esquema ou membership no papel proprietário ao runtime.
As funções continuam vulneráveis a um administrador comprometido: triggers não
substituem backup ou controles administrativos. Não há garantia de auditoria
individual com identidade compartilhada entre pessoas.

O papel de conexão por escopo é uma prova de isolamento no banco. A futura
integração com autenticação existente precisa de projeto próprio para mapear
usuários autorizados a escopos sem deixar o navegador escolher credenciais.
Não conectar uma credencial multiescopo e alegar isolamento individual automático.

## Executar PostgreSQL local efêmero

Não há dependência das credenciais atuais. Use somente os valores fictícios abaixo.
O runner recusa bancos diferentes de rt_contract_test e hosts fora de loopback.
Também recusa esquema rt_persist já existente; não apaga dados para repetir testes.

Em ambiente de desenvolvimento com Docker:
    docker run --name rt-contract-test --rm -d -p 127.0.0.1:55432:5432 \
      -e POSTGRES_USER=postgres -e POSTGRES_PASSWORD=ephemeral-test-only \
      -e POSTGRES_DB=rt_contract_test postgres:16
    python -m pip install -r database/rt/requirements-test.txt
    export RT_TEST_DSN='postgresql://postgres:ephemeral-test-only@127.0.0.1:55432/rt_contract_test'
    python database/rt/test_contract.py
    docker restart rt-contract-test
    python database/rt/test_contract.py --verify-restart

Aguardar pg_isready antes do primeiro teste e depois do restart.
Ao terminar, remover apenas o container efêmero criado para este ensaio.
A configuração Docker deve seguir as regras de proxy/CA do ambiente executor.
A suíte usa SET SESSION AUTHORIZATION com administrador LOCAL para simular LOGINs;
runtime real não recebe essa permissão.

GitHub Actions executa somente estes testes na branch dedicada, com PostgreSQL 16
efêmero e sem repository Secrets. Não executa a suíte completa nem publica nada.
Disponibilidade e eventuais minutos de Actions dependem das políticas existentes;
não contratar plano ou habilitar cobrança para executar.

## Supabase de homologação — configuração futura, não executada

1. Proprietário cria projeto exclusivo de homologação, sem contratação paga.
   Se plano gratuito não estiver disponível, parar e decidir alternativa local.
2. Selecionar região adequada e confirmar retenção/limites/pausa do plano.
   Não assumir backups/PITR em plano gratuito.
3. Usar SQL Editor administrativo para aplicar 001_contract.sql em banco vazio.
4. Criar LOGIN restrito e conceder rt_persist_client; cadastrar somente empresas,
   condomínios e escopos fictícios. Senha forte gerada fora do repositório/chat.
5. Guardar conexão em gerenciador de segredos próprio da homologação; não alterar
   Secrets do app atual. Conexão TLS verificada (sslmode=verify-full com CA adequada).
6. Manter esquema privado fora da lista de schemas expostos da Data API.
7. Validar session_user, permissões negadas e testes equivalentes em ambiente
   dedicado. O runner local propositalmente NÃO aceita URL Supabase.
8. Documentar backup lógico externo e fazer restauração antes de considerar uso.
   Nenhum projeto Supabase ou credencial foi criado por estes arquivos.

## Recuperação e reversão

A fonte é PostgreSQL, com synchronous_commit=on, fsync=on e armazenamento durável.
Arquivo local/Streamlit não participa do protocolo. Novas conexões e novos processos
recuperam histórico e conflitos pelas funções; retentativas preservam UUID original.
Teste --verify-restart verifica uma revisão fictícia fixa e sua idempotência após
reinício do banco. O teste em novo processo usa /tmp e não lê arquivos de rascunho.

Backup administrativo: pg_dump no formato custom, com esquema rt_persist e
conteúdo; proteger/encriptar fora da instância. Roles são globais e não fazem parte
do pg_dump comum: documentar/recriar rt_persist_client e LOGINs restritos sem
exportar senhas para logs ou Git. Restaurar primeiro em banco isolado, preservando
ownership/ACLs e funções; conferir session_user, escopos, heads, hashes, revisões,
save_requests e outbox, além de repetir as consultas de recuperação.
Backups configurados e teste pg_restore ainda são requisitos de homologação futura;
o ensaio de restart não comprova recuperação de desastre/região.

Não usar DROP SCHEMA CASCADE como rollback quando houver dados. Reversão destas
etapas: manter banco intacto e não integrar ao aplicativo. Para ensaio descartável,
remover somente banco/container identificado como fictício. Não há down migration
destrutiva automatizada. Migração de rascunhos reais pertence a etapa posterior.

## Critérios e limites da prova

Os testes cobrem oito disputas simultâneas, repetição simultânea da mesma operação,
resposta descartada depois do commit, retentativa obsoleta, restauração,
rollback integral, outbox transacional, acessos cruzados e imutabilidade.
Não comprovam entrega no Sheets, migração real, backup gerenciado, offline no
navegador, nem atribuição individual pelo PIN atual. Nada que não chegou ao banco
pode ser recuperado por este contrato. Etapas 3 a 9 continuam não autorizadas.
