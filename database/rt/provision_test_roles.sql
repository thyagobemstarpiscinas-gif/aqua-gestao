-- SOMENTE PARA REVISAO. Nao executar sem autorizacao especifica.
-- Executar via psql direto, TLS verify-full, com ON_ERROR_STOP.
-- Requer SET rt_homolog.expected_marker='32 hex'; na MESMA sessao.
-- Nao contem senhas. Roles sao criadas com NOLOGIN ate configuracao privada.
BEGIN;
DO $guard$
DECLARE expected text := current_setting('rt_homolog.expected_marker', true);
BEGIN
    IF current_database() <> 'postgres' OR session_user <> 'postgres'
       OR current_user <> session_user THEN
        RAISE EXCEPTION 'Destino ou identidade administrativa inesperados';
    END IF;
    IF expected IS NULL OR expected !~ '^[a-f0-9]{32}$' THEN
        RAISE EXCEPTION 'Marcador esperado nao configurado nesta sessao';
    END IF;
    IF (SELECT shobj_description(oid,'pg_database') FROM pg_database
        WHERE datname=current_database()) IS DISTINCT FROM
       'acqua-gestao-homologacao:' || expected THEN
        RAISE EXCEPTION 'Marcador do banco nao corresponde a homologacao autorizada';
    END IF;
    IF NOT coalesce((SELECT ssl FROM pg_stat_ssl WHERE pid=pg_backend_pid()), false) THEN
        RAISE EXCEPTION 'TLS obrigatorio; verify-full tambem deve ser confirmado no cliente';
    END IF;
    IF to_regnamespace('rt_persist') IS NULL THEN
        RAISE EXCEPTION 'Esquema RT ausente; este roteiro nao aplica migracoes';
    END IF;
    IF (SELECT array_agg(version ORDER BY version) FROM rt_persist.schema_migrations)
       IS DISTINCT FROM ARRAY[1] THEN
        RAISE EXCEPTION 'Versao da migracao inesperada';
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname='rt_persist_client'
        AND NOT rolcanlogin AND NOT rolsuper AND NOT rolcreatedb
        AND NOT rolcreaterole AND NOT rolreplication AND NOT rolbypassrls) THEN
        RAISE EXCEPTION 'Papel base ausente ou com atributos inesperados';
    END IF;
END;
$guard$;

DO $empty$
DECLARE t text; occupied boolean;
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles
        WHERE rolname IN ('rt_test_a','rt_test_b','rt_test_none')) THEN
        RAISE EXCEPTION 'Papel de teste preexistente: parar, nao reutilizar nem alterar';
    END IF;
    FOREACH t IN ARRAY ARRAY['companies','condominiums','access_scopes','documents',
                             'revisions','save_requests','outbox'] LOOP
        EXECUTE format('SELECT EXISTS(SELECT 1 FROM rt_persist.%I)', t) INTO occupied;
        IF occupied THEN
            RAISE EXCEPTION 'Dados RT preexistentes; nao limpar para provisionar';
        END IF;
    END LOOP;
END;
$empty$;
CREATE ROLE rt_test_a NOLOGIN INHERIT NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION NOBYPASSRLS CONNECTION LIMIT 4;
CREATE ROLE rt_test_b NOLOGIN INHERIT NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION NOBYPASSRLS CONNECTION LIMIT 4;
CREATE ROLE rt_test_none NOLOGIN INHERIT NOSUPERUSER NOCREATEDB NOCREATEROLE NOREPLICATION NOBYPASSRLS CONNECTION LIMIT 4;
-- PostgreSQL 17: herdar EXECUTE, sem SET ROLE e sem delegar membership.
GRANT rt_persist_client TO rt_test_a, rt_test_b, rt_test_none WITH INHERIT TRUE;
GRANT rt_persist_client TO rt_test_a, rt_test_b, rt_test_none WITH SET FALSE;
GRANT rt_persist_client TO rt_test_a, rt_test_b, rt_test_none WITH ADMIN FALSE;
COMMIT;
