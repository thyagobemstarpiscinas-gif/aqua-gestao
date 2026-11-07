-- Executar SOMENTE depois de configurar as tres senhas com \password privado.
-- Mesma validacao de destino; nao configura nem imprime senhas.
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

DO $roles$
DECLARE r record;
BEGIN
    IF (SELECT count(*) FROM pg_roles
        WHERE rolname IN ('rt_test_a','rt_test_b','rt_test_none')) <> 3 THEN
        RAISE EXCEPTION 'Provisionamento dos tres papeis incompleto';
    END IF;
    FOR r IN SELECT oid,rolname,rolcanlogin,rolsuper,rolcreatedb,rolcreaterole,
                    rolreplication,rolbypassrls,rolinherit
             FROM pg_roles WHERE rolname IN ('rt_test_a','rt_test_b','rt_test_none') LOOP
        IF r.rolcanlogin OR r.rolsuper OR r.rolcreatedb OR r.rolcreaterole
           OR r.rolreplication OR r.rolbypassrls OR NOT r.rolinherit THEN
            RAISE EXCEPTION 'Atributos divergentes: interromper habilitacao';
        END IF;
        IF (SELECT count(*) FROM pg_auth_members WHERE member=r.oid) <> 1
           OR NOT EXISTS (
             SELECT 1 FROM pg_auth_members m JOIN pg_roles g ON g.oid=m.roleid
             WHERE m.member=r.oid AND g.rolname='rt_persist_client'
               AND m.inherit_option AND NOT m.set_option AND NOT m.admin_option
           ) THEN
            RAISE EXCEPTION 'Membership divergente: interromper habilitacao';
        END IF;
    END LOOP;
END;
$roles$;
ALTER ROLE rt_test_a LOGIN;
ALTER ROLE rt_test_b LOGIN;
ALTER ROLE rt_test_none LOGIN;
COMMIT;
