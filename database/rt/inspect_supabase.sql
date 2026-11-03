-- SOMENTE LEITURA. Executar administrativamente no projeto de homologacao.
SELECT current_database(), session_user, current_user, version();
SELECT ssl, version, cipher FROM pg_stat_ssl WHERE pid=pg_backend_pid();
SELECT to_regnamespace('rt_persist') AS schema_rt;
SELECT rolname, rolcanlogin, rolsuper, rolcreatedb, rolcreaterole, rolbypassrls
FROM pg_roles WHERE rolname IN
('postgres','rt_persist_client','rt_test_a','rt_test_b','rt_test_none',
 'anon','authenticated','service_role');
-- Pode ser executado antes da migracao: retorna zero objetos RT.
SELECT n.nspname, pg_get_userbyid(n.nspowner) AS owner, n.nspacl
FROM pg_namespace n WHERE n.nspname='rt_persist';
SELECT c.relname, pg_get_userbyid(c.relowner) AS owner, c.relacl
FROM pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace
WHERE n.nspname='rt_persist' AND c.relkind IN ('r','p');
SELECT p.oid::regprocedure AS function, pg_get_userbyid(p.proowner) AS owner,
       p.prosecdef, p.proconfig, p.proacl
FROM pg_proc p JOIN pg_namespace n ON n.oid=p.pronamespace
WHERE n.nspname='rt_persist';
-- Privilegios efetivos dos papeis Data API, incluindo heranca e PUBLIC.
SELECT r.rolname, n.nspname,
       has_schema_privilege(r.oid,n.oid,'USAGE') AS usage,
       has_schema_privilege(r.oid,n.oid,'CREATE') AS create_schema_objects
FROM pg_roles r CROSS JOIN pg_namespace n
WHERE r.rolname IN ('anon','authenticated','service_role') AND n.nspname='rt_persist';
SELECT r.rolname,c.relname,
       has_table_privilege(r.oid,c.oid,'SELECT,INSERT,UPDATE,DELETE,TRUNCATE,REFERENCES,TRIGGER')
         AS any_table_privilege,
       has_any_column_privilege(r.oid,c.oid,'SELECT,INSERT,UPDATE,REFERENCES')
         AS any_column_privilege
FROM pg_roles r CROSS JOIN pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace
WHERE r.rolname IN ('anon','authenticated','service_role')
  AND n.nspname='rt_persist' AND c.relkind IN ('r','p');
SELECT r.rolname,p.oid::regprocedure AS function,
       has_function_privilege(r.oid,p.oid,'EXECUTE') AS execute
FROM pg_roles r CROSS JOIN pg_proc p JOIN pg_namespace n ON n.oid=p.pronamespace
WHERE r.rolname IN ('anon','authenticated','service_role') AND n.nspname='rt_persist';
-- Defaults globais ou por esquema podem conceder acesso a objetos futuros.
SELECT pg_get_userbyid(d.defaclrole) AS owner,
       coalesce(n.nspname,'GLOBAL') AS schema, d.defaclobjtype, d.defaclacl
FROM pg_default_acl d LEFT JOIN pg_namespace n ON n.oid=d.defaclnamespace
WHERE d.defaclnamespace=0 OR n.nspname='rt_persist';
