-- SOMENTE LEITURA. Executar em CADA conexao real separada (A, B, none).
SELECT session_user,current_user,current_database();
SELECT ssl,version,cipher FROM pg_stat_ssl WHERE pid=pg_backend_pid();
SELECT rolname,rolcanlogin,rolsuper,rolcreatedb,rolcreaterole,rolreplication,rolbypassrls
FROM pg_roles WHERE rolname=session_user;
SELECT has_database_privilege(session_user,current_database(),'CONNECT') AS can_connect,
       has_schema_privilege(session_user,'rt_persist','USAGE') AS schema_usage,
       has_schema_privilege(session_user,'rt_persist','CREATE') AS schema_create;
SELECT c.relname,
       has_table_privilege(session_user,c.oid,
         'SELECT,INSERT,UPDATE,DELETE,TRUNCATE,REFERENCES,TRIGGER') AS any_table_access,
       has_any_column_privilege(session_user,c.oid,
         'SELECT,INSERT,UPDATE,REFERENCES') AS any_column_access
FROM pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace
WHERE n.nspname='rt_persist' AND c.relkind IN ('r','p');
SELECT p.oid::regprocedure AS function,
       has_function_privilege(session_user,p.oid,'EXECUTE') AS can_execute
FROM pg_proc p JOIN pg_namespace n ON n.oid=p.pronamespace
WHERE n.nspname='rt_persist';
SELECT g.rolname,m.admin_option,m.inherit_option,m.set_option
FROM pg_auth_members m JOIN pg_roles g ON g.oid=m.roleid
JOIN pg_roles u ON u.oid=m.member WHERE u.rolname=session_user;
-- Resultado esperado: somente rt_persist_client; inherit=true, admin=false, set=false.
-- Em transacao separada, testar SELECT * FROM rt_persist.revisions LIMIT 0;
-- Deve falhar com 42501. Fazer ROLLBACK apos a falha. Nao testar com administrador.
