"""Executor explicito: esquema ja instalado, LOGINs reais, TLS e fixtures novas."""
import argparse
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import tempfile
import threading
import time
import tomllib
import unittest

import psycopg
from psycopg import sql
from psycopg.pq import TransactionStatus
import test_contract as base
import jit_dpapi

CONFIG = None
CONFIG_PATH = None
LOCAL = False
ROLES = ("rt_test_a", "rt_test_b", "rt_test_none")
# Somente os tres cenarios recusados na abertura de conexao JIT da rodada de 2026-11-04.
SELECTED = ("test_setting_cannot_forge_authorization",
            "test_simultaneous_same_operation_is_one_revision",
            "test_stale_retry_cannot_regress_head")
RT_TABLES = ("companies", "condominiums", "access_scopes", "documents",
             "revisions", "save_requests", "outbox")
# Intervalo minimo entre aberturas JIT; 0 preserva o comportamento dos outros modos.
JIT_MIN_INTERVAL = 0.0
SELECTED_JIT_INTERVAL = 2.0
_JIT_PACE = threading.Lock()
_JIT_LAST = [float("-inf")]


def pace_jit():
    """Espaca autenticacoes JIT sem retry: falha continua sendo falha."""
    if JIT_MIN_INTERVAL <= 0:
        return
    with _JIT_PACE:
        wait = _JIT_LAST[0] + JIT_MIN_INTERVAL - time.monotonic()
        if wait > 0:
            time.sleep(wait)
        _JIT_LAST[0] = time.monotonic()


class Refused(RuntimeError):
    def __init__(self, message, *, evidence=None):
        super().__init__(message)
        # Somente indicadores booleanos calculados pelo executor, nunca credenciais.
        self.evidence = evidence or {}


def private_file(path):
    p = Path(path)
    if not p.is_absolute() or not p.is_file() or p.is_symlink():
        raise Refused("arquivo privado invalido")
    if os.name != "nt" and (p.stat().st_uid != os.getuid()
                            or stat.S_IMODE(p.stat().st_mode) & 0o077):
        raise Refused("arquivo privado precisa pertencer ao usuario e usar modo 0600")
    return str(p)


def validate(cfg, local=False):
    expected = {"environment", "project_name", "project_ref", "host", "port",
                "database", "marker", "ca_file", "passfile", "admin_user"}
    if set(cfg) - {"connection_mode", "test_auth_mode", "jit_token_file"} != expected:
        raise Refused("campos de autorizacao ausentes ou desconhecidos")
    mode = cfg.get("connection_mode", "direct")
    if mode not in ("direct", "session_pooler"):
        raise Refused("modo de conexao nao autorizado")
    if cfg["environment"] != "homologacao" or cfg["project_name"] != "acqua-gestao-homologacao":
        raise Refused("homologacao nao autorizada")
    if not isinstance(cfg["port"], int) or isinstance(cfg["port"], bool) or not 1 <= cfg["port"] <= 65535:
        raise Refused("porta invalida")
    if not isinstance(cfg["marker"], str) or not re.fullmatch(r"[a-f0-9]{32}", cfg["marker"]):
        raise Refused("marcador de homologacao invalido")
    if local:
        if mode != "direct":
            raise Refused("pooler nao permitido no destino local")
        if (cfg["project_ref"], cfg["host"], cfg["database"]) != (
                "local-ci", "localhost", "rt_contract_test"):
            raise Refused("destino local invalido")
    else:
        if not isinstance(cfg["project_ref"], str) or not re.fullmatch(r"[a-z0-9]{20}", cfg["project_ref"]):
            raise Refused("project ref invalido")
        if cfg["database"] != "postgres":
            raise Refused("banco fora da autorizacao")
        if mode == "direct":
            if cfg["host"] != "db." + cfg["project_ref"] + ".supabase.co":
                raise Refused("endpoint direto fora da autorizacao")
        elif not isinstance(cfg["host"], str) or not re.fullmatch(
                r"aws-[0-9]+-sa-east-1\.pooler\.supabase\.com", cfg["host"]):
            raise Refused("Session Pooler fora da regiao de homologacao autorizada")
        if cfg["port"] != 5432:
            raise Refused("somente porta 5432; Transaction Pooler nao autorizado")
    if cfg["admin_user"] != "postgres":
        raise Refused("administrador de homologacao inesperado")
    auth = cfg.get("test_auth_mode", "passfile")
    if auth not in ("passfile", "jit_dpapi"):
        raise Refused("modo de autenticacao de teste nao autorizado")
    if auth == "jit_dpapi":
        if (local or mode != "session_pooler" or cfg["port"] != 5432
                or cfg["project_ref"] != "xfumoyxmhpxmqidmxjhq"
                or cfg["host"] != "aws-1-sa-east-1.pooler.supabase.com"):
            raise Refused("JIT restrito ao Session Pooler autorizado de homologacao")
        if (not isinstance(cfg.get("jit_token_file"), str)
                or not Path(cfg["jit_token_file"]).is_absolute()):
            raise Refused("arquivo DPAPI absoluto obrigatorio para JIT")
    elif "jit_token_file" in cfg:
        raise Refused("arquivo DPAPI exige modo JIT explicito")
    for key in ("ca_file", "passfile"):
        if not isinstance(cfg[key], str) or not Path(cfg[key]).is_absolute():
            raise Refused("caminhos absolutos obrigatorios")
    return cfg


def load_config(path, local=False):
    global CONFIG, CONFIG_PATH, LOCAL
    CONFIG_PATH = private_file(str(Path(path).resolve()))
    with open(CONFIG_PATH, "rb") as stream:
        CONFIG = validate(tomllib.load(stream), local)
    LOCAL = local
    private_file(CONFIG["passfile"])
    if not Path(CONFIG["ca_file"]).is_file():
        raise Refused("CA ausente")
    # Nao aceitar defaults libpq ocultos por variaveis de ambiente.
    for key in list(os.environ):
        if key.startswith("PG"):
            os.environ.pop(key)


def connection_user(user):
    """Nome de roteamento do pooler; a identidade SQL continua sendo o LOGIN."""
    if user not in (*ROLES, CONFIG["admin_user"]):
        raise Refused("LOGIN fora da lista autorizada")
    if CONFIG.get("connection_mode", "direct") == "session_pooler":
        return user + "." + CONFIG["project_ref"]
    return user


def verify_session_settings(c, *, read_only=False):
    effective = c.execute(
        "SELECT current_setting('statement_timeout')::interval = interval '20 seconds',"
        "current_setting('lock_timeout')::interval = interval '10 seconds',"
        "current_setting('transaction_read_only'),"
        "current_setting('default_transaction_read_only')"
    ).fetchone()
    if not effective or effective[:2] != (True, True):
        raise Refused("timeouts efetivos de sessao divergentes")
    if read_only and effective[2:] != ("on", "on"):
        raise Refused("inspecao exige transacoes somente de leitura")


def begin_inspection(c):
    """Revalidar a protecao antes de consultar em cada nova transacao."""
    if c.info.transaction_status != TransactionStatus.IDLE:
        raise Refused("inspecao exige nova transacao explicita")
    c.execute("BEGIN READ ONLY")
    verify_session_settings(c, read_only=True)


def connect(principal=None, *, read_only=False):
    if CONFIG is None:
        raise Refused("configuracao nao carregada")
    user = principal or CONFIG["admin_user"]
    if user not in (*ROLES, CONFIG["admin_user"]):
        raise Refused("LOGIN fora da lista autorizada")
    c = None
    try:
        credentials = dict(passfile=CONFIG["passfile"], options="")
        if CONFIG.get("test_auth_mode", "passfile") == "jit_dpapi" and user in ROLES:
            # Revalidar o manifesto tambem nas chamadas programaticas. Sem fallback.
            validate(CONFIG, LOCAL)
            pace_jit()
            credentials = dict(password=jit_dpapi.read_token(CONFIG["jit_token_file"]),
                               passfile=os.devnull, options="-c jit=true")
        try:
            c = psycopg.connect(
                host=CONFIG["host"], port=CONFIG["port"], dbname=CONFIG["database"],
                user=connection_user(user), sslmode="verify-full",
                sslrootcert=CONFIG["ca_file"], connect_timeout=10,
                application_name="acqua-rt-homolog-contract",
                # Nao confiar em GUCs de startup: o Session Pooler pode ignora-las.
                prepare_threshold=None, autocommit=True, **credentials)
        finally:
            credentials.clear()
        identity = c.execute(
            "SELECT session_user,current_user,current_database(),"
            "shobj_description(oid,'pg_database') "
            "FROM pg_database WHERE datname=current_database()"
        ).fetchone()
        marker = "acqua-gestao-homologacao:" + CONFIG["marker"]
        if identity != (user, user, CONFIG["database"], marker):
            raise Refused("identidade ou marcador de destino divergente")
        tls = c.execute("SELECT ssl FROM pg_stat_ssl WHERE pid=pg_backend_pid()").fetchone()
        evidence = dict(identity_matches=True, marker_matches=True,
                        client_tls=bool(c.pgconn.ssl_in_use),
                        client_verify_full=c.info.get_parameters().get("sslmode") == "verify-full",
                        backend_ssl_row_present=bool(tls),
                        backend_tls=bool(tls and tls[0] is True))
        if not evidence["client_tls"] or not evidence["client_verify_full"]:
            raise Refused("TLS verify-full do cliente nao confirmado", evidence=evidence)
        if not evidence["backend_tls"]:
            raise Refused("TLS do backend nao confirmado", evidence=evidence)
        # Somente SQL de sessao, depois das barreiras de identidade/marcador/TLS.
        c.execute("SET SESSION statement_timeout = '20s'")
        c.execute("SET SESSION lock_timeout = '10s'")
        if read_only:
            c.execute("SET SESSION default_transaction_read_only = on")
            begin_inspection(c)
        else:
            verify_session_settings(c)
        if principal:
            flags = c.execute(
                "SELECT rolcanlogin,rolsuper,rolcreatedb,rolcreaterole,rolreplication,rolbypassrls "
                "FROM pg_roles WHERE rolname=session_user").fetchone()
            if flags != (True, False, False, False, False, False):
                raise Refused("LOGIN de teste possui atributos inadequados")
            if c.execute("SELECT pg_has_role(session_user,'postgres','MEMBER')").fetchone()[0]:
                raise Refused("LOGIN de teste herda administrador")
        if not read_only:
            c.autocommit = False
        return c
    except Exception as error:
        if c is not None:
            c.close()
        raise Refused("conexao recusada; tipo=" + type(error).__name__,
                      evidence=error.evidence if isinstance(error, Refused) else None) from None


def inspect_session():
    """Duas sessoes reais; SELECTs e controle transacional, sem DDL/DML/SET ROLE."""
    snapshot_sql = (
        "SELECT pg_backend_pid(),session_user,current_user,current_database(),"
        "current_setting('transaction_read_only'),"
        "current_setting('statement_timeout')::interval = interval '20 seconds',"
        "current_setting('lock_timeout')::interval = interval '10 seconds',"
        "current_setting('transaction_isolation')"
    )
    snapshots = []
    with connect(read_only=True) as a, connect(read_only=True) as b:
        for c in (a, b):
            first = c.execute(snapshot_sql).fetchone()
            expected = (CONFIG["admin_user"], CONFIG["admin_user"], CONFIG["database"],
                        "on", True, True, "read committed")
            if first[1:] != expected:
                raise Refused("identidade, timeouts ou isolamento de sessao divergente")
            c.commit()
            begin_inspection(c)
            if c.execute(snapshot_sql).fetchone() != first:
                raise Refused("sessao nao preservada apos COMMIT")
            c.execute("SAVEPOINT rt_readonly_probe")
            c.execute("SELECT 1")
            c.execute("ROLLBACK TO SAVEPOINT rt_readonly_probe")
            c.execute("RELEASE SAVEPOINT rt_readonly_probe")
            c.rollback()
            begin_inspection(c)
            if c.execute(snapshot_sql).fetchone() != first:
                raise Refused("sessao nao preservada apos ROLLBACK")
            c.rollback()
            snapshots.append(first)
        if snapshots[0][0] == snapshots[1][0]:
            raise Refused("conexoes simultaneas sem backends independentes")
    return dict(session_user=CONFIG["admin_user"], current_user=CONFIG["admin_user"],
                database=CONFIG["database"], sslmode="verify-full", marker_matches=True,
                read_only=True, independent_backends=True, stable_sessions=True,
                timeouts_preserved=True, isolation="read committed", savepoint_rollback=True)


def inspect():
    if CONFIG.get("connection_mode", "direct") == "session_pooler":
        inspect_session()
    with connect(read_only=True) as c:
        if c.execute("SELECT version FROM rt_persist.schema_migrations").fetchall() != [(1,)]:
            raise Refused("versao de esquema inesperada")
        if c.execute("SHOW synchronous_commit").fetchone()[0] != "on":
            raise Refused("synchronous_commit precisa estar on")
        if c.execute("SHOW fsync").fetchone()[0] != "on":
            raise Refused("fsync precisa estar on")
        tables = c.execute(
            "SELECT oid FROM pg_class WHERE relnamespace='rt_persist'::regnamespace "
            "AND relkind IN ('r','p')").fetchall()
        funcs = c.execute(
            "SELECT oid FROM pg_proc WHERE pronamespace='rt_persist'::regnamespace").fetchall()
        # PUBLIC nao e LOGIN. Conferir ACL explicita/default em todos os objetos.
        public = c.execute("""
            SELECT count(*) FROM (
              SELECT a.grantee FROM pg_namespace n,
                LATERAL aclexplode(coalesce(n.nspacl,acldefault('n',n.nspowner))) a
                WHERE n.nspname='rt_persist'
              UNION ALL
              SELECT a.grantee FROM pg_class t,
                LATERAL aclexplode(coalesce(t.relacl,acldefault('r',t.relowner))) a
                WHERE t.relnamespace='rt_persist'::regnamespace AND t.relkind IN ('r','p')
              UNION ALL
              SELECT a.grantee FROM pg_proc f,
                LATERAL aclexplode(coalesce(f.proacl,acldefault('f',f.proowner))) a
                WHERE f.pronamespace='rt_persist'::regnamespace
            ) acl WHERE grantee=0
        """).fetchone()[0]
        if public:
            raise Refused("PUBLIC possui privilegios RT")
        for role in ("anon", "authenticated", "service_role"):
            exists = c.execute("SELECT 1 FROM pg_roles WHERE rolname=%s", (role,)).fetchone()
            if not exists:
                if not LOCAL:
                    raise Refused("papel Supabase esperado ausente")
                continue
            if c.execute("SELECT has_schema_privilege(%s,'rt_persist','USAGE') "
                         "OR has_schema_privilege(%s,'rt_persist','CREATE')", (role, role)).fetchone()[0]:
                raise Refused("papel Data API possui privilegios RT")
            for (oid,) in tables:
                if c.execute("SELECT has_table_privilege(%s,%s,'SELECT,INSERT,UPDATE,DELETE,"
                             "TRUNCATE,REFERENCES,TRIGGER') OR "
                             "has_any_column_privilege(%s,%s,'SELECT,INSERT,UPDATE,REFERENCES')",
                             (role, oid, role, oid)).fetchone()[0]:
                    raise Refused("papel Data API possui acesso a tabela RT")
            for (oid,) in funcs:
                if c.execute("SELECT has_function_privilege(%s,%s,'EXECUTE')", (role, oid)).fetchone()[0]:
                    raise Refused("papel Data API possui acesso a funcao RT")
        for role in ROLES:
            if c.execute("SELECT has_schema_privilege(%s,\'rt_persist\',\'CREATE\')", (role,)).fetchone()[0]:
                raise Refused("LOGIN restrito pode criar objetos no esquema")
            for (oid,) in tables:
                if c.execute("SELECT has_table_privilege(%s,%s,'SELECT,INSERT,UPDATE,DELETE,"
                             "TRUNCATE,REFERENCES,TRIGGER') OR "
                             "has_any_column_privilege(%s,%s,'SELECT,INSERT,UPDATE,REFERENCES')",
                             (role, oid, role, oid)).fetchone()[0]:
                    raise Refused("LOGIN restrito possui acesso direto a tabela")
    for role in ROLES:
        with connect(role, read_only=True) as c:
            if not c.execute("SELECT pg_has_role(session_user,'rt_persist_client','USAGE')").fetchone()[0]:
                raise Refused("LOGIN sem heranca do contrato")


class SupabaseTests(base.ContractTests):
    @classmethod
    def setUpClass(cls):
        inspect()
        # Sem migracao, CREATE ROLE, DELETE, DROP ou TRUNCATE.
        with connect() as c:
            c.execute("LOCK TABLE rt_persist.companies IN EXCLUSIVE MODE")
            for table in ("companies", "condominiums", "access_scopes", "documents",
                          "revisions", "save_requests", "outbox"):
                # Exige area RT vazia para impedir mistura com dados anteriores.
                exists = c.execute(sql.SQL("SELECT EXISTS(SELECT 1 FROM rt_persist.{})")
                                   .format(sql.Identifier(table))).fetchone()[0]
                if exists:
                    raise Refused("fixtures ou dados existentes; nenhuma limpeza executada")
            c.execute("INSERT INTO rt_persist.companies VALUES (%s,'EMPRESA A FICTICIA'),"
                      "(%s,'EMPRESA B FICTICIA')", (base.C1, base.C2))
            for company, condo in ((base.C1, base.D1), (base.C1, base.D2),
                                   (base.C2, base.D1), (base.C1, base.MARKER)):
                c.execute("INSERT INTO rt_persist.condominiums VALUES (%s,%s,'FICTICIO')",
                          (company, condo))
            for company, condo, role in ((base.C1, base.D1, "rt_test_a"),
                                         (base.C1, base.D2, "rt_test_a"),
                                         (base.C2, base.D1, "rt_test_b"),
                                         (base.C1, base.MARKER, "rt_test_a")):
                c.execute("INSERT INTO rt_persist.access_scopes VALUES (%s,%s,%s,%s)",
                          (role, company, condo, base.PERIOD))

    def test_new_process_without_local_files_recovers(self):
        with connect("rt_test_a") as c:
            base.save(c, 0, "RECUPERAVEL_SEM_DISCO_STREAMLIT",
                      base.uuid.UUID("60000000-0000-0000-0000-000000000006"),
                      condo=base.MARKER)
        command = [sys.executable, str(Path(__file__).resolve()),
                   "--config", CONFIG_PATH, "--recover"]
        if LOCAL:
            command.append("--local")
        with tempfile.TemporaryDirectory() as directory:
            result = subprocess.run(command, check=True, capture_output=True,
                                    text=True, cwd=directory, timeout=45,
                                    env=recovery_environment())
        self.assertIn("RECOVERY_OK", result.stdout)


def row_fingerprints(c):
    """md5 de cada linha RT existente; somente leitura, sem expor conteudo."""
    return {table: frozenset(r[0] for r in c.execute(
                sql.SQL("SELECT md5(row_to_json(t)::text) FROM rt_persist.{} t")
                .format(sql.Identifier(table))).fetchall())
            for table in RT_TABLES}


def selected_preflight(c):
    """Exige as fixtures da rodada anterior intactas; nunca insere nem apaga base."""
    companies = {r[0] for r in c.execute("SELECT company_id FROM rt_persist.companies").fetchall()}
    if companies != {base.C1, base.C2}:
        raise Refused("empresas ficticias da rodada anterior divergentes")
    condos = set(c.execute("SELECT company_id,condominium_id FROM rt_persist.condominiums").fetchall())
    if not {(base.C1, base.D1), (base.C1, base.D2), (base.C2, base.D1),
            (base.C1, base.MARKER)} <= condos:
        raise Refused("condominios ficticios da rodada anterior ausentes")
    scopes = set(c.execute("SELECT principal,company_id,condominium_id,period "
                           "FROM rt_persist.access_scopes").fetchall())
    if not {("rt_test_a", base.C1, base.D1, base.PERIOD), ("rt_test_a", base.C1, base.D2, base.PERIOD),
            ("rt_test_b", base.C2, base.D1, base.PERIOD),
            ("rt_test_a", base.C1, base.MARKER, base.PERIOD)} <= scopes:
        raise Refused("escopos ficticios da rodada anterior ausentes")
    if {s for s in scopes if s[0] != "rt_test_a"} != {("rt_test_b", base.C2, base.D1, base.PERIOD)}:
        raise Refused("escopos inesperados para rt_test_b ou rt_test_none")
    if c.execute("SELECT count(*) FROM rt_persist.revisions WHERE condominium_id=%s",
                 (base.MARKER,)).fetchone()[0] != 1:
        raise Refused("marcador de recuperacao da rodada anterior ausente")
    return row_fingerprints(c)


class SupabaseSelectedTests(base.ContractTests):
    """Somente SELECTED; metodos e assercoes herdados sem alteracao."""
    @classmethod
    def setUpClass(cls):
        inspect()
        with connect(read_only=True) as c:
            cls.preexisting = selected_preflight(c)
            c.rollback()

    @classmethod
    def tearDownClass(cls):
        with connect(read_only=True) as c:
            after = row_fingerprints(c)
            c.rollback()
        for table, rows in cls.preexisting.items():
            if not rows <= after[table]:
                raise Refused("dados da rodada anterior alterados: " + table)


def selected_suite():
    suite = unittest.TestSuite(SupabaseSelectedTests(name) for name in SELECTED)
    if suite.countTestCases() != 3 or len(set(SELECTED)) != 3:
        raise Refused("selecao de cenarios alterada")
    return suite


def recovery_environment():
    """O filho relera o mesmo manifesto e DPAPI; nenhum token via argv/ambiente."""
    return {key: value for key, value in os.environ.items()
            if not key.upper().startswith("PG") and key.upper() not in (
                "SUPABASE_ACCESS_TOKEN", "SUPABASE_MANAGEMENT_API_TOKEN", "RT_JIT_TOKEN")}


class SafeResult(unittest.TextTestResult):
    # unittest padrao inclui tracebacks, SQL e mensagens libpq. Somente tipo/SQLSTATE.
    def _exc_info_to_string(self, err, test):
        typ, value, _ = err
        return "detalhe omitido; tipo=" + typ.__name__ + "; SQLSTATE=" + str(
            getattr(value, "sqlstate", None))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--local", action="store_true")
    parser.add_argument("--inspect", action="store_true")
    parser.add_argument("--inspect-session", action="store_true")
    parser.add_argument("--recover", action="store_true")
    parser.add_argument("--run-fixtures", action="store_true")
    parser.add_argument("--run-selected", action="store_true")
    args = parser.parse_args()
    global JIT_MIN_INTERVAL
    try:
        if sum((args.inspect, args.inspect_session, args.recover, args.run_fixtures,
                args.run_selected)) != 1:
            raise Refused("selecionar uma unica operacao")
        load_config(args.config, args.local)
        base.connect = connect
        if args.inspect_session:
            print("SESSION_INSPECTION_OK " + json.dumps(inspect_session(), sort_keys=True))
        elif args.inspect:
            inspect()
            print("INSPECTION_OK")
        elif args.recover:
            base.verify_marker()
        elif args.run_selected:
            suite = selected_suite()
            JIT_MIN_INTERVAL = SELECTED_JIT_INTERVAL
            result = unittest.TextTestRunner(verbosity=2, resultclass=SafeResult).run(suite)
            return 0 if result.wasSuccessful() else 1
        else:
            suite = unittest.defaultTestLoader.loadTestsFromTestCase(SupabaseTests)
            if suite.countTestCases() != 15:
                raise Refused("quantidade de cenarios alterada")
            result = unittest.TextTestRunner(verbosity=2, resultclass=SafeResult).run(suite)
            return 0 if result.wasSuccessful() else 1
        return 0
    except Exception as error:
        if args.inspect_session and isinstance(error, Refused) and error.evidence:
            print("SESSION_INSPECTION_STOP " + json.dumps(error.evidence, sort_keys=True),
                  file=sys.stderr)
        print("EXECUCAO_RECUSADA: revisar configuracao, TLS, identidade, ACLs ou fixtures. "
              "Nenhuma mensagem bruta de conexao sera exibida.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
