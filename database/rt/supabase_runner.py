"""Executor explicito: esquema ja instalado, LOGINs reais, TLS e fixtures novas."""
import argparse
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import tempfile
import tomllib
import unittest

import psycopg
from psycopg import sql
import test_contract as base

CONFIG = None
CONFIG_PATH = None
LOCAL = False
ROLES = ("rt_test_a", "rt_test_b", "rt_test_none")


class Refused(RuntimeError):
    pass


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
    if set(cfg) != expected:
        raise Refused("campos de autorizacao ausentes ou desconhecidos")
    if cfg["environment"] != "homologacao" or cfg["project_name"] != "acqua-gestao-homologacao":
        raise Refused("homologacao nao autorizada")
    if not isinstance(cfg["port"], int) or isinstance(cfg["port"], bool) or not 1 <= cfg["port"] <= 65535:
        raise Refused("porta invalida")
    if not isinstance(cfg["marker"], str) or not re.fullmatch(r"[a-f0-9]{32}", cfg["marker"]):
        raise Refused("marcador de homologacao invalido")
    if local:
        if (cfg["project_ref"], cfg["host"], cfg["database"]) != (
                "local-ci", "localhost", "rt_contract_test"):
            raise Refused("destino local invalido")
    else:
        if not isinstance(cfg["project_ref"], str) or not re.fullmatch(r"[a-z0-9]{20}", cfg["project_ref"]):
            raise Refused("project ref invalido")
        if cfg["host"] != "db." + cfg["project_ref"] + ".supabase.co" or cfg["database"] != "postgres":
            raise Refused("endpoint direto fora da autorizacao")
        if cfg["port"] != 5432:
            raise Refused("porta Supabase direta invalida")
    if cfg["admin_user"] != "postgres":
        raise Refused("administrador de homologacao inesperado")
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


def connect(principal=None):
    if CONFIG is None:
        raise Refused("configuracao nao carregada")
    user = principal or CONFIG["admin_user"]
    if user not in (*ROLES, CONFIG["admin_user"]):
        raise Refused("LOGIN fora da lista autorizada")
    c = None
    try:
        c = psycopg.connect(
            host=CONFIG["host"], port=CONFIG["port"], dbname=CONFIG["database"],
            user=user, passfile=CONFIG["passfile"], sslmode="verify-full",
            sslrootcert=CONFIG["ca_file"], connect_timeout=10,
            application_name="acqua-rt-homolog-contract",
            options="-c statement_timeout=20000 -c lock_timeout=10000",
            prepare_threshold=None, autocommit=True)
        identity = c.execute(
            "SELECT session_user,current_user,current_database(),"
            "shobj_description(oid,'pg_database') "
            "FROM pg_database WHERE datname=current_database()"
        ).fetchone()
        marker = "acqua-gestao-homologacao:" + CONFIG["marker"]
        if identity != (user, user, CONFIG["database"], marker):
            raise Refused("identidade ou marcador de destino divergente")
        tls = c.execute("SELECT ssl FROM pg_stat_ssl WHERE pid=pg_backend_pid()").fetchone()
        if not tls or tls[0] is not True or not c.info.ssl_in_use:
            raise Refused("TLS nao confirmado")
        if c.info.get_parameters().get("sslmode") != "verify-full":
            raise Refused("verify-full nao confirmado")
        if principal:
            flags = c.execute(
                "SELECT rolcanlogin,rolsuper,rolcreatedb,rolcreaterole,rolreplication,rolbypassrls "
                "FROM pg_roles WHERE rolname=session_user").fetchone()
            if flags != (True, False, False, False, False, False):
                raise Refused("LOGIN de teste possui atributos inadequados")
            if c.execute("SELECT pg_has_role(session_user,'postgres','MEMBER')").fetchone()[0]:
                raise Refused("LOGIN de teste herda administrador")
        c.autocommit = False
        return c
    except Exception:
        if c is not None:
            c.close()
        raise Refused("conexao recusada; revisar destino, identidade e TLS sem publicar diagnostico bruto") from None


def inspect():
    with connect() as c:
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
            for (oid,) in tables:
                if c.execute("SELECT has_table_privilege(%s,%s,'SELECT,INSERT,UPDATE,DELETE,"
                             "TRUNCATE,REFERENCES,TRIGGER') OR "
                             "has_any_column_privilege(%s,%s,'SELECT,INSERT,UPDATE,REFERENCES')",
                             (role, oid, role, oid)).fetchone()[0]:
                    raise Refused("LOGIN restrito possui acesso direto a tabela")
    for role in ROLES:
        with connect(role) as c:
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
                                    text=True, cwd=directory, timeout=45)
        self.assertIn("RECOVERY_OK", result.stdout)


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
    parser.add_argument("--recover", action="store_true")
    parser.add_argument("--run-fixtures", action="store_true")
    args = parser.parse_args()
    try:
        if sum((args.inspect, args.recover, args.run_fixtures)) != 1:
            raise Refused("selecionar uma unica operacao")
        load_config(args.config, args.local)
        base.connect = connect
        if args.inspect:
            inspect()
            print("INSPECTION_OK")
        elif args.recover:
            base.verify_marker()
        else:
            suite = unittest.defaultTestLoader.loadTestsFromTestCase(SupabaseTests)
            if suite.countTestCases() != 15:
                raise Refused("quantidade de cenarios alterada")
            result = unittest.TextTestRunner(verbosity=2, resultclass=SafeResult).run(suite)
            return 0 if result.wasSuccessful() else 1
        return 0
    except Exception:
        print("EXECUCAO_RECUSADA: revisar configuracao, TLS, identidade, ACLs ou fixtures. "
              "Nenhuma mensagem bruta de conexao sera exibida.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
