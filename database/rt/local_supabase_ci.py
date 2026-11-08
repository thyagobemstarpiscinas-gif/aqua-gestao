"""Bootstrap EXCLUSIVO do PostgreSQL efemero local do CI; nunca aceita URL remota."""
import json
from pathlib import Path
import sys

import psycopg
from psycopg import sql

root = Path("work/rt-supabase-ci").resolve()
root.mkdir(parents=True, exist_ok=True)
ca = root / "server.crt"
marker = "1" * 32
# Credenciais ficticias apenas do container descartavel, sem variaveis de destino.
with psycopg.connect(host="localhost", port=5432, dbname="rt_contract_test",
                     user="postgres", password="ephemeral-test-only",
                     sslmode="verify-full", sslrootcert=str(ca),
                     autocommit=True) as c:
    if c.execute("SELECT current_database()").fetchone()[0] != "rt_contract_test":
        raise RuntimeError("local only")
    if c.execute("SELECT to_regnamespace('rt_persist')").fetchone()[0]:
        raise RuntimeError("not empty")
    c.execute(Path("database/rt/001_contract.sql").read_text(encoding="utf-8"))
    c.execute(sql.SQL("COMMENT ON DATABASE rt_contract_test IS {}").format(
        sql.Literal("acqua-gestao-homologacao:" + marker)))
    for role in ("rt_test_a", "rt_test_b", "rt_test_none"):
        c.execute(sql.SQL("CREATE ROLE {} LOGIN INHERIT NOSUPERUSER NOCREATEDB "
                         "NOCREATEROLE NOREPLICATION NOBYPASSRLS PASSWORD {}").format(
                             sql.Identifier(role), sql.Literal("fixture-" + role + "-only")))
        c.execute(sql.SQL("GRANT rt_persist_client TO {}").format(sql.Identifier(role)))
    for role in ("anon", "authenticated", "service_role"):
        c.execute(sql.SQL("CREATE ROLE {} NOLOGIN").format(sql.Identifier(role)))
passfile = root / "pgpass"
passfile.write_text(
    "localhost:5432:rt_contract_test:postgres:ephemeral-test-only\n"
    + "".join("localhost:5432:rt_contract_test:" + role + ":fixture-" + role + "-only\n"
              for role in ("rt_test_a", "rt_test_b", "rt_test_none")),
    encoding="utf-8")
passfile.chmod(0o600)
config = dict(environment="homologacao", project_name="acqua-gestao-homologacao",
              project_ref="local-ci", host="localhost", port=5432,
              database="rt_contract_test", marker=marker, admin_user="postgres",
              ca_file=str(ca), passfile=str(passfile))
manifest = root / "authorization.toml"
manifest.write_text("".join(k + " = " + json.dumps(v) + "\n" for k, v in config.items()),
                    encoding="utf-8")
manifest.chmod(0o600)
print("LOCAL_BOOTSTRAP_OK")
