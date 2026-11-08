"""PostgreSQL real, dados ficticios; recusa hosts/bancos nao locais de teste."""
import concurrent.futures
import datetime
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import unittest
import uuid

import psycopg
from psycopg import sql
from psycopg.conninfo import conninfo_to_dict
from psycopg.types.json import Jsonb

DSN = os.environ.get("RT_TEST_DSN", "")
C1 = uuid.UUID("10000000-0000-0000-0000-000000000001")
C2 = uuid.UUID("20000000-0000-0000-0000-000000000002")
D1 = uuid.UUID("30000000-0000-0000-0000-000000000003")
D2 = uuid.UUID("40000000-0000-0000-0000-000000000004")
PERIOD = datetime.date(2026, 10, 1)
MARKER = uuid.UUID("50000000-0000-0000-0000-000000000005")


def guard():
    cfg = conninfo_to_dict(DSN)
    if (cfg.get("dbname") != "rt_contract_test"
            or cfg.get("host") not in ("127.0.0.1", "localhost", "::1")
            or cfg.get("service") or cfg.get("hostaddr")):
        raise RuntimeError("Somente PostgreSQL loopback / rt_contract_test")


def connect(principal=None):
    guard()
    c = psycopg.connect(DSN, autocommit=True)
    if principal:
        c.execute(sql.SQL("SET SESSION AUTHORIZATION {}").format(sql.Identifier(principal)))
    c.autocommit = False
    return c


def save(c, base, text, operation=None, company=C1, condo=D1, period=PERIOD,
         restored=None, body=None):
    return c.execute(
        "SELECT rt_persist.save(%s,%s,%s,%s,%s,%s,%s)",
        (company, condo, period, operation or uuid.uuid4(), base,
         Jsonb(body if body is not None else {"campos": {"rel_diagnostico": text}}),
         restored),
    ).fetchone()[0]


def history(c, company=C1, condo=D1, period=PERIOD):
    return c.execute("SELECT rt_persist.history(%s,%s,%s)",
                     (company, condo, period)).fetchone()[0]


def verify_marker():
    with connect("rt_test_a") as c:
        rows = history(c, condo=MARKER)
        assert len(rows) == 1
        assert rows[0]["payload"]["campos"]["rel_diagnostico"] == "RECUPERAVEL_SEM_DISCO_STREAMLIT"
        n = c.execute("SELECT rt_persist.save(%s,%s,%s,%s,%s,%s,%s)",
                      (C1, MARKER, PERIOD,
                       uuid.UUID("60000000-0000-0000-0000-000000000006"),
                       0, Jsonb({"campos": {"rel_diagnostico":
                                            "RECUPERAVEL_SEM_DISCO_STREAMLIT"}}), None)
                      ).fetchone()[0]
        assert n["revision"] == 1
    print("RECOVERY_OK: revisao e idempotencia recuperadas do PostgreSQL")


class ContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Um banco dedicado vazio; nunca DROP de esquema preexistente.
        with connect() as c:
            if c.execute("SELECT to_regnamespace('rt_persist')").fetchone()[0]:
                raise RuntimeError("Banco nao vazio: use um novo banco efemero")
            for role in ("rt_test_a", "rt_test_b", "rt_test_none"):
                c.execute(sql.SQL("CREATE ROLE {} NOLOGIN").format(sql.Identifier(role)))
            c.execute(Path(__file__).with_name("001_contract.sql").read_text(encoding="utf-8"))
        with connect() as c:
            for role in ("rt_test_a", "rt_test_b", "rt_test_none"):
                c.execute(sql.SQL("GRANT rt_persist_client TO {}").format(sql.Identifier(role)))
            c.execute("INSERT INTO rt_persist.companies VALUES (%s,'EMPRESA A FICTICIA'),"
                      "(%s,'EMPRESA B FICTICIA')", (C1, C2))
            for company, condo in ((C1, D1), (C1, D2), (C2, D1), (C1, MARKER)):
                c.execute("INSERT INTO rt_persist.condominiums VALUES (%s,%s,'FICTICIO')",
                          (company, condo))
            for company, condo, principal in ((C1, D1, "rt_test_a"),
                                             (C1, D2, "rt_test_a"),
                                             (C2, D1, "rt_test_b"),
                                             (C1, MARKER, "rt_test_a")):
                c.execute("INSERT INTO rt_persist.access_scopes VALUES (%s,%s,%s,%s)",
                          (principal, company, condo, PERIOD))

    def setUp(self):
        self.condo = uuid.uuid4()
        with connect() as c:
            c.execute("INSERT INTO rt_persist.condominiums VALUES (%s,%s,'FICTICIO')",
                      (C1, self.condo))
            c.execute("INSERT INTO rt_persist.access_scopes VALUES (%s,%s,%s,%s)",
                      ("rt_test_a", C1, self.condo, PERIOD))

    def write(self, base, text, operation=None, **kw):
        with connect("rt_test_a") as c:
            return save(c, base, text, operation, condo=self.condo, **kw)

    def rows(self):
        with connect("rt_test_a") as c:
            return history(c, condo=self.condo)

    def race(self, jobs):
        barrier = threading.Barrier(len(jobs))
        def run(job):
            with connect("rt_test_a") as c:
                barrier.wait(timeout=15)
                return save(c, condo=self.condo, **job)
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(jobs)) as pool:
            futures = [pool.submit(run, j) for j in jobs]
            return [f.result(timeout=30) for f in futures]

    def test_concurrent_writers_preserve_loser_and_one_head(self):
        for base in range(8):
            results = self.race([{"base": base, "text": f"A-{base}"},
                                 {"base": base, "text": f"B-{base}"}])
            self.assertEqual(sorted(r["status"] for r in results), ["conflict", "saved"])
        self.assertEqual(len(self.rows()), 8)
        with connect("rt_test_a") as c:
            conflicts = c.execute("SELECT rt_persist.conflicts(%s,%s,%s)",
                                  (C1, self.condo, PERIOD)).fetchone()[0]
        self.assertEqual(len(conflicts), 8)
        texts = {r["payload"]["campos"]["rel_diagnostico"] for r in self.rows() + conflicts}
        self.assertEqual(texts, {f"{p}-{i}" for p in ("A", "B") for i in range(8)})

    def test_simultaneous_same_operation_is_one_revision(self):
        op = uuid.uuid4()
        results = self.race([{"base": 0, "text": "MESMA", "operation": op}] * 2)
        self.assertEqual(results[0], results[1])
        self.assertEqual(len(self.rows()), 1)

    def test_commit_response_discarded_retry_after_new_save(self):
        op = uuid.uuid4()
        # Commit efetivo; resultado descartado pelo cliente (resposta perdida).
        self.write(0, "PRIMEIRA", op)
        self.write(1, "SEGUNDA")
        replay = self.write(0, "PRIMEIRA", op)
        self.assertEqual(replay["revision"], 1)
        self.assertEqual(len(self.rows()), 2)

    def test_operation_reuse_different_payload_rejected(self):
        op = uuid.uuid4()
        self.write(0, "ORIGINAL", op)
        with self.assertRaises(psycopg.errors.InvalidParameterValue):
            self.write(0, "DIFERENTE", op)
        self.assertEqual(len(self.rows()), 1)

    def test_stale_retry_cannot_regress_head(self):
        self.write(0, "UM")
        conflict_op = uuid.uuid4()
        conflict = self.write(0, "CONFLITO", conflict_op)
        self.write(1, "DOIS")
        self.assertEqual(self.write(0, "CONFLITO", conflict_op), conflict)
        self.assertEqual(self.rows()[-1]["payload"]["campos"]["rel_diagnostico"], "DOIS")

    def test_restore_adds_revision_without_mutation(self):
        self.write(0, "ANTIGA")
        self.write(1, "ATUAL")
        self.write(2, "ANTIGA", restored=1)
        rows = self.rows()
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[-1]["restored_from"], 1)
        self.assertEqual(rows[1]["payload"]["campos"]["rel_diagnostico"], "ATUAL")

    def test_rollback_removes_revision_request_and_outbox(self):
        with connect("rt_test_a") as c:
            save(c, 0, "ABORTADA", condo=self.condo)
            c.rollback()
        self.assertEqual(self.rows(), [])
        with connect() as c:
            for table in ("save_requests", "outbox"):
                n = c.execute(sql.SQL("SELECT count(*) FROM rt_persist.{} "
                                     "WHERE condominium_id = %s").format(sql.Identifier(table)),
                              (self.condo,)).fetchone()[0]
                self.assertEqual(n, 0)

    def test_outbox_atomic_with_save(self):
        self.write(0, "CONFIRMADA")
        with connect() as c:
            n = c.execute("SELECT count(*) FROM rt_persist.outbox "
                          "WHERE condominium_id=%s AND revision=1 AND state='pending'",
                          (self.condo,)).fetchone()[0]
        self.assertEqual(n, 1)

    def test_cross_company_condo_period_denied(self):
        cases = [(C2, D1, PERIOD), (C1, uuid.uuid4(), PERIOD),
                 (C1, self.condo, datetime.date(2026, 11, 1))]
        for company, condo, period in cases:
            with self.subTest(company=company, condo=condo, period=period):
                with self.assertRaises(psycopg.errors.InsufficientPrivilege):
                    with connect("rt_test_a") as c:
                        save(c, 0, "NEGADA", company=company, condo=condo, period=period)
                with self.assertRaises(psycopg.errors.InsufficientPrivilege):
                    with connect("rt_test_a") as c:
                        history(c, company=company, condo=condo, period=period)

    def test_setting_cannot_forge_authorization(self):
        with self.assertRaises(psycopg.errors.InsufficientPrivilege):
            with connect("rt_test_none") as c:
                c.execute("SELECT set_config('app.company_id',%s,true)", (str(C1),))
                history(c, condo=D1)

    def test_same_condominium_id_in_two_companies_is_separate(self):
        with connect("rt_test_a") as c:
            save(c, 0, "A", condo=D1)
        with connect("rt_test_b") as c:
            save(c, 0, "B", company=C2, condo=D1)
            self.assertEqual(history(c, company=C2)[0]["payload"]["campos"]["rel_diagnostico"], "B")

    def test_direct_tables_and_internal_function_denied(self):
        for statement in ("SELECT * FROM rt_persist.revisions",
                          "DELETE FROM rt_persist.documents",
                          "TRUNCATE rt_persist.revisions",
                          "SELECT rt_persist.assert_scope(NULL,NULL,NULL)"):
            with self.subTest(statement=statement):
                with self.assertRaises(psycopg.errors.InsufficientPrivilege):
                    with connect("rt_test_a") as c:
                        c.execute(statement)

    def test_immutability_trigger(self):
        self.write(0, "IMUTAVEL")
        for table in ("revisions", "save_requests"):
            with self.assertRaises(psycopg.errors.ObjectNotInPrerequisiteState):
                with connect() as c:
                    c.execute(sql.SQL("DELETE FROM rt_persist.{} WHERE condominium_id=%s")
                              .format(sql.Identifier(table)), (self.condo,))

    def test_invalid_identity_payload_and_future_base(self):
        for body, base in (({"campos": {}, "identity": {}}, 0),
                           ({"campos": []}, 0), ({"campos": {}}, 20)):
            with self.assertRaises(psycopg.errors.InvalidParameterValue):
                self.write(base, "INVALIDA", body=body)
        self.assertEqual(self.rows(), [])

    def test_new_process_without_local_files_recovers(self):
        with connect("rt_test_a") as c:
            save(c, 0, "RECUPERAVEL_SEM_DISCO_STREAMLIT",
                 uuid.UUID("60000000-0000-0000-0000-000000000006"), condo=MARKER)
        result = subprocess.run([sys.executable, str(Path(__file__).resolve()),
                                 "--verify-restart"],
                                check=True, capture_output=True, text=True, cwd="/tmp",
                                timeout=30)
        self.assertIn("RECOVERY_OK", result.stdout)


if __name__ == "__main__":
    guard()
    if sys.argv[1:] == ["--verify-restart"]:
        verify_marker()
    else:
        unittest.main(verbosity=2)
