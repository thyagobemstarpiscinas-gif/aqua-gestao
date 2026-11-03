"""Validacao offline e TLS negativo no container local, sem endpoint Supabase."""
import copy
import os
from pathlib import Path
import unittest
from unittest.mock import patch

import psycopg
import supabase_runner as runner


class GuardTests(unittest.TestCase):
    def setUp(self):
        self.cfg = dict(environment="homologacao", project_name="acqua-gestao-homologacao",
                        project_ref="a" * 20, host="db." + "a" * 20 + ".supabase.co",
                        port=5432, database="postgres", marker="1" * 32,
                        admin_user="postgres", ca_file="/tmp/ca", passfile="/tmp/pgpass")

    def test_exact_remote_allowlist(self):
        self.assertEqual(runner.validate(self.cfg), self.cfg)

    def test_production_name_refused(self):
        self.cfg["environment"] = "producao"
        with self.assertRaises(runner.Refused):
            runner.validate(self.cfg)

    def test_host_database_and_port_refused(self):
        for key, value in (("host", "db.outro.supabase.co"),
                           ("database", "outro"), ("port", 6543)):
            cfg = dict(self.cfg, **{key: value})
            with self.subTest(key=key), self.assertRaises(runner.Refused):
                runner.validate(cfg)

    def test_unknown_connection_options_refused(self):
        for key in ("password", "sslmode", "service", "hostaddr", "options"):
            cfg = dict(self.cfg, **{key: "invalid"})
            with self.subTest(key=key), self.assertRaises(runner.Refused):
                runner.validate(cfg)

    def test_marker_required(self):
        self.cfg["marker"] = ""
        with self.assertRaises(runner.Refused):
            runner.validate(self.cfg)

    def test_local_flag_cannot_enable_remote(self):
        with self.assertRaises(runner.Refused):
            runner.validate(self.cfg, local=True)

    def test_local_target_cannot_use_remote_mode(self):
        self.cfg.update(project_ref="local-ci", host="localhost", database="rt_contract_test")
        runner.validate(self.cfg, local=True)
        with self.assertRaises(runner.Refused):
            runner.validate(self.cfg)

    def test_error_output_does_not_include_connection_exception(self):
        runner.CONFIG = dict(self.cfg)
        with patch.object(runner.psycopg, "connect", side_effect=RuntimeError("SECRET_SENTINEL")):
            with self.assertRaises(runner.Refused) as caught:
                runner.connect()
        self.assertNotIn("SECRET_SENTINEL", str(caught.exception))


@unittest.skipUnless(os.environ.get("RT_LOCAL_TLS_TEST") == "1", "local TLS fixture only")
class TLSNegativeTests(unittest.TestCase):
    def test_wrong_hostname_fails_verification(self):
        # Certificado do CI possui apenas DNS:localhost; endereco IP nao coincide.
        with self.assertRaises(psycopg.OperationalError):
            psycopg.connect(host="127.0.0.1", port=5432, dbname="rt_contract_test",
                            user="postgres", password="ephemeral-test-only",
                            sslmode="verify-full",
                            sslrootcert=str(Path("work/rt-supabase-ci/server.crt").resolve()),
                            connect_timeout=5)

    def test_wrong_ca_fails_verification(self):
        with self.assertRaises(psycopg.OperationalError):
            psycopg.connect(host="localhost", port=5432, dbname="rt_contract_test",
                            user="postgres", password="ephemeral-test-only",
                            sslmode="verify-full",
                            sslrootcert=str(Path("work/rt-supabase-ci/untrusted.crt").resolve()),
                            connect_timeout=5)


if __name__ == "__main__":
    unittest.main(verbosity=2)
