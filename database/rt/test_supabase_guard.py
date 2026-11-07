"""Validacao offline e TLS negativo no container local, sem endpoint Supabase."""
import copy
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import psycopg
import supabase_runner as runner


class GuardTests(unittest.TestCase):
    def setUp(self):
        temporary = Path(tempfile.gettempdir()).resolve()
        self.cfg = dict(environment="homologacao", project_name="acqua-gestao-homologacao",
                        project_ref="a" * 20, host="db." + "a" * 20 + ".supabase.co",
                        port=5432, database="postgres", marker="1" * 32,
                        admin_user="postgres", ca_file=str(temporary / "ca"),
                        passfile=str(temporary / "pgpass"))

    def test_exact_remote_allowlist(self):
        self.assertEqual(runner.validate(self.cfg), self.cfg)

    def test_session_pooler_requires_explicit_mode_and_sao_paulo(self):
        cfg = dict(self.cfg, connection_mode="session_pooler",
                   host="aws-1-sa-east-1.pooler.supabase.com")
        self.assertEqual(runner.validate(cfg), cfg)
        for changes in ({"connection_mode": "direct"}, {"connection_mode": "transaction"},
                        {"host": "aws-1-us-east-1.pooler.supabase.com"},
                        {"host": "aws-1-sa-east-1.pooler.supabase.com.attacker.example"},
                        {"host": "aws-1-sa-east-1.pooler.supabase.com/path"},
                        {"port": 6543}, {"admin_user": "postgres." + "a" * 20},
                        {"project_ref": "invalid"}):
            with self.subTest(changes=changes), self.assertRaises(runner.Refused):
                runner.validate(dict(cfg, **changes))
        cfg.pop("connection_mode")
        with self.assertRaises(runner.Refused):
            runner.validate(cfg)

    def test_local_mode_cannot_use_pooler(self):
        cfg = dict(self.cfg, connection_mode="session_pooler", project_ref="local-ci",
                   host="localhost", database="rt_contract_test")
        with self.assertRaises(runner.Refused):
            runner.validate(cfg, local=True)

    def test_pooler_wire_user_cannot_replace_database_identity(self):
        cfg = dict(self.cfg, connection_mode="session_pooler",
                   host="aws-1-sa-east-1.pooler.supabase.com")
        for logical_user in ("postgres", *runner.ROLES):
            wire_user = logical_user + "." + cfg["project_ref"]
            for sql_user in (logical_user, wire_user, "unexpected_role"):
                connection = MagicMock()
                connection.pgconn.ssl_in_use = True
                connection.info.get_parameters.return_value = {"sslmode": "verify-full"}
                connection.info.transaction_status = runner.TransactionStatus.IDLE
                connection.execute.return_value.fetchone.side_effect = [
                    (sql_user, sql_user, "postgres", "acqua-gestao-homologacao:" + cfg["marker"]),
                    (True,), (True, True, "on", "on"),
                    (True, False, False, False, False, False), (False,)]
                with self.subTest(login=logical_user, identity=sql_user), \
                     patch.object(runner, "CONFIG", cfg), \
                     patch.object(runner.psycopg, "connect", return_value=connection) as dial:
                    principal = None if logical_user == "postgres" else logical_user
                    if sql_user == logical_user:
                        self.assertIs(runner.connect(principal, read_only=True), connection)
                    else:
                        with self.assertRaises(runner.Refused):
                            runner.connect(principal, read_only=True)
                        connection.close.assert_called_once()
                    self.assertEqual(dial.call_args.kwargs["user"], wire_user)
                    self.assertEqual(dial.call_args.kwargs["sslmode"], "verify-full")
                    self.assertEqual(dial.call_args.kwargs["sslrootcert"], cfg["ca_file"])
                    self.assertEqual(dial.call_args.kwargs["options"], "")

    def test_pooler_rejects_wrong_marker_and_missing_backend_tls(self):
        cfg = dict(self.cfg, connection_mode="session_pooler",
                   host="aws-1-sa-east-1.pooler.supabase.com")
        for marker, tls in (("acqua-gestao-homologacao:" + "2" * 32, True),
                            ("acqua-gestao-homologacao:" + cfg["marker"], False)):
            connection = MagicMock()
            connection.pgconn.ssl_in_use = True
            connection.info.get_parameters.return_value = {"sslmode": "verify-full"}
            connection.execute.return_value.fetchone.side_effect = [
                ("postgres", "postgres", "postgres", marker), (tls,)]
            with patch.object(runner, "CONFIG", cfg), \
                 patch.object(runner.psycopg, "connect", return_value=connection):
                with self.assertRaises(runner.Refused) as caught:
                    runner.connect()
                connection.close.assert_called_once()
                if tls is False:
                    self.assertEqual(caught.exception.evidence, dict(
                        identity_matches=True, marker_matches=True, client_tls=True,
                        client_verify_full=True, backend_ssl_row_present=True, backend_tls=False))

    def test_client_tls_cannot_be_replaced_by_backend_tls(self):
        for client_tls, sslmode in ((False, "verify-full"), (True, "require")):
            connection = MagicMock()
            connection.pgconn.ssl_in_use = client_tls
            connection.info.get_parameters.return_value = {"sslmode": sslmode}
            connection.execute.return_value.fetchone.side_effect = [
                ("postgres", "postgres", "postgres", "acqua-gestao-homologacao:" + self.cfg["marker"]),
                (True,)]
            with self.subTest(tls=client_tls, mode=sslmode), \
                 patch.object(runner, "CONFIG", self.cfg), \
                 patch.object(runner.psycopg, "connect", return_value=connection):
                with self.assertRaises(runner.Refused) as caught:
                    runner.connect()
                self.assertTrue(caught.exception.evidence["backend_tls"])
                self.assertTrue(all(isinstance(value, bool)
                                    for value in caught.exception.evidence.values()))
                connection.close.assert_called_once()

    def test_session_probe_refuses_backend_reuse_state_loss_and_write_mode(self):
        first = (123, "postgres", "postgres", "postgres", "on", True, True, "read committed")
        second = (456, *first[1:])
        for fault in (None, "commit_switch", "rollback_switch", "shared_backend",
                      "timeout_ignored", "writable", "identity_changed"):
            a = MagicMock()
            b = MagicMock()
            a.__enter__.return_value = a
            b.__enter__.return_value = b
            a.info.transaction_status = runner.TransactionStatus.IDLE
            b.info.transaction_status = runner.TransactionStatus.IDLE
            a_rows = [first, first, first]
            b_rows = [second, second, second]
            if fault == "commit_switch":
                a_rows[1] = second
            elif fault == "rollback_switch":
                a_rows[2] = second
            elif fault == "shared_backend":
                b_rows = list(a_rows)
            elif fault == "timeout_ignored":
                a_rows[0] = (*first[:5], False, *first[6:])
            elif fault == "writable":
                a_rows[0] = (*first[:4], "off", *first[5:])
            elif fault == "identity_changed":
                a_rows[0] = (123, "postgres", "unexpected_role", *first[3:])
            effective = (True, True, "on", "on")
            a.execute.return_value.fetchone.side_effect = [a_rows[0], effective, a_rows[1], effective, a_rows[2]]
            b.execute.return_value.fetchone.side_effect = [b_rows[0], effective, b_rows[1], effective, b_rows[2]]
            with self.subTest(fault=fault), patch.object(runner, "CONFIG", self.cfg), \
                 patch.object(runner, "connect", side_effect=[a, b]) as connect:
                if fault is None:
                    result = runner.inspect_session()
                    self.assertTrue(result["stable_sessions"])
                    self.assertNotIn(self.cfg["marker"], str(result))
                    self.assertNotIn(self.cfg["passfile"], str(result))
                else:
                    with self.assertRaises(runner.Refused):
                        runner.inspect_session()
                self.assertTrue(all(call.kwargs == {"read_only": True}
                                    for call in connect.call_args_list))

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


class PoolerSessionRegressionTests(unittest.TestCase):
    def setUp(self):
        temporary = Path(tempfile.gettempdir()).resolve()
        self.cfg = dict(environment="homologacao", connection_mode="session_pooler",
                        project_name="acqua-gestao-homologacao", project_ref="a" * 20,
                        host="aws-1-sa-east-1.pooler.supabase.com", port=5432,
                        database="postgres", marker="1" * 32, admin_user="postgres",
                        ca_file=str(temporary / "ca"), passfile=str(temporary / "pgpass"))

    def backend(self, *, ignored=(), begin_read_only=True, bad_identity=False, backend_tls=True):
        """Backend que ignora startup como o pooler observado; SQL pode aplicar ou falhar."""
        c = MagicMock()
        c.autocommit = True
        c.info.transaction_status = runner.TransactionStatus.IDLE
        c.info.get_parameters.return_value = {"sslmode": "verify-full"}
        c.pgconn.ssl_in_use = True
        state = dict(statement_timeout="2min", lock_timeout="0", default_read_only="off",
                     transaction_read_only="off")

        def execute(statement, *args):
            result = MagicMock()
            if statement.startswith("SELECT session_user,current_user,current_database()"):
                user = "wrong_role" if bad_identity else "postgres"
                result.fetchone.return_value = (user, user, "postgres",
                                                "acqua-gestao-homologacao:" + self.cfg["marker"])
            elif statement.startswith("SELECT ssl FROM pg_stat_ssl"):
                result.fetchone.return_value = (backend_tls,)
            elif statement == "SET SESSION statement_timeout = '20s'":
                if "statement_timeout" not in ignored:
                    state["statement_timeout"] = "20s"
            elif statement == "SET SESSION lock_timeout = '10s'":
                if "lock_timeout" not in ignored:
                    state["lock_timeout"] = "10s"
            elif statement == "SET SESSION default_transaction_read_only = on":
                if "default_read_only" not in ignored:
                    state["default_read_only"] = "on"
            elif statement == "BEGIN READ ONLY":
                c.info.transaction_status = runner.TransactionStatus.INTRANS
                state["transaction_read_only"] = "on" if begin_read_only else "off"
            elif statement.startswith("SELECT current_setting('statement_timeout')"):
                result.fetchone.return_value = (state["statement_timeout"] == "20s",
                    state["lock_timeout"] == "10s", state["transaction_read_only"],
                    state["default_read_only"])
            else:
                raise AssertionError("Consulta inesperada no teste")
            return result

        c.execute.side_effect = execute
        return c, state

    def test_ignored_startup_is_replaced_by_verified_sql(self):
        c, state = self.backend()
        with patch.object(runner, "CONFIG", self.cfg), \
             patch.object(runner.psycopg, "connect", return_value=c) as dial:
            self.assertIs(runner.connect(read_only=True), c)
        self.assertEqual(dial.call_args.kwargs["options"], "")
        self.assertEqual(state, dict(statement_timeout="20s", lock_timeout="10s",
                                    default_read_only="on", transaction_read_only="on"))
        statements = [call.args[0] for call in c.execute.call_args_list]
        self.assertLess(statements.index("SELECT ssl FROM pg_stat_ssl WHERE pid=pg_backend_pid()"),
                        statements.index("SET SESSION statement_timeout = '20s'"))
        self.assertIn("BEGIN READ ONLY", statements)
        c.close.assert_not_called()

    def test_no_session_configuration_before_identity_and_tls_guards(self):
        for changes in ({"bad_identity": True}, {"backend_tls": False}):
            c, _ = self.backend(**changes)
            with self.subTest(changes=changes), patch.object(runner, "CONFIG", self.cfg), \
                 patch.object(runner.psycopg, "connect", return_value=c):
                with self.assertRaises(runner.Refused):
                    runner.connect(read_only=True)
                self.assertFalse(any(call.args[0].startswith("SET ")
                                     for call in c.execute.call_args_list))
                c.close.assert_called_once()

    def test_sql_settings_must_really_apply(self):
        for ignored in ("statement_timeout", "lock_timeout", "default_read_only"):
            c, _ = self.backend(ignored=(ignored,))
            with self.subTest(ignored=ignored), patch.object(runner, "CONFIG", self.cfg), \
                 patch.object(runner.psycopg, "connect", return_value=c):
                with self.assertRaises(runner.Refused):
                    runner.connect(read_only=True)
                c.close.assert_called_once()

    def test_begin_read_only_must_really_apply(self):
        c, _ = self.backend(begin_read_only=False)
        with patch.object(runner, "CONFIG", self.cfg), \
             patch.object(runner.psycopg, "connect", return_value=c):
            with self.assertRaises(runner.Refused):
                runner.connect(read_only=True)
        c.close.assert_called_once()

    def test_non_inspection_does_not_enable_read_only(self):
        c, state = self.backend()
        with patch.object(runner, "CONFIG", self.cfg), \
             patch.object(runner.psycopg, "connect", return_value=c):
            runner.connect()
        self.assertEqual(state["default_read_only"], "off")
        self.assertFalse(c.autocommit)
        self.assertFalse(any(call.args[0] == "BEGIN READ ONLY"
                             for call in c.execute.call_args_list))

    def test_new_inspection_transaction_rechecks_settings_before_queries(self):
        for effective in ((True, True, "on", "on"), (False, True, "on", "on"),
                          (True, False, "on", "on"), (True, True, "off", "on")):
            c = MagicMock()
            c.info.transaction_status = runner.TransactionStatus.IDLE
            c.execute.return_value.fetchone.return_value = effective
            with self.subTest(effective=effective):
                if effective == (True, True, "on", "on"):
                    runner.begin_inspection(c)
                else:
                    with self.assertRaises(runner.Refused):
                        runner.begin_inspection(c)
                self.assertEqual(c.execute.call_args_list[0].args, ("BEGIN READ ONLY",))
                self.assertEqual(c.execute.call_count, 2)

    def test_nested_begin_is_refused_instead_of_silently_ignored(self):
        c = MagicMock()
        c.info.transaction_status = runner.TransactionStatus.INTRANS
        with self.assertRaises(runner.Refused):
            runner.begin_inspection(c)
        c.execute.assert_not_called()

    def test_configuration_error_is_sanitized(self):
        c, _ = self.backend()
        execute = c.execute.side_effect

        def fail_on_set(statement, *args):
            if statement.startswith("SET "):
                raise psycopg.ProgrammingError("SECRET_SENTINEL")
            return execute(statement, *args)

        c.execute.side_effect = fail_on_set
        with patch.object(runner, "CONFIG", self.cfg), \
             patch.object(runner.psycopg, "connect", return_value=c):
            with self.assertRaises(runner.Refused) as caught:
                runner.connect(read_only=True)
        self.assertNotIn("SECRET_SENTINEL", str(caught.exception))
        c.close.assert_called_once()


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
