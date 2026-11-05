"""Regressoes offline do modo --run-selected: sem rede, sem banco, sem token real."""
import contextlib
import io
import sys
import unittest
import uuid
from unittest.mock import MagicMock, patch

import supabase_runner as runner
import test_contract as base

EXPECTED = ("test_setting_cannot_forge_authorization",
            "test_simultaneous_same_operation_is_one_revision",
            "test_stale_retry_cannot_regress_head")
CONDOS = [(base.C1, base.D1), (base.C1, base.D2), (base.C2, base.D1), (base.C1, base.MARKER)]
SCOPES = [("rt_test_a", base.C1, base.D1, base.PERIOD), ("rt_test_a", base.C1, base.D2, base.PERIOD),
          ("rt_test_b", base.C2, base.D1, base.PERIOD), ("rt_test_a", base.C1, base.MARKER, base.PERIOD)]
WRITES = ("INSERT", "UPDATE", "DELETE", "TRUNCATE", "DROP", "ALTER", "CREATE", "GRANT", "REVOKE")


def fake_db(companies=None, condos=None, scopes=None, marker=1, rows=None):
    """Conexao simulada que responde apenas aos SELECTs da pre-verificacao."""
    rows = rows if rows is not None else {t: {"h-" + t} for t in runner.RT_TABLES}
    c = MagicMock()
    def execute(query, params=None):
        text = query if isinstance(query, str) else query.as_string(None)
        result = MagicMock()
        if "FROM rt_persist.companies" in text and "md5" not in text:
            result.fetchall.return_value = [(x,) for x in (companies or [base.C1, base.C2])]
        elif "FROM rt_persist.condominiums" in text and "md5" not in text:
            result.fetchall.return_value = condos if condos is not None else CONDOS
        elif "FROM rt_persist.access_scopes" in text and "md5" not in text:
            result.fetchall.return_value = scopes if scopes is not None else SCOPES
        elif "rt_persist.revisions WHERE" in text:
            result.fetchone.return_value = (marker,)
        elif "md5" in text:
            table = next(t for t in runner.RT_TABLES if '"' + t + '"' in text)
            result.fetchall.return_value = [(h,) for h in rows[table]]
        else:
            raise AssertionError("consulta inesperada")
        return result
    c.execute.side_effect = execute
    c.__enter__.return_value = c
    return c


def executed(c):
    return [(q if isinstance(q, str) else q.as_string(None)).upper()
            for q in (call.args[0] for call in c.execute.call_args_list)]


class SelectedTests(unittest.TestCase):
    def test_exactly_three_named_scenarios_from_the_failed_run(self):
        self.assertEqual(runner.SELECTED, EXPECTED)
        suite = runner.selected_suite()
        self.assertEqual(suite.countTestCases(), 3)
        self.assertEqual(tuple(t._testMethodName for t in suite), EXPECTED)
        self.assertNotIn("test_new_process_without_local_files_recovers",
                         [t._testMethodName for t in suite])

    def test_assertions_are_inherited_unchanged(self):
        for name in EXPECTED:
            self.assertIs(getattr(runner.SupabaseSelectedTests, name),
                          getattr(base.ContractTests, name))
        self.assertIs(runner.SupabaseSelectedTests.setUp, base.ContractTests.setUp)

    def test_full_suite_still_has_fifteen_scenarios(self):
        self.assertEqual(len(unittest.defaultTestLoader.getTestCaseNames(runner.SupabaseTests)), 15)

    def test_preflight_is_read_only_and_returns_fingerprints(self):
        c = fake_db()
        snapshot = runner.selected_preflight(c)
        self.assertEqual(set(snapshot), set(runner.RT_TABLES))
        self.assertFalse(any(q.lstrip().startswith(WRITES) for q in executed(c)))

    def test_preflight_refuses_missing_or_unexpected_prior_fixtures(self):
        cases = [dict(companies=[base.C1]),
                 dict(companies=[base.C1, base.C2, uuid.uuid4()]),
                 dict(condos=CONDOS[:3]),
                 dict(scopes=SCOPES[:3]),
                 dict(scopes=SCOPES + [("rt_test_none", base.C1, base.D1, base.PERIOD)]),
                 dict(scopes=SCOPES + [("rt_test_b", base.C1, base.D1, base.PERIOD)]),
                 dict(marker=0), dict(marker=2)]
        for changes in cases:
            with self.subTest(changes=list(changes)), self.assertRaises(runner.Refused):
                runner.selected_preflight(fake_db(**changes))

    def test_preflight_accepts_extra_per_test_condos_of_rt_test_a(self):
        extra = uuid.uuid4()
        runner.selected_preflight(fake_db(condos=CONDOS + [(base.C1, extra)],
                                          scopes=SCOPES + [("rt_test_a", base.C1, extra, base.PERIOD)]))

    def test_setup_class_inspects_and_never_inserts_base_fixtures(self):
        c = fake_db()
        with patch.object(runner, "inspect") as inspect, patch.object(runner, "connect", return_value=c) as dial:
            runner.SupabaseSelectedTests.setUpClass()
        inspect.assert_called_once()
        dial.assert_called_once_with(read_only=True)
        self.assertFalse(any(q.lstrip().startswith(WRITES) for q in executed(c)))

    def test_teardown_refuses_if_prior_rows_changed_or_disappeared(self):
        before = {t: frozenset({"h-" + t}) for t in runner.RT_TABLES}
        with patch.object(runner.SupabaseSelectedTests, "preexisting", before, create=True):
            grown = {t: {"h-" + t, "novo"} for t in runner.RT_TABLES}
            with patch.object(runner, "connect", return_value=fake_db(rows=grown)):
                runner.SupabaseSelectedTests.tearDownClass()
            changed = dict(grown, revisions={"outra"})
            with patch.object(runner, "connect", return_value=fake_db(rows=changed)):
                with self.assertRaises(runner.Refused):
                    runner.SupabaseSelectedTests.tearDownClass()

    def test_jit_pacing_only_when_enabled_and_without_retry(self):
        clock = [100.0]
        sleeps = []
        def sleep(s):
            sleeps.append(round(s, 3)); clock[0] += s
        with patch.object(runner.time, "monotonic", side_effect=lambda: clock[0]), patch.object(runner.time, "sleep", side_effect=sleep), patch.object(runner, "_JIT_LAST", [float("-inf")]):
            with patch.object(runner, "JIT_MIN_INTERVAL", 0.0):
                runner.pace_jit(); runner.pace_jit()
            self.assertEqual(sleeps, [])
            with patch.object(runner, "JIT_MIN_INTERVAL", 2.0):
                runner.pace_jit(); clock[0] += 0.5; runner.pace_jit()
            self.assertEqual(sleeps, [1.5])

    def test_cli_selected_is_exclusive_and_sets_pacing(self):
        for extra in (["--run-fixtures"], ["--inspect"], ["--recover"]):
            argv = ["supabase_runner.py", "--config", "x.toml", "--run-selected", *extra]
            with self.subTest(extra=extra), patch.object(sys, "argv", argv), patch.object(runner, "load_config") as load, contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(runner.main(), 1)
                load.assert_not_called()
        ran = MagicMock(); ran.run.return_value.wasSuccessful.return_value = True
        argv = ["supabase_runner.py", "--config", "x.toml", "--run-selected"]
        with patch.object(sys, "argv", argv), patch.object(runner, "load_config"), patch.object(runner, "JIT_MIN_INTERVAL", 0.0), patch.object(runner.unittest, "TextTestRunner", return_value=ran):
            self.assertEqual(runner.main(), 0)
            self.assertEqual(runner.JIT_MIN_INTERVAL, runner.SELECTED_JIT_INTERVAL)
            self.assertEqual(ran.run.call_args.args[0].countTestCases(), 3)


if __name__ == "__main__": unittest.main(verbosity=2)
