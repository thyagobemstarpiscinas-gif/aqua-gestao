"""Regressoes offline JIT: drivers simulados e DPAPI Windows com valor sintetico."""
import contextlib
import ctypes
from ctypes import wintypes as w
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import psycopg
import jit_dpapi
import supabase_runner as runner

# Valor sintetico; nao e token emitido pelo Supabase.
SYNTHETIC = "sbp_" + "offline_test_only_" * 3


def config():
    temp = Path(tempfile.gettempdir()).resolve()
    return dict(environment="homologacao", project_name="acqua-gestao-homologacao",
                project_ref="xfumoyxmhpxmqidmxjhq", host="aws-1-sa-east-1.pooler.supabase.com",
                port=5432, database="postgres", marker="1" * 32, admin_user="postgres",
                connection_mode="session_pooler", test_auth_mode="jit_dpapi",
                ca_file=str(temp / "offline-ca.crt"), passfile=str(temp / "offline-admin.pgpass"),
                jit_token_file=str(temp / "offline-token.dpapi"))


def backend(cfg, user, *, identity=None, backend_tls=True, client_tls=True,
            sslmode="verify-full", effective=(True, True, "on", "on"), admin=False):
    c = MagicMock()
    c.pgconn.ssl_in_use = client_tls
    c.info.get_parameters.return_value = {"sslmode": sslmode}
    c.info.transaction_status = runner.TransactionStatus.IDLE
    c.execute.return_value.fetchone.side_effect = [
        identity or (user, user, "postgres", "acqua-gestao-homologacao:" + cfg["marker"]),
        (backend_tls,), effective, (True, admin, False, False, False, False), (admin,)]
    return c


class JitTests(unittest.TestCase):
    def setUp(self):
        self.cfg = config()
        self.addCleanup(patch.stopall)
        patch.object(runner, "CONFIG", self.cfg).start()
        patch.object(runner, "LOCAL", False).start()

    def test_explicit_mode_and_exact_destination_required(self):
        self.assertEqual(runner.validate(self.cfg), self.cfg)
        for changes in ({"test_auth_mode": "jit"}, {"jit_token_file": "relative.dpapi"},
                        {"connection_mode": "direct"}, {"port": 6543},
                        {"project_ref": "a" * 20},
                        {"host": "aws-2-sa-east-1.pooler.supabase.com"},
                        {"host": self.cfg["host"] + ".example"}):
            with self.subTest(changes=changes), self.assertRaises(runner.Refused):
                runner.validate(dict(self.cfg, **changes))
        for key in ("jit_token_file", "test_auth_mode"):
            candidate = dict(self.cfg); candidate.pop(key)
            with self.assertRaises(runner.Refused):
                runner.validate(candidate)
        with self.assertRaises(runner.Refused):
            runner.validate(self.cfg, local=True)

    def test_jit_only_for_three_real_principals(self):
        for role in runner.ROLES:
            c = backend(self.cfg, role)
            with self.subTest(role=role), patch.object(jit_dpapi, "read_token", return_value=SYNTHETIC) as decrypt, patch.object(runner.psycopg, "connect", return_value=c) as dial:
                self.assertIs(runner.connect(role, read_only=True), c)
                decrypt.assert_called_once_with(self.cfg["jit_token_file"])
                options = dial.call_args.kwargs
                self.assertEqual(options["user"], role + "." + self.cfg["project_ref"])
                self.assertEqual(options["options"], "-c jit=true")
                self.assertEqual(options["password"], SYNTHETIC)
                self.assertEqual(options["passfile"], os.devnull)
                self.assertEqual(options["sslmode"], "verify-full")
                self.assertEqual(options["sslrootcert"], self.cfg["ca_file"])
                self.assertIn("BEGIN READ ONLY", [call.args[0] for call in c.execute.call_args_list])
                self.assertFalse(any("SET ROLE" in call.args[0] or "SESSION AUTHORIZATION" in call.args[0] for call in c.execute.call_args_list))

    def test_admin_uses_existing_passfile_without_decrypting(self):
        for principal in (None, "postgres"):
            c = backend(self.cfg, "postgres")
            with self.subTest(principal=principal), patch.object(jit_dpapi, "read_token") as decrypt, patch.object(runner.psycopg, "connect", return_value=c) as dial:
                runner.connect(principal, read_only=True)
                decrypt.assert_not_called()
                self.assertEqual(dial.call_args.kwargs["passfile"], self.cfg["passfile"])
                self.assertEqual(dial.call_args.kwargs["options"], "")
                self.assertNotIn("password", dial.call_args.kwargs)

    def test_unknown_probe_and_wire_names_refused_before_any_decrypt_or_dial(self):
        for role in ("rt_jit_probe", "service_role", "rt_test_a." + self.cfg["project_ref"]):
            with self.subTest(role=role), patch.object(jit_dpapi, "read_token") as decrypt, patch.object(runner.psycopg, "connect") as dial:
                with self.assertRaises(runner.Refused): runner.connect(role)
                decrypt.assert_not_called(); dial.assert_not_called()

    def test_missing_corrupt_or_inaccessible_dpapi_never_falls_back(self):
        for error in (FileNotFoundError(SYNTHETIC), PermissionError(SYNTHETIC), jit_dpapi.TokenRefused(SYNTHETIC)):
            with self.subTest(error=type(error).__name__), patch.object(jit_dpapi, "read_token", side_effect=error), patch.object(runner.psycopg, "connect") as dial:
                with self.assertRaises(runner.Refused) as caught: runner.connect("rt_test_a")
                dial.assert_not_called()
                self.assertNotIn(SYNTHETIC, str(caught.exception))

    def test_auth_failure_attempts_once_and_sanitizes_error(self):
        with patch.object(jit_dpapi, "read_token", return_value=SYNTHETIC), patch.object(runner.psycopg, "connect", side_effect=psycopg.OperationalError(SYNTHETIC)) as dial:
            with self.assertRaises(runner.Refused) as caught: runner.connect("rt_test_a")
            dial.assert_called_once()
            self.assertNotIn(SYNTHETIC, str(caught.exception))
            self.assertIsNone(caught.exception.__cause__)

    def test_jit_preserves_identity_marker_client_and_backend_tls_guards(self):
        role = "rt_test_a"
        valid = (role, role, "postgres", "acqua-gestao-homologacao:" + self.cfg["marker"])
        cases = [{"identity": ("postgres", "postgres", *valid[2:])},
                 {"identity": (role, "postgres", *valid[2:])},
                 {"identity": (role, role, "other", valid[3])},
                 {"identity": (*valid[:3], "wrong-marker")},
                 {"backend_tls": False}, {"client_tls": False}, {"sslmode": "require"}]
        for changes in cases:
            c = backend(self.cfg, role, **changes)
            with self.subTest(changes=changes), patch.object(jit_dpapi, "read_token", return_value=SYNTHETIC), patch.object(runner.psycopg, "connect", return_value=c):
                with self.assertRaises(runner.Refused): runner.connect(role, read_only=True)
                c.close.assert_called_once()
                self.assertFalse(any(call.args[0].startswith("SET ") for call in c.execute.call_args_list))

    def test_jit_refuses_ignored_sql_timeouts_and_read_only(self):
        for effective in ((False, True, "on", "on"), (True, False, "on", "on"),
                          (True, True, "off", "on"), (True, True, "on", "off")):
            c = backend(self.cfg, "rt_test_a", effective=effective)
            with self.subTest(effective=effective), patch.object(jit_dpapi, "read_token", return_value=SYNTHETIC), patch.object(runner.psycopg, "connect", return_value=c):
                with self.assertRaises(runner.Refused): runner.connect("rt_test_a", read_only=True)
                c.close.assert_called_once()

    def test_jit_admin_attributes_still_refused(self):
        c = backend(self.cfg, "rt_test_a", admin=True)
        with patch.object(jit_dpapi, "read_token", return_value=SYNTHETIC), patch.object(runner.psycopg, "connect", return_value=c):
            with self.assertRaises(runner.Refused): runner.connect("rt_test_a", read_only=True)
            c.close.assert_called_once()

    def test_legacy_password_mode_keeps_original_routing(self):
        self.cfg.pop("test_auth_mode"); self.cfg.pop("jit_token_file")
        for role in runner.ROLES:
            c = backend(self.cfg, role)
            with self.subTest(role=role), patch.object(jit_dpapi, "read_token") as decrypt, patch.object(runner.psycopg, "connect", return_value=c) as dial:
                runner.connect(role)
                decrypt.assert_not_called()
                self.assertNotIn("password", dial.call_args.kwargs)
                self.assertEqual(dial.call_args.kwargs["options"], "")
                self.assertEqual(dial.call_args.kwargs["passfile"], self.cfg["passfile"])

    def test_recovery_reuses_only_manifest_and_sanitized_environment(self):
        c = MagicMock(); c.__enter__.return_value = c
        private_manifest = str(Path(tempfile.gettempdir()).resolve() / "offline-jit.toml")
        with patch.object(runner, "CONFIG_PATH", private_manifest), patch.object(runner, "connect", return_value=c), patch.object(runner.base, "save"), patch.dict(os.environ, {"PGPASSWORD": SYNTHETIC, "SUPABASE_ACCESS_TOKEN": SYNTHETIC, "RT_JIT_TOKEN": SYNTHETIC}), patch.object(runner.subprocess, "run", return_value=MagicMock(stdout="RECOVERY_OK")) as child:
            runner.SupabaseTests("test_new_process_without_local_files_recovers").test_new_process_without_local_files_recovers()
            call = child.call_args
            self.assertEqual(call.args[0][2:], ["--config", private_manifest, "--recover"])
            self.assertNotIn(SYNTHETIC, repr(call.args))
            self.assertNotIn(SYNTHETIC, repr(call.kwargs))
            self.assertFalse(any(k.upper().startswith("PG") for k in call.kwargs["env"]))
            self.assertFalse(Path(call.kwargs["cwd"]).exists())
            self.assertEqual(self.cfg["test_auth_mode"], "jit_dpapi")

    def test_fifteen_scenarios_and_original_recovery_assertion_remain(self):
        names = unittest.defaultTestLoader.getTestCaseNames(runner.SupabaseTests)
        self.assertEqual(len(names), 15)
        self.assertIn("test_new_process_without_local_files_recovers", names)


class TokenParsingTests(unittest.TestCase):
    def test_raw_and_powershell_dpapi_encodings(self):
        cipher = b"\x01\x00\x00\x00" + b"\xab" * 20
        self.assertEqual(jit_dpapi._ciphertext(cipher), (cipher, False))
        for raw in (cipher.hex().encode(), b"\xef\xbb\xbf" + cipher.hex().encode(), cipher.hex().encode("utf-16")):
            self.assertEqual(jit_dpapi._ciphertext(raw), (cipher, True))

    def test_plaintext_aes_securestring_and_bad_hex_refused(self):
        for value in (SYNTHETIC.encode(), b"01000000xz", b"010000001", b"76492d1116743f0423413b16050a5345"):
            with self.subTest(value_type="synthetic"), self.assertRaises(Exception): jit_dpapi._ciphertext(value)

    def test_invalid_decrypted_values_never_return_or_leak(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "synthetic.dpapi"; p.write_bytes(b"cipher")
            for token in ("", "x", "sbp_short", SYNTHETIC + "\n", " " + SYNTHETIC, "ey.fake.jwt", SYNTHETIC + "\x00"):
                with self.subTest(kind="synthetic"), patch.object(jit_dpapi, "private_token_file", return_value=p), patch.object(jit_dpapi, "_ciphertext", return_value=(b"cipher", False)), patch.object(jit_dpapi, "_unprotect", return_value=token):
                    with self.assertRaises(jit_dpapi.TokenRefused) as caught: jit_dpapi.read_token(str(p))
                    self.assertNotIn(SYNTHETIC, str(caught.exception))

    def test_dpapi_errors_sanitized_and_no_plaintext_fallback(self):
        with patch.object(jit_dpapi, "private_token_file", side_effect=OSError(SYNTHETIC)), patch.object(jit_dpapi, "_unprotect") as decrypt:
            with self.assertRaises(jit_dpapi.TokenRefused) as caught: jit_dpapi.read_token("missing")
            decrypt.assert_not_called(); self.assertNotIn(SYNTHETIC, str(caught.exception))


@unittest.skipUnless(os.name == "nt", "DPAPI integra somente Windows")
class WindowsDpapiTests(unittest.TestCase):
    @staticmethod
    def protect(text, secure_string=True):
        raw = text.encode("utf-16-le" if secure_string else "utf-8")
        buf = ctypes.create_string_buffer(raw)
        source = jit_dpapi.Blob(len(raw), ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte)))
        target = jit_dpapi.Blob()
        crypt = ctypes.WinDLL("crypt32"); kernel = ctypes.WinDLL("kernel32")
        crypt.CryptProtectData.argtypes = [ctypes.POINTER(jit_dpapi.Blob), w.LPCWSTR, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, w.DWORD, ctypes.POINTER(jit_dpapi.Blob)]
        crypt.CryptProtectData.restype = w.BOOL
        kernel.LocalFree.argtypes = [ctypes.c_void_p]
        kernel.LocalFree.restype = ctypes.c_void_p
        if not crypt.CryptProtectData(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(target)):
            raise RuntimeError("synthetic DPAPI protection failed")
        try:
            cipher = ctypes.string_at(target.data, target.size)
            return cipher.hex().encode() if secure_string else cipher
        finally: kernel.LocalFree(target.data)

    @staticmethod
    def restrict(path):
        sid = subprocess.check_output(["powershell.exe", "-NoProfile", "-Command", "[Security.Principal.WindowsIdentity]::GetCurrent().User.Value"], text=True).strip()
        subprocess.run(["icacls.exe", str(path), "/inheritance:r", "/grant:r", "*" + sid + ":(F)"], check=True, capture_output=True)

    def test_real_windows_dpapi_and_private_acl_both_formats(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "synthetic.dpapi"
            for secure in (True, False):
                p.write_bytes(self.protect(SYNTHETIC, secure)); self.restrict(p)
                self.assertEqual(jit_dpapi.read_token(str(p)), SYNTHETIC)

    def test_real_windows_acl_rejects_everyone_and_missing_file(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "synthetic.dpapi"
            with self.assertRaises(jit_dpapi.TokenRefused): jit_dpapi.read_token(str(p))
            p.write_bytes(self.protect(SYNTHETIC)); self.restrict(p)
            subprocess.run(["icacls.exe", str(p), "/grant", "*S-1-1-0:(R)"], check=True, capture_output=True)
            with self.assertRaises(jit_dpapi.TokenRefused): jit_dpapi.read_token(str(p))

    def test_token_file_cannot_be_inside_a_worktree(self):
        with self.assertRaises(jit_dpapi.TokenRefused):
            jit_dpapi.private_token_file(str(Path(runner.__file__).resolve()))

    def test_real_windows_corrupt_ciphertext_fails_closed(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "synthetic.dpapi"
            p.write_bytes(b"\x01\x00\x00\x00" + b"invalid" * 12); self.restrict(p)
            with self.assertRaises(jit_dpapi.TokenRefused): jit_dpapi.read_token(str(p))

    def test_real_subprocess_reloads_manifest_and_dpapi_from_empty_cwd(self):
        with tempfile.TemporaryDirectory() as d:
            folder = Path(d); cfg = config()
            p = folder / "synthetic.dpapi"; p.write_bytes(self.protect(SYNTHETIC)); self.restrict(p)
            cfg.update(jit_token_file=str(p), ca_file=str(folder / "ca"), passfile=str(folder / "admin.pgpass"))
            Path(cfg["ca_file"]).write_text("offline only")
            Path(cfg["passfile"]).write_text("offline only")
            manifest = folder / "authorization.toml"
            manifest.write_text("\n".join(k + " = " + json.dumps(v) for k, v in cfg.items()), encoding="utf-8")
            child = folder / "offline_recovery.py"
            # Real child/main/DPAPI/verify_marker; only DB transport replaced. No network.
            child.write_text('''import sys
from unittest.mock import MagicMock
import supabase_runner as r
EXPECTED = "sbp_" + "offline_test_only_" * 3
def dial(**kw):
    assert kw['password'] == EXPECTED
    assert kw['options'] == '-c jit=true'
    assert kw['passfile'] == __import__('os').devnull
    assert kw['user'] == 'rt_test_a.xfumoyxmhpxmqidmxjhq'
    c=MagicMock();c.__enter__.return_value=c
    c.pgconn.ssl_in_use=True;c.info.get_parameters.return_value={'sslmode':'verify-full'}
    c.execute.return_value.fetchone.side_effect=[
      ('rt_test_a','rt_test_a','postgres','acqua-gestao-homologacao:'+'1'*32),
      (True,), (True,True,'off','off'), (True,False,False,False,False,False), (False,),
      ([{'payload':{'campos':{'rel_diagnostico':'RECUPERAVEL_SEM_DISCO_STREAMLIT'}}}],),
      ({'revision':1},)]
    return c
r.psycopg.connect=dial
raise SystemExit(r.main())
''', encoding="utf-8")
            empty = folder / "empty"; empty.mkdir()
            env = runner.recovery_environment()
            env["PYTHONPATH"] = os.pathsep.join([str(Path(runner.__file__).parent), *sys.path])
            result = subprocess.run([sys.executable, "-B", str(child), "--config", str(manifest), "--recover"], cwd=empty, env=env, capture_output=True, text=True, timeout=45)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("RECOVERY_OK", result.stdout)
            self.assertNotIn(SYNTHETIC, result.stdout + result.stderr)
            self.assertEqual(list(empty.iterdir()), [])


if __name__ == "__main__": unittest.main(verbosity=2)
