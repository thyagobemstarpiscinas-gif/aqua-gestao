"""Instrumentação com a proteção local de encerramento/limpeza dos AppTests.

Uso: python tests/diagnosticar_timeouts.py --output-dir PASTA teste1 teste2
Mesmos bloqueios de rede/Drive de rodar_isolado.py. Suíte completa só com --full.
"""
import argparse
import ast
import collections
import contextlib
import ctypes
import faulthandler
import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import threading
import time
import traceback
import unittest
from unittest import mock
from urllib.parse import urlsplit


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--full", action="store_true")
    parser.add_argument("tests", nargs="*")
    args = parser.parse_args(argv)
    if not args.full and not args.tests:
        parser.error("informe seletores ou --full")
    root = Path(__file__).resolve().parents[1]
    os.chdir(root)
    sys.path[:0] = [str(root / "tests"), str(root)]
    output = Path(args.output_dir)
    output.mkdir(parents=True, exist_ok=False)
    os.environ.update(CODESPACES="true", HTTP_PROXY="http://127.0.0.1:9",
                      HTTPS_PROXY="http://127.0.0.1:9", NO_PROXY="localhost,127.0.0.1")
    events = (output / "eventos.jsonl").open("w", encoding="utf-8", buffering=1)
    stacks = (output / "stacks.txt").open("w", encoding="utf-8", buffering=1)
    lock = threading.RLock()
    counter = collections.Counter()
    active = {}
    run_number = 0
    current_test = None
    t0 = time.perf_counter()

    def emit(event, **data):
        with lock:
            events.write(json.dumps({"event": event, "wall": time.time(), "perf": time.perf_counter() - t0,
                "cpu": time.process_time(), "thread": threading.get_ident(), "test": current_test,
                **data}, ensure_ascii=False, default=str) + "\n")

    def memory():
        if os.name != "nt":
            return {}
        class Counters(ctypes.Structure):
            _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong),
                *[(n, ctypes.c_size_t) for n in ("PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage", "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]]
        c = Counters()
        c.cb = ctypes.sizeof(c)
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.GetCurrentProcess.restype = ctypes.c_void_p
        psapi = ctypes.WinDLL("psapi", use_last_error=True)
        psapi.GetProcessMemoryInfo.argtypes = [ctypes.c_void_p, ctypes.POINTER(Counters), ctypes.c_ulong]
        if psapi.GetProcessMemoryInfo(kernel.GetCurrentProcess(), ctypes.byref(c), c.cb):
            return {"rss": c.WorkingSetSize, "peak_rss": c.PeakWorkingSetSize, "commit": c.PagefileUsage}
        return {"error": ctypes.get_last_error()}

    def snapshot():
        state = {"threads": [{"name": t.name, "id": t.ident, "alive": t.is_alive(), "daemon": t.daemon}
                             for t in threading.enumerate()], "memory": memory(),
                 "mocks": [{"class": type(p).__name__, "attribute": getattr(p, "attribute", None)}
                           for p in mock._patch._active_patches], "calls": dict(counter),
                 "codespaces": os.environ.get("CODESPACES")}
        for name in ("gspread", "google.oauth2.service_account"):
            mod = sys.modules.get(name)
            state[name] = {"id": id(mod), "file": getattr(mod, "__file__", None),
                           "fake_sheet": hasattr(mod, "PLANILHA")}
        oauth = sys.modules.get("google.oauth2")
        state["service_account_parent_matches"] = (getattr(oauth, "service_account", None)
                                                   is sys.modules.get("google.oauth2.service_account"))
        st = sys.modules.get("streamlit")
        if st:
            state["secrets_id"] = id(st.secrets)  # nunca registra conteúdo de secrets
            from streamlit.runtime import Runtime
            state["runtime_id"] = id(Runtime._instance) if Runtime._instance else None
            for name, attr in (("cache_data_api", "_data_caches"), ("cache_resource_api", "_resource_caches")):
                module = sys.modules.get("streamlit.runtime.caching." + name)
                manager = getattr(module, attr, None)
                state[name] = {"manager_id": id(manager), "caches": len(getattr(manager, "_function_caches", {}))}
        return state

    def dump(reason):
        with lock:
            stacks.write(f"\n=== {reason}; wall={time.time()}; perf={time.perf_counter()-t0}; cpu={time.process_time()}; test={current_test} ===\n")
            names = {t.ident: t.name for t in threading.enumerate()}
            for ident, frame in sys._current_frames().items():
                stacks.write(f"\nTHREAD {ident} {names.get(ident, '?')}\n")
                stacks.writelines(traceback.format_stack(frame))
            stacks.flush()
        emit("stack", reason=reason)

    def deny(kind):
        def blocked(*_a, **_k):
            counter[kind] += 1
            if counter[kind] <= 3:
                emit("network_blocked", kind=kind, stack=traceback.format_stack(limit=10))
            raise RuntimeError("Teste isolado: acesso à rede bloqueado")
        return blocked

    socket.socket.connect = deny("socket.connect")
    socket.socket.connect_ex = deny("socket.connect_ex")
    socket.create_connection = deny("socket.create_connection")
    import googleapiclient.discovery
    googleapiclient.discovery.build = deny("drive.build")
    from streamlit.testing.v1 import AppTest
    from apptest_protegido import instalar
    instalar()
    from streamlit.testing.v1.local_script_runner import LocalScriptRunner
    from streamlit.runtime.scriptrunner import magic
    from streamlit.runtime.scriptrunner.script_cache import ScriptCache

    original_run = LocalScriptRunner.run
    original_app_run = AppTest._run
    original_sleep = time.sleep
    original_popen = subprocess.Popen.__init__
    original_cleanups = unittest.TestCase.doCleanups
    original_parse = ast.parse
    original_magic = magic.add_magic
    original_bytecode = ScriptCache.get_bytecode
    import requests.sessions
    original_request = requests.sessions.Session.request

    def script_run(runner, *a, **k):
        nonlocal run_number
        run_number += 1
        number = run_number
        started = time.perf_counter()
        active[number] = (started, runner)
        emit("script_start", run=number, timeout=k.get("timeout"), script=runner.script_path,
             snapshot=snapshot())
        try:
            return original_run(runner, *a, **k)
        except BaseException as e:
            emit("script_error", run=number, error=repr(e))
            dump("script_error")
            raise
        finally:
            active.pop(number, None)
            th = runner._script_thread
            emit("script_end", run=number, elapsed=time.perf_counter() - started,
                 stopped=runner.script_stopped(), script_thread_alive=th.is_alive() if th else None,
                 events=[str(e) for e in runner.events[-5:]], snapshot=snapshot())

    def app_run(app, *a, **k):
        before = snapshot()
        try:
            return original_app_run(app, *a, **k)
        finally:
            emit("app_return", before=before, after=snapshot())

    def sleep(seconds):
        if seconds >= .05:
            counter["sleep"] += 1
            emit("sleep", seconds=seconds, stack=traceback.format_stack(limit=5))
        return original_sleep(seconds)

    def popen(process, *a, **k):
        counter["subprocess"] += 1
        command = a[0] if a else k.get("args")
        executable = command[0] if isinstance(command, (list, tuple)) and command else str(command).split(" ")[0]
        emit("subprocess_start", executable=str(executable))
        result = original_popen(process, *a, **k)
        emit("subprocess_created", pid=process.pid)
        return result

    def request(session, method, url, *a, **k):
        counter["http"] += 1
        started = time.perf_counter()
        emit("http_start", method=method, host=urlsplit(url).hostname)
        try:
            return original_request(session, method, url, *a, **k)
        finally:
            emit("http_end", elapsed=time.perf_counter() - started)

    def cleanups(case):
        path = getattr(case, "pasta", None)
        before = snapshot()
        started = time.perf_counter()
        try:
            return original_cleanups(case)
        finally:
            emit("cleanups", elapsed=time.perf_counter() - started, before=before, after=snapshot(),
                 temp=str(path) if path else None, temp_exists=path.exists() if path else None)

    def phase(name, function, *a, **k):
        started, cpu = time.perf_counter(), time.process_time()
        emit("phase_start", name=name)
        try:
            return function(*a, **k)
        finally:
            emit("phase_end", name=name, elapsed=time.perf_counter()-started,
                 cpu_elapsed=time.process_time()-cpu)

    def parse(source, filename="<unknown>", *a, **k):
        if str(filename).endswith("app.py"):
            return phase("ast.parse", original_parse, source, filename, *a, **k)
        return original_parse(source, filename, *a, **k)

    def magic_code(source, path):
        return phase("magic.add_magic", original_magic, source, path)

    def bytecode(cache, path):
        emit("bytecode_cache", cache_id=id(cache), hit=os.path.abspath(path) in cache._cache)
        return phase("ScriptCache.get_bytecode", original_bytecode, cache, path)

    # Eventos locais de funções relevantes: não altera seus corpos ou retornos.
    monitoring = getattr(sys, "monitoring", None)
    tool_id = 5
    tracked = {"converter_docx_para_pdf", "_renderizar_relatorio_rt", "preencher_relatorio_mensal_docx",
               "gerar_pdf_relatorio_rt_premium_reportlab", "_relatorio_rt_sincronizar_pendentes", "conectar_sheets"}
    enabled = False
    if monitoring and monitoring.get_tool(tool_id) is None:
        monitoring.use_tool_id(tool_id, "diagnostico_apptest")
        def entered(code, *_):
            counter["function:" + code.co_name] += 1
            emit("function_start", name=code.co_name)
        def returned(code, *_):
            emit("function_end", name=code.co_name)
        monitoring.register_callback(tool_id, monitoring.events.PY_START, entered)
        monitoring.register_callback(tool_id, monitoring.events.PY_RETURN, returned)
        enabled = True
        def audit(event, arg):
            if enabled and event == "exec":
                def visit(code):
                    if getattr(code, "co_name", None) in tracked and code.co_filename.endswith("app.py"):
                        monitoring.set_local_events(tool_id, code, monitoring.events.PY_START | monitoring.events.PY_RETURN)
                    for c in getattr(code, "co_consts", ()):
                        if hasattr(c, "co_consts"):
                            visit(c)
                visit(arg[0])
        sys.addaudithook(audit)

    stop = threading.Event()
    def watch():
        last_dump = {}
        while not stop.wait(10):
            emit("sample", active=[{"run": n, "elapsed": time.perf_counter()-start,
                                   "runner_thread": getattr(runner._script_thread, "ident", None)}
                                  for n, (start, runner) in list(active.items())], snapshot=snapshot())
            for n, (start, _) in list(active.items()):
                elapsed = time.perf_counter() - start
                if elapsed >= 20 and elapsed - last_dump.get(n, -1000) >= 30:
                    dump(f"run {n}, elapsed {elapsed:.3f}")
                    last_dump[n] = elapsed
    watcher = threading.Thread(target=watch, name="diagnostico-watchdog", daemon=True)

    class Result(unittest.TextTestResult):
        def startTest(self, test):
            nonlocal current_test
            current_test = test.id()
            self.started = time.perf_counter()
            self.test_status = None
            emit("test_start", snapshot=snapshot())
            super().startTest(test)
        def stopTest(self, test):
            emit("test_end", elapsed=time.perf_counter()-self.started,
                 status=self.test_status or 'sem_resultado', snapshot=snapshot())
            super().stopTest(test)
        def addSuccess(self, test):
            self.test_status = 'aprovado'
            super().addSuccess(test)
        def addFailure(self, test, err):
            self.test_status = 'falha'
            emit('test_failure', error=''.join(traceback.format_exception(*err)))
            super().addFailure(test, err)
        def addSkip(self, test, reason):
            self.test_status = 'ignorado'
            emit('test_skip', reason=reason)
            super().addSkip(test, reason)
        def addExpectedFailure(self, test, err):
            self.test_status = 'falha_esperada'
            super().addExpectedFailure(test, err)
        def addUnexpectedSuccess(self, test):
            self.test_status = 'sucesso_inesperado'
            super().addUnexpectedSuccess(test)
        def addSubTest(self, test, subtest, err):
            if err is not None:
                self.test_status = 'falha'
                emit('test_subtest_failure', error=''.join(traceback.format_exception(*err)))
            super().addSubTest(test, subtest, err)
        def addError(self, test, err):
            self.test_status = 'erro'
            emit("test_error", error="".join(traceback.format_exception(*err)))
            dump("test_error")
            super().addError(test, err)

    emit("process_start", pid=os.getpid(), executable=sys.executable, tests=args.tests, snapshot=snapshot())
    suite = (unittest.defaultTestLoader.discover("tests") if args.full else
             unittest.TestSuite(unittest.defaultTestLoader.loadTestsFromName(n) for n in args.tests))
    watcher.start()
    faulthandler.dump_traceback_later(90, repeat=True, file=stacks)
    try:
        with contextlib.ExitStack() as patches:
            patches.enter_context(mock.patch.object(LocalScriptRunner, "run", script_run))
            patches.enter_context(mock.patch.object(AppTest, "_run", app_run))
            patches.enter_context(mock.patch.object(time, "sleep", sleep))
            patches.enter_context(mock.patch.object(subprocess.Popen, "__init__", popen))
            patches.enter_context(mock.patch.object(requests.sessions.Session, "request", request))
            patches.enter_context(mock.patch.object(unittest.TestCase, "doCleanups", cleanups))
            patches.enter_context(mock.patch.object(ast, "parse", parse))
            patches.enter_context(mock.patch.object(magic, "add_magic", magic_code))
            patches.enter_context(mock.patch.object(ScriptCache, "get_bytecode", bytecode))
            result = unittest.TextTestRunner(verbosity=2, resultclass=Result).run(suite)
        summary = {"run": result.testsRun, "failures": len(result.failures), "errors": len(result.errors),
            "skipped": [(t.id(), reason) for t, reason in result.skipped],
            "elapsed": time.perf_counter()-t0, "calls": dict(counter), "success": result.wasSuccessful()}
        (output / "resultado.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
        emit("process_end", summary=summary, snapshot=snapshot())
        return int(not result.wasSuccessful())
    finally:
        enabled = False
        faulthandler.cancel_dump_traceback_later()
        stop.set()
        watcher.join(2)
        if monitoring and monitoring.get_tool(tool_id) == "diagnostico_apptest":
            monitoring.free_tool_id(tool_id)
        events.close()
        stacks.close()


if __name__ == "__main__":
    raise SystemExit(main())
