"""Proteção processual dos AppTests; não altera o Streamlit instalado."""
import faulthandler
import functools
import os
import sys
import threading


def _fatal(motivo):
    # Não usar uma exceção: unittest continuaria e executaria os cleanups.
    sys.stderr.write("\nFATAL AppTest: " + motivo + "\n")
    faulthandler.dump_traceback(file=sys.stderr, all_threads=True)
    sys.stderr.flush()
    os._exit(86)


def instalar():
    import streamlit as st
    from streamlit.runtime import Runtime
    from streamlit.runtime.pages_manager import PagesManager
    from streamlit.testing.v1 import AppTest
    from streamlit.testing.v1.local_script_runner import LocalScriptRunner

    if getattr(AppTest._run, "_protegido", False):
        return
    original = AppTest._run
    init = LocalScriptRunner.__init__
    run = LocalScriptRunner.run
    local = threading.local()

    def encerrar(runners):
        try:
            for runner in runners:
                thread = runner._script_thread
                if thread is not None and thread.is_alive():
                    runner.request_stop()
                    thread.join(5)
                if thread is not None and thread.is_alive():
                    _fatal("ScriptRunner ainda ativo antes da restauração dos globais")
        except BaseException as erro:
            _fatal("não foi possível confirmar encerramento do runner: " + repr(erro))

    @functools.wraps(run)
    def rodar(runner, *args, **kwargs):
        try:
            return run(runner, *args, **kwargs)
        finally:
            # Antes de o próprio AppTest restaurar seus globais/configurações.
            encerrar([runner])

    @functools.wraps(init)
    def registrar(runner, *args, **kwargs):
        init(runner, *args, **kwargs)
        if hasattr(local, "runners"):
            local.runners.append(runner)

    @functools.wraps(original)
    def executar(app, widget_state=None, timeout=None):
        if hasattr(local, "runners"):
            _fatal("AppTests aninhados não são suportados")
        runtime, secrets = Runtime._instance, st.secrets
        pages = PagesManager.uses_pages_directory
        runners = local.runners = []
        terminou = threading.Event()
        limite = app.default_timeout if timeout is None else timeout

        def vigiar():
            if terminou.wait(limite):
                return
            sys.stderr.write("\nAppTest atingiu o limite; solicitando parada do runner.\n")
            faulthandler.dump_traceback(file=sys.stderr, all_threads=True)
            for runner in list(runners):
                runner.request_stop()
            # Apenas margem para encerrar, nunca para continuar o script/teste.
            if not terminou.wait(5):
                _fatal("runner/chamada não encerrou após limite e solicitação de parada; suíte interrompida")

        watcher = threading.Thread(target=vigiar, name="apptest-limite", daemon=True)
        watcher.start()
        try:
            return original(app, widget_state, timeout)
        finally:
            # Só devolver globais e permitir cleanups depois do encerramento real.
            encerrar(runners)
            Runtime._instance, st.secrets = runtime, secrets
            PagesManager.uses_pages_directory = pages
            del local.runners
            terminou.set()
            watcher.join()

    executar._protegido = True
    LocalScriptRunner.__init__ = registrar
    LocalScriptRunner.run = rodar
    AppTest._run = executar
