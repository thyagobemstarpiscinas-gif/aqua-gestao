"""Sonda de diagnóstico: mede a limpeza do AppTest quando o runner lança erro.

Erro sintético, sem executar app.py, sem modificar fixtures ou o Streamlit instalado.
"""
import json
import socket
import threading
import sys
from unittest.mock import patch


def negar(*a, **k):
    raise RuntimeError("Sonda sem rede")


socket.socket.connect = negar
socket.socket.connect_ex = negar
socket.create_connection = negar
import googleapiclient.discovery
googleapiclient.discovery.build = negar
import streamlit as st
from streamlit.runtime import Runtime
from streamlit.testing.v1 import AppTest
from streamlit.testing.v1.local_script_runner import LocalScriptRunner
from apptest_protegido import instalar

instalar()

if "--runner-preso" in sys.argv:
    def preso(runner, *args, **kwargs):
        runner._script_thread = threading.Thread(target=threading.Event().wait,
                                                 name="runner-sintetico-preso", daemon=True)
        runner._script_thread.start()
        raise RuntimeError("ERRO_COM_THREAD_PRESA")
    with patch.object(LocalScriptRunner, "run", preso):
        try:
            AppTest.from_string("pass", default_timeout=300).run()
        finally:
            print("ERRO: CLEANUP_EXECUTADO_COM_THREAD_PRESA", flush=True)
    raise AssertionError("processo deveria ter sido interrompido")

runtime_antes, secrets_antes = Runtime._instance, st.secrets
threads_antes = len(threading.enumerate())
try:
    app = AppTest.from_string("import streamlit as st\nst.write('dado ficticio')", default_timeout=300)
    app.secrets["admin"] = {"pin": "pin-ficticio-sonda"}
    with patch.object(LocalScriptRunner, "run", side_effect=RuntimeError("ERRO_SINTETICO_DO_RUNNER")):
        try:
            app.run()
        except RuntimeError as e:
            assert str(e) == "ERRO_SINTETICO_DO_RUNNER"
            assert Runtime._instance is runtime_antes
            assert st.secrets is secrets_antes
            assert not any(t.name in ("ScriptRunner.scriptThread", "apptest-limite")
                           for t in threading.enumerate())
            print(json.dumps({"erro": str(e), "app_executado": False,
                "runtime_antes_nulo": runtime_antes is None,
                "runtime_depois_nulo": Runtime._instance is None,
                "runtime_restaurado": Runtime._instance is runtime_antes,
                "secrets_restaurados": st.secrets is secrets_antes,
                "threads_antes": threads_antes, "threads_depois": len(threading.enumerate())}))
finally:
    # Limpa somente os objetos criados por esta sonda descartável.
    Runtime._instance, st.secrets = runtime_antes, secrets_antes
