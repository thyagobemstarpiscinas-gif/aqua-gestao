"""Executa a suíte sem rede e sem credenciais de produção.

Uso: python tests/rodar_isolado.py
     python tests/rodar_isolado.py --modules test_rt_protecao_integrada
"""
import argparse
import json
import os
from pathlib import Path
import socket
import sys
import unittest


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--modules", nargs="+")
    parser.add_argument("--result-json")
    args = parser.parse_args()
    raiz = Path(__file__).resolve().parents[1]
    os.chdir(raiz)
    sys.path[:0] = [str(raiz / "tests"), str(raiz)]
    os.environ.update(CODESPACES="true", HTTP_PROXY="http://127.0.0.1:9",
                      HTTPS_PROXY="http://127.0.0.1:9", NO_PROXY="localhost,127.0.0.1")

    def negar_rede(*_args, **_kwargs):
        raise RuntimeError("Teste isolado: acesso à rede bloqueado")

    socket.socket.connect = negar_rede
    socket.socket.connect_ex = negar_rede
    socket.create_connection = negar_rede
    import googleapiclient.discovery
    googleapiclient.discovery.build = negar_rede
    loader = unittest.defaultTestLoader
    suite = (unittest.TestSuite(loader.loadTestsFromName(n) for n in args.modules)
             if args.modules else loader.discover("tests"))
    resultado = unittest.TextTestRunner(verbosity=2).run(suite)
    if args.result_json:
        Path(args.result_json).write_text(json.dumps({
            "run": resultado.testsRun, "failures": len(resultado.failures),
            "errors": len(resultado.errors), "skipped": resultado.skipped,
            "success": resultado.wasSuccessful(),
        }), encoding="utf-8")
    return 0 if resultado.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
