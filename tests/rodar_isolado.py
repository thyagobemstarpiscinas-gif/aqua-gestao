"""Executa a suíte sem rede e sem credenciais de produção.

Uso: python tests/rodar_isolado.py
     python tests/rodar_isolado.py --modules test_rt_protecao_integrada
     python tests/rodar_isolado.py --diagnostics-dir PASTA_NOVA
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
    parser.add_argument("--diagnostics-dir")
    parser.add_argument("--inventory-json")
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
    if args.inventory_json:
        import collections
        loader = unittest.TestLoader()
        def flatten(suite):
            for item in suite:
                if isinstance(item, unittest.TestSuite):
                    yield from flatten(item)
                else:
                    yield item
        tests = list(flatten(loader.discover('tests')))
        if loader.errors:
            raise RuntimeError('\n'.join(loader.errors))
        rows, positions, unique = [], collections.defaultdict(list), {}
        for index, test in enumerate(tests, 1):
            cls = type(test)
            app = (any(base.__name__ == '_AppReal' for base in cls.__mro__)
                   or cls.__module__.endswith('apptest')
                   or getattr(cls.setUp, '__module__', '') == 'test_relatorio_rt_campos')
            row = {'id': test.id(), 'classe': cls.__module__ + '.' + cls.__qualname__,
                   'modulo': cls.__module__, 'apptest': app, 'ordem': index}
            rows.append(row)
            unique.setdefault(test.id(), row)
            positions[test.id()].append(index)
        inventory = {'descobertos': len(rows), 'unicos': len(unique),
            'diferenca_contagem_anterior_213': len(rows)-213, 'ocorrencias': rows,
            'duplicados': {key: value for key, value in positions.items() if len(value) > 1},
            'testes': list(unique.values())}
        with Path(args.inventory_json).open('x', encoding='utf-8') as stream:
            json.dump(inventory, stream, ensure_ascii=False, indent=2)
        print(f"DESCOBERTOS={len(rows)} UNICOS={len(unique)}", flush=True)
        return 0
    if args.diagnostics_dir:
        from diagnosticar_timeouts import main as diagnosticar
        seletores = args.modules or ["--full"]
        return diagnosticar(["--output-dir", args.diagnostics_dir, *seletores])
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
