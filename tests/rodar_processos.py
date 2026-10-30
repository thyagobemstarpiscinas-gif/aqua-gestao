"""Grupos sequenciais, cada um em um novo Python, sem remover temporários.

python tests/rodar_processos.py --inventory INVENTARIO.json --output-dir PASTA_NOVA
"""
import argparse
import collections
import csv
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time


DOCUMENTOS = ('test_rt_protecao_integrada.TestDocumentosProtegidos.'
              'test_d1_b_c_d_e_documentos_sem_dados_de_outro_cliente')


def publicar(path, value):
    temporary = path.with_suffix(path.suffix + '.tmp')
    with temporary.open('w', encoding='utf-8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def fontes(root):
    paths = [root / 'app.py', *sorted((root / 'utils').rglob('*.py')),
             root / 'requirements.txt', root / 'tests' / 'apptest_protegido.py']
    return {str(path.relative_to(root)): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--inventory', required=True)
    parser.add_argument('--output-dir', required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=False)
    inventory = json.loads(Path(args.inventory).read_text(encoding='utf-8'))
    rows = inventory['testes']
    ids = [row['id'] for row in rows]
    if len(set(ids)) != len(ids) or DOCUMENTOS not in ids:
        raise RuntimeError('Inventário único inválido ou teste de documentos ausente')
    publicar(output / 'inventario.json', inventory)
    before = fontes(root)
    publicar(output / 'fontes-antes.json', before)
    started = time.perf_counter()
    pending = [[DOCUMENTOS]]
    previous = None
    for row in rows:
        if row['id'] == DOCUMENTOS:
            continue
        key = ('app', row['classe']) if row['apptest'] else ('unit', row['modulo'])
        limit = 3 if row['apptest'] else 50
        if key != previous or len(pending[-1]) >= limit:
            pending.append([])
        pending[-1].append(row['id'])
        previous = key
    publicar(output / 'plano-inicial.json', pending)
    attempts, processes = [], []

    def consolidate():
        tests = []
        for row in rows:
            history = [item for item in attempts if item['teste'] == row['id']]
            last = history[-1] if history else {'teste': row['id'], 'resultado': 'nao_executado',
                     'duracao_s': None, 'grupo': None, 'codigo_saida': None}
            tests.append({**last, 'tentativas': len(history)})
        preserved = fontes(root) == before
        result = {'descobertos': inventory['descobertos'], 'unicos': len(rows),
                  'contagens': dict(collections.Counter(item['resultado'] for item in tests)),
                  'duracao_total_s': time.perf_counter() - started, 'testes': tests,
                  'tentativas': attempts, 'processos': processes, 'fontes_preservadas': preserved,
                  'sucesso': preserved and all(item['resultado'] == 'aprovado' for item in tests)
                      and all(item['codigo_saida'] == 0 and item['finalizacao_registrada'] for item in processes)}
        publicar(output / 'consolidado.json', result)
        temporary = output / 'resultados-por-teste.csv.tmp'
        with temporary.open('w', encoding='utf-8', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=['teste', 'resultado', 'duracao_s',
                'grupo', 'codigo_saida', 'tentativas'], extrasaction='ignore')
            writer.writeheader()
            writer.writerows(tests)
        os.replace(temporary, output / 'resultados-por-teste.csv')
        return result

    consolidate()
    print(f"INVENTARIO descobertos={inventory['descobertos']} unicos={len(rows)} "
          f"grupos_iniciais={len(pending)}", flush=True)
    number = 0
    while pending:
        names = pending.pop(0)
        number += 1
        group = f'grupo-{number:03d}'
        folder = output / group
        temporary = output / 'temporarios' / group
        temporary.mkdir(parents=True, exist_ok=False)
        env = os.environ.copy()
        env.update(TEMP=str(temporary), TMP=str(temporary), TMPDIR=str(temporary))
        for name in ('GOOGLE_APPLICATION_CREDENTIALS', 'GCP_SERVICE_ACCOUNT_JSON'):
            env.pop(name, None)
        command = [sys.executable, str(root / 'tests' / 'rodar_isolado.py'),
                   '--diagnostics-dir', str(folder), '--modules', *names]
        publicar(output / f'{group}-plano.json', {'grupo': group, 'testes': names, 'comando': command})
        print(f'INICIO {group} testes={len(names)} primeiro={names[0]}', flush=True)
        t0 = time.perf_counter()
        with (output / f'{group}.log').open('x', encoding='utf-8') as log:
            child = subprocess.run(command, cwd=root, env=env, stdout=log, stderr=subprocess.STDOUT)
        end_wall = time.time()
        events = []
        journal = folder / 'eventos.jsonl'
        if journal.exists():
            for line in journal.read_text(encoding='utf-8').splitlines():
                try:
                    events.append(json.loads(line))
                except json.JSONDecodeError:
                    # Linha incompleta por crash não comprova resultado algum.
                    break
        beginnings = {event['test']: event for event in events if event['event'] == 'test_start'}
        endings = {event['test']: event for event in events if event['event'] == 'test_end'}
        retry = []
        for name in names:
            event = endings.get(name)
            status = event.get('status', 'sem_resultado') if event else (
                'interrompido' if name in beginnings else 'nao_executado')
            duration = event['elapsed'] if event else (
                max(0, end_wall - beginnings[name]['wall']) if name in beginnings else None)
            attempts.append({'teste': name, 'resultado': status, 'duracao_s': duration,
                             'grupo': group, 'codigo_saida': child.returncode,
                             'duracao_incompleta': event is None})
            if status in ('interrompido', 'nao_executado') and len(names) > 1:
                retry.append([name])
        process = {'grupo': group, 'codigo_saida': child.returncode,
                   'duracao_s': time.perf_counter() - t0, 'testes': names,
                   'finalizacao_registrada': any(event['event'] == 'process_end' for event in events)}
        processes.append(process)
        publicar(output / f'{group}-resultado.json', process)
        summary = consolidate()
        print(f"FIM {group} exit={child.returncode} duracao={process['duracao_s']:.3f}s "
              f"consolidado={summary['contagens']}", flush=True)
        # Reduz somente grupos interrompidos. Preserva todas as tentativas anteriores.
        # Um singleton que cair não é repetido indefinidamente.
        pending[:0] = retry
    final = consolidate()
    publicar(output / 'fontes-depois.json', fontes(root))
    print(json.dumps({'sucesso': final['sucesso'], 'contagens': final['contagens'],
                      'duracao_total_s': final['duracao_total_s']}, ensure_ascii=False), flush=True)
    return 0 if final['sucesso'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
