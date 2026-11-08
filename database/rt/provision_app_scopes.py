"""Escopos de homologação para os rascunhos RT do app (dados fictícios).

Para cada origem "empresa|id:C00N|MM/AAAA" cria, se faltar, o condomínio derivado
(utils.rt_postgres.condominium_id) na empresa fictícia configurada e o escopo do
principal da empresa. Usa o administrador do manifesto (passfile, verify-full,
marcador). Padrão: simulação somente leitura; grava apenas com --aplicar.
Nunca cria empresa, papel, membership ou privilégio; nunca apaga nem altera linhas.

    python database/rt/provision_app_scopes.py --manifesto M.toml --config APP.toml \
        --origem "aqua_gestao|id:C001|09/2026" --origem "bem_star|id:C003|09/2026" [--aplicar]
"""
import argparse
import json
from pathlib import Path
import re
import sys
import tomllib

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import supabase_runner as runner  # noqa: E402
from utils import rt_postgres  # noqa: E402

CLIENTE_FICTICIO = re.compile(r"^id:C\d{3}$")


def plano(cfg, origens):
    itens = []
    for origem in origens:
        empresa = str(origem).split("|", 1)[0]
        ident = rt_postgres.identidade(cfg, empresa, origem)
        if not CLIENTE_FICTICIO.match(ident.cliente):
            raise runner.Refused("somente clientes fictícios C00N da homologação")
        if ident.principal not in runner.ROLES:
            raise runner.Refused("principal fora dos LOGINs de teste")
        itens.append(ident)
    if len({i.doc for i in itens}) != len(itens):
        raise runner.Refused("origem repetida")
    return itens


def aplicar(c, itens, gravar):
    resumo = []
    empresas = {r[0] for r in c.execute("SELECT company_id FROM rt_persist.companies").fetchall()}
    for i in itens:
        if i.company_id not in empresas:
            raise runner.Refused("empresa fictícia ausente; este script não cria empresas")
        tem_condo = c.execute("SELECT 1 FROM rt_persist.condominiums WHERE company_id=%s AND condominium_id=%s",
                              (i.company_id, i.condominium_id)).fetchone() is not None
        tem_escopo = c.execute("SELECT 1 FROM rt_persist.access_scopes WHERE principal=%s AND company_id=%s "
                               "AND condominium_id=%s AND period=%s",
                               (i.principal, i.company_id, i.condominium_id, i.period)).fetchone() is not None
        if gravar and not tem_condo:
            c.execute("INSERT INTO rt_persist.condominiums VALUES (%s,%s,%s)",
                      (i.company_id, i.condominium_id, "FICTICIO APP " + i.cliente[3:]))
        if gravar and not tem_escopo:
            c.execute("INSERT INTO rt_persist.access_scopes VALUES (%s,%s,%s,%s)",
                      (i.principal, i.company_id, i.condominium_id, i.period))
        resumo.append({"origem": f"{i.empresa}|{i.cliente}|{i.period:%m/%Y}", "principal": i.principal,
                       "condominio_existia": tem_condo, "escopo_existia": tem_escopo})
    return resumo


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--manifesto", required=True)
    p.add_argument("--config", required=True)
    p.add_argument("--origem", action="append", required=True)
    p.add_argument("--aplicar", action="store_true")
    a = p.parse_args()
    try:
        with open(a.config, "rb") as f:
            cfg = rt_postgres.validar_config(tomllib.load(f))
        itens = plano(cfg, a.origem)
        runner.load_config(a.manifesto)
        with runner.connect(read_only=not a.aplicar) as c:
            resumo = aplicar(c, itens, a.aplicar)
            if a.aplicar:
                c.commit()
            else:
                c.rollback()
        print(("ESCOPOS_APLICADOS " if a.aplicar else "SIMULACAO ") + json.dumps(resumo, ensure_ascii=False))
        return 0
    except Exception as e:  # noqa: BLE001
        print("PROVISIONAMENTO_RECUSADO tipo=" + type(e).__name__
              + (" msg=" + str(e) if isinstance(e, (runner.Refused, ValueError)) else ""), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
