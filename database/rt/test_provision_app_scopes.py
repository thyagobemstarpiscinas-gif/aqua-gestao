"""Regressões offline do provisionamento de escopos do app: sem banco, sem rede."""
import unittest
import uuid
from unittest.mock import MagicMock

import provision_app_scopes as prov
import supabase_runner as runner
from utils import rt_postgres

C1 = uuid.UUID("10000000-0000-0000-0000-000000000001")
C2 = uuid.UUID("20000000-0000-0000-0000-000000000002")
CFG = {"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "rt_test_a"},
                    "bem_star": {"company_id": str(C2), "principal": "rt_test_b"}}}


def conexao(empresas=(C1, C2), condos=(), escopos=()):
    c = MagicMock()
    def execute(sql, params=None):
        r = MagicMock()
        if "FROM rt_persist.companies" in sql:
            r.fetchall.return_value = [(e,) for e in empresas]
        elif "FROM rt_persist.condominiums" in sql:
            r.fetchone.return_value = (1,) if tuple(params) in condos else None
        elif "FROM rt_persist.access_scopes" in sql:
            r.fetchone.return_value = (1,) if tuple(params) in escopos else None
        return r
    c.execute.side_effect = execute
    return c


def inserts(c):
    return [call.args for call in c.execute.call_args_list if call.args[0].startswith("INSERT")]


class ProvisionTests(unittest.TestCase):
    def test_plano_somente_clientes_ficticios_e_principais_de_teste(self):
        itens = prov.plano(CFG, ["aqua_gestao|id:C001|09/2026", "bem_star|id:C003|09/2026"])
        self.assertEqual([(i.principal, i.company_id) for i in itens], [("rt_test_a", C1), ("rt_test_b", C2)])
        self.assertEqual(itens[0].condominium_id, rt_postgres.condominium_id("aqua_gestao", "id:C001"))
        for origens in (["aqua_gestao|id:12345|09/2026"], ["aqua_gestao|nome:alfa|09/2026"],
                        ["aqua_gestao|id:C001|09/2026"] * 2):
            with self.subTest(origens=origens), self.assertRaises(runner.Refused):
                prov.plano(CFG, origens)
        with self.assertRaises(ValueError):
            prov.plano(CFG, ["aqua_gestao|id:C001|13/2026"])
        outro = {"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "postgres"}}}
        with self.assertRaises(runner.Refused):
            prov.plano(outro, ["aqua_gestao|id:C001|09/2026"])

    def test_simulacao_nao_grava(self):
        c = conexao()
        resumo = prov.aplicar(c, prov.plano(CFG, ["aqua_gestao|id:C001|09/2026"]), gravar=False)
        self.assertEqual(inserts(c), [])
        self.assertEqual(resumo[0]["escopo_existia"], False)

    def test_aplica_somente_o_que_falta_sem_update_ou_delete(self):
        i = prov.plano(CFG, ["aqua_gestao|id:C001|09/2026"])[0]
        c = conexao(condos={(i.company_id, i.condominium_id)})
        prov.aplicar(c, [i], gravar=True)
        feitos = inserts(c)
        self.assertEqual(len(feitos), 1)
        self.assertIn("access_scopes", feitos[0][0])
        self.assertEqual(feitos[0][1], (i.principal, i.company_id, i.condominium_id, i.period))
        self.assertFalse(any(k in call.args[0].upper() for call in c.execute.call_args_list
                             for k in ("UPDATE", "DELETE", "TRUNCATE", "DROP", "CREATE", "GRANT")))
        c2 = conexao(condos={(i.company_id, i.condominium_id)},
                     escopos={(i.principal, i.company_id, i.condominium_id, i.period)})
        prov.aplicar(c2, [i], gravar=True)
        self.assertEqual(inserts(c2), [])

    def test_empresa_ausente_e_recusada_sem_criar(self):
        c = conexao(empresas=(C2,))
        with self.assertRaises(runner.Refused):
            prov.aplicar(c, prov.plano(CFG, ["aqua_gestao|id:C001|09/2026"]), gravar=True)
        self.assertEqual(inserts(c), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
