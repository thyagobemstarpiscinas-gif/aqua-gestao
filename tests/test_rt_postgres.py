"""Adaptador de rascunhos RT para o contrato rt_persist, sobre o duplo em memória.

Sem rede e sem banco: a validação no PostgreSQL real é feita na homologação.
"""
import ast
import base64
import datetime
import hashlib
import json
import os
import sys
import unittest
import uuid
from pathlib import Path
from unittest import mock

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(Path(__file__).resolve().parent / "e2e" / "fakes"))

from utils import rt_postgres as pg  # noqa: E402
from rt_persist_falso import ContratoFalso  # noqa: E402

C1 = uuid.UUID("10000000-0000-0000-0000-000000000001")
C2 = uuid.UUID("20000000-0000-0000-0000-000000000002")
CFG = {"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "rt_test_a"},
                    "bem_star": {"company_id": str(C2), "principal": "rt_test_b"}}}
ALFA = "aqua_gestao|id:C1|09/2026"
SET = datetime.date(2026, 9, 1)


def rascunho(origem=ALFA, empresa="aqua_gestao", texto="DIAGNOSTICO_FICTICIO", fotos=None):
    campos = {"rel_origem_dados": origem, "rel_nome_condominio": "Condomínio Alfa Fictício",
              "rel_diagnostico": texto, "rel_analise_ph_0": "7,3"}
    if fotos is not None:
        campos["rel_fotos_preservadas"] = json.dumps({"identidade": origem, "fotos": fotos})
    return {"empresa_ativa": empresa, "rel_analises_total": 12, "salvo_em": "01/10/2026 10:00:00",
            "_salvo_iso": "2026-10-01T10:00:00", "campos": campos}


class AdaptadorTests(unittest.TestCase):
    def setUp(self):
        self.contrato = ContratoFalso()
        self.repo = pg.Repositorio(CFG, self.contrato.conectar)
        self.alfa = self.repo.identidade("aqua_gestao", ALFA)
        self.contrato.conceder("rt_test_a", C1, self.alfa.condominium_id, SET)

    def test_identidade_deterministica_e_restrita(self):
        self.assertEqual(self.alfa.company_id, C1)
        self.assertEqual(self.alfa.condominium_id, pg.condominium_id("aqua_gestao", "id:C1"))
        self.assertEqual((self.alfa.period, self.alfa.principal), (SET, "rt_test_a"))
        # Mesmo cliente em outra empresa e outro cliente na mesma empresa nunca coincidem.
        self.assertNotEqual(pg.condominium_id("bem_star", "id:C1"), self.alfa.condominium_id)
        self.assertNotEqual(pg.condominium_id("aqua_gestao", "id:C2"), self.alfa.condominium_id)
        for empresa, origem in (("bem_star", ALFA), ("aqua_gestao", ""), ("aqua_gestao", "aqua_gestao|id:C1|13/2026"),
                                ("aqua_gestao", "aqua_gestao|C1|09/2026"), ("aqua_gestao", "aqua_gestao|id:C1|9/2026")):
            with self.subTest(origem=origem), self.assertRaises(ValueError):
                self.repo.identidade(empresa, origem)

    def test_configuracao_recusa_empresa_principal_invalidos(self):
        for cfg in ({}, {"empresas": {}}, {"empresas": {"outra": CFG["empresas"]["aqua_gestao"]}},
                    {"empresas": {"aqua_gestao": {"company_id": "x", "principal": "rt_test_a"}}},
                    {"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "rt_test_a; drop"}}},
                    {"empresas": {"aqua_gestao": CFG["empresas"]["aqua_gestao"],
                                  "bem_star": dict(CFG["empresas"]["bem_star"], principal="rt_test_a")}}):
            with self.subTest(cfg=cfg), self.assertRaises((pg.RepositorioIndisponivel, ValueError)):
                pg.Repositorio(cfg, self.contrato.conectar)

    def test_salvar_reabrir_e_fotos_preservadas(self):
        png = b"\x89PNG-ficticio"
        foto = {"nome": "alfa.png", "sha256": hashlib.sha256(png).hexdigest(),
                "bytes": base64.b64encode(png).decode(), "legenda": "MARCADOR", "piscina": "Piscina Infantil"}
        r = self.repo.salvar("thyago#s#a", rascunho(fotos=[foto]), 0)
        self.assertEqual((r["status"], r["revision"]), ("saved", 1))
        dados, rev = self.repo.ultima("aqua_gestao", ALFA)
        self.assertEqual(rev, 1)
        self.assertEqual(dados["campos"]["rel_diagnostico"], "DIAGNOSTICO_FICTICIO")
        self.assertEqual(json.loads(dados["campos"]["rel_fotos_preservadas"])["fotos"], [foto])
        self.assertEqual(dados["_pg"], {"doc": self.alfa.doc, "revision": 1})
        self.assertNotIn("identity", dados)
        payload = self.contrato.revisoes[(C1, self.alfa.condominium_id, SET)][0]["payload"]
        self.assertEqual(payload["identity"], self.alfa.envelope())

    def test_base_desatualizada_vira_conflito_sem_sobrescrever(self):
        self.repo.salvar("thyago#s#a", rascunho(texto="SESSAO_A_V1"), 0)
        self.repo.salvar("thyago#s#a", rascunho(texto="SESSAO_A_V2"), 1)
        with self.assertRaises(pg.ConflitoVersao) as caught:
            self.repo.salvar("thyago#s#b", rascunho(texto="SESSAO_B_BASE_1"), 1)
        self.assertIn("divergente", str(caught.exception))
        dados, rev = self.repo.ultima("aqua_gestao", ALFA)
        self.assertEqual((rev, dados["campos"]["rel_diagnostico"]), (2, "SESSAO_A_V2"))
        conflitos = self.contrato.conflitos(C1, self.alfa.condominium_id, SET)
        self.assertEqual([c["payload"]["campos"]["rel_diagnostico"] for c in conflitos], ["SESSAO_B_BASE_1"])

    def test_reenvio_identico_e_idempotente(self):
        dados = rascunho()
        self.assertEqual(self.repo.salvar("thyago#s#a", dados, 0)["revision"], 1)
        self.assertEqual(self.repo.salvar("thyago#s#a", dados, 0)["revision"], 1)
        self.assertEqual(len(self.contrato.revisoes[(C1, self.alfa.condominium_id, SET)]), 1)

    def test_isolamento_entre_empresas_condominios_e_periodos(self):
        self.repo.salvar("thyago#s#a", rascunho(), 0)
        # Sem escopo: outro condomínio, outro período e a outra empresa são negados.
        for empresa, origem in (("aqua_gestao", "aqua_gestao|id:C9|09/2026"),
                                ("aqua_gestao", "aqua_gestao|id:C1|10/2026"),
                                ("bem_star", "bem_star|id:C1|09/2026")):
            with self.subTest(origem=origem):
                with self.assertRaises(PermissionError):
                    self.repo.salvar("thyago#s#x", rascunho(origem=origem, empresa=empresa), 0)
                with self.assertRaises(PermissionError):
                    self.repo.ultima(empresa, origem)
        # O principal da outra empresa não lê o documento de Alfa nem com a identidade correta.
        outro = pg.Repositorio({"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "rt_test_b"}}},
                               self.contrato.conectar)
        with self.assertRaises(PermissionError):
            outro.ultima("aqua_gestao", ALFA)
        self.assertEqual(len(self.contrato.revisoes), 1)

    def test_rascunho_de_outra_empresa_ou_com_identidade_forjada_e_recusado(self):
        with self.assertRaises(ValueError):
            self.repo.salvar("thyago#s#a", rascunho(origem="bem_star|id:C1|09/2026"), 0)
        forjado = dict(rascunho(), identity={"company_id": str(C2), "condominium_id": str(uuid.uuid4()), "period": "2026-09-01"})
        r = self.repo.salvar("thyago#s#a", forjado, 0)  # envelope recalculado: o metadado forjado não passa
        self.assertEqual(r["revision"], 1)
        payload = self.contrato.revisoes[(C1, self.alfa.condominium_id, SET)][0]["payload"]
        self.assertEqual(payload["identity"], self.alfa.envelope())

    def test_revisao_adulterada_nao_e_restaurada(self):
        self.repo.salvar("thyago#s#a", rascunho(), 0)
        rev = self.contrato.revisoes[(C1, self.alfa.condominium_id, SET)][0]
        rev["payload"]["campos"]["rel_origem_dados"] = "aqua_gestao|id:C2|09/2026"
        with self.assertRaises(pg.RepositorioIndisponivel):
            self.repo.ultima("aqua_gestao", ALFA)

    def test_falha_de_conexao_nao_vira_ausencia(self):
        def falha(_):
            raise OSError("sem rede")
        repo = pg.Repositorio(CFG, falha)
        with self.assertRaises(pg.RepositorioIndisponivel):
            repo.ultima("aqua_gestao", ALFA)
        with self.assertRaises(pg.RepositorioIndisponivel):
            repo.salvar("thyago#s#a", rascunho(), 0)

    def test_nao_commita_sem_sucesso_e_conflito_fica_registrado(self):
        self.repo.salvar("thyago#s#a", rascunho(texto="V1"), 0)
        with self.assertRaises(pg.ConflitoVersao):
            self.repo.salvar("thyago#s#b", rascunho(texto="CONFLITO"), 0)
        self.assertEqual(len(self.contrato.conflitos(C1, self.alfa.condominium_id, SET)), 1)
        with self.assertRaises(ValueError):
            self.repo.salvar("thyago#s#c", rascunho(), -1)


class SelecaoBackendTests(unittest.TestCase):
    """Funções de seleção extraídas do app.py real (sem executar o app)."""

    @classmethod
    def setUpClass(cls):
        arvore = ast.parse((RAIZ / "app.py").read_text(encoding="utf-8"))
        nomes = {"_rascunho_pg_ativo", "_rascunho_backend", "_rascunho_usa_sheets", "_rascunho_destino", "_rt_pg_repositorio"}
        corpo = [n for n in arvore.body if isinstance(n, ast.FunctionDef) and n.name in nomes]
        cls.ns = {"os": os, "Path": Path, "__file__": str(RAIZ / "app.py")}
        exec(compile(ast.Module(body=corpo, type_ignores=[]), "app.py", "exec"), cls.ns)

    def test_padrao_continua_planilha(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(self.ns["_rascunho_backend"](), "sheets")
            self.assertTrue(self.ns["_rascunho_usa_sheets"]())
        with mock.patch.dict(os.environ, {"CODESPACES": "true"}, clear=True):
            self.assertEqual(self.ns["_rascunho_backend"](), "local")
        with mock.patch.dict(os.environ, {"ACQUA_RT_RASCUNHO_BACKEND": "postgres"}, clear=True):
            self.assertEqual(self.ns["_rascunho_backend"](), "postgres")
            self.assertFalse(self.ns["_rascunho_usa_sheets"]())
        with mock.patch.dict(os.environ, {"ACQUA_RT_RASCUNHO_BACKEND": "sheets-qualquer"}, clear=True):
            self.assertEqual(self.ns["_rascunho_backend"](), "sheets")

    def test_textos_do_destino(self):
        d = self.ns["_rascunho_destino"]
        with mock.patch.dict(os.environ, {}, clear=True):  # produção: textos da planilha inalterados
            self.assertEqual([d(), d("a"), d("em"), d("o"), d("o_min")],
                             ["planilha", "à planilha", "na planilha", "A planilha", "a planilha"])
        with mock.patch.dict(os.environ, {"ACQUA_RT_RASCUNHO_BACKEND": "postgres"}, clear=True):
            self.assertEqual([d("a"), d("em"), d("o"), d("o_min")],
                             ["ao PostgreSQL de homologação", "no PostgreSQL de homologação",
                              "O PostgreSQL de homologação", "o PostgreSQL de homologação"])

    def test_postgres_exige_configuracao_de_homologacao(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            casos = {"ausente": None,
                     "producao": 'ambiente = "producao"\nconexao = "fabrica_teste"\n',
                     "sem_conexao": 'ambiente = "homologacao"\n',
                     "conexao_desconhecida": 'ambiente = "homologacao"\nconexao = "dsn"\n'}
            for nome, texto in casos.items():
                caminho = Path(d) / f"{nome}.toml"
                if texto is not None:
                    caminho.write_text(texto + '[empresas.aqua_gestao]\ncompany_id = "%s"\nprincipal = "rt_test_a"\n' % C1,
                                       encoding="utf-8")
                with self.subTest(nome=nome), mock.patch.dict(os.environ, {"ACQUA_RT_PG_CONFIG": str(caminho)}), \
                        mock.patch.dict(pg.FABRICAS, {"teste": ContratoFalso().conectar}):
                    with self.assertRaises(pg.RepositorioIndisponivel):
                        self.ns["_rt_pg_repositorio"]()
            ok = Path(d) / "ok.toml"
            ok.write_text('ambiente = "homologacao"\nconexao = "fabrica_teste"\n[empresas.aqua_gestao]\n'
                          f'company_id = "{C1}"\nprincipal = "rt_test_a"\n', encoding="utf-8")
            with mock.patch.dict(os.environ, {"ACQUA_RT_PG_CONFIG": str(ok)}):
                with mock.patch.dict(pg.FABRICAS, {}, clear=True):
                    with self.assertRaises(pg.RepositorioIndisponivel):
                        self.ns["_rt_pg_repositorio"]()  # sem fábrica registrada por teste
                with mock.patch.dict(pg.FABRICAS, {"teste": ContratoFalso().conectar}):
                    self.assertIsInstance(self.ns["_rt_pg_repositorio"](), pg.Repositorio)


if __name__ == "__main__":
    unittest.main(verbosity=2)
