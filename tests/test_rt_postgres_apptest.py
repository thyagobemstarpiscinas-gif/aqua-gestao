"""App real (AppTest) com backend PostgreSQL de homologação sobre o duplo do contrato.

Planilha falsa em memória, sem rede. Cobre: salvar rascunho RT, reabrir em outra
sessão, gerar o relatório mensal (dados, foto e marca), proteção contra sobrescrita
entre duas sessões e isolamento por condomínio. A validação no PostgreSQL real é
feita na homologação.
"""
import io
import json
import os
import re
import sys
import unittest
import uuid
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent / "e2e" / "fakes"))

import test_rt_relatorio_apptest as _base  # noqa: E402

AppTest = _base.AppTest
from rt_persist_falso import ContratoFalso  # noqa: E402
from utils import rt_postgres as pg  # noqa: E402
from utils.rt_contexto import ler_fotos  # noqa: E402

C1 = uuid.UUID("10000000-0000-0000-0000-000000000001")
C2 = uuid.UUID("20000000-0000-0000-0000-000000000002")
ALFA = "aqua_gestao|id:C1|09/2026"


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestRascunhoRTPostgres(unittest.TestCase):
    _entrar_e_carregar = _base.TestRelatorioRTProfissional._entrar_e_carregar
    _mes = _base.TestRelatorioRTProfissional._mes
    _carregar = _base.TestRelatorioRTProfissional._carregar
    _preparar_alfa = _base.TestRelatorioRTProfissional._preparar_alfa
    _texto = staticmethod(_base.TestRelatorioRTProfissional._texto)
    _evidencia = _base.TestRelatorioRTProfissional._evidencia

    def setUp(self):
        _base.TestRelatorioRTProfissional.setUp(self)
        self.contrato = ContratoFalso()
        cfg = {"empresas": {"aqua_gestao": {"company_id": str(C1), "principal": "rt_test_a"},
                            "bem_star": {"company_id": str(C2), "principal": "rt_test_b"}}}
        self.alfa = pg.identidade(cfg, "aqua_gestao", ALFA)
        self.contrato.conceder("rt_test_a", C1, self.alfa.condominium_id, self.alfa.period)  # só Alfa 09/2026
        caminho = self.pasta / "rt_pg_homologacao.toml"
        caminho.write_text('ambiente = "homologacao"\nconexao = "fabrica_teste"\n'
                           f'[empresas.aqua_gestao]\ncompany_id = "{C1}"\nprincipal = "rt_test_a"\n'
                           f'[empresas.bem_star]\ncompany_id = "{C2}"\nprincipal = "rt_test_b"\n', encoding="utf-8")
        p = mock.patch.dict(os.environ, {"ACQUA_RT_RASCUNHO_BACKEND": "postgres", "ACQUA_RT_PG_CONFIG": str(caminho)})
        p.start()
        self.addCleanup(p.stop)
        f = mock.patch.dict(pg.FABRICAS, {"teste": self.contrato.conectar})
        f.start()
        self.addCleanup(f.stop)

    def _revisoes(self):
        return self.contrato.revisoes.get((C1, self.alfa.condominium_id, self.alfa.period), [])

    def _salvar(self, at):
        next(b for b in at.button if b.label == "💾 Salvar rascunho").click().run()
        self.assertFalse(at.exception, at.exception)

    def test_salvar_reabrir_gerar_relatorio_e_proteger_contra_sobrescrita(self):
        from PIL import Image
        self._preparar_alfa()
        planilha = sys.modules["gspread"].PLANILHA
        # Sessão 1: dados de setembro, foto com legenda e diagnóstico; salvamento explícito.
        at1 = self._entrar_e_carregar()
        self._mes(at1, "09")
        self.assertEqual(at1.session_state["rel_origem_dados"], ALFA)
        buf = io.BytesIO()
        Image.new("RGB", (23, 29), (10, 120, 200)).save(buf, format="PNG")
        next(w for w in at1.file_uploader if w.label == "Upload de fotos do relatório").upload(
            "alfa.png", buf.getvalue(), "image/png").run()
        next(w for w in at1.text_input if w.label == "Legenda — foto 1").input("MARCADOR_LEGENDA_PG").run()
        at1.text_area(key="rel_diagnostico").input("DIAGNOSTICO_SESSAO_1").run()
        self._salvar(at1)
        self.assertTrue(any("PostgreSQL de homologação" in s.value for s in at1.success),
                        [e.value for e in at1.error] + [w.value for w in at1.warning])
        revs = self._revisoes()
        self.assertGreaterEqual(len(revs), 1)
        ultimo = revs[-1]["payload"]
        self.assertEqual(ultimo["identity"], self.alfa.envelope())
        self.assertEqual(ultimo["campos"]["rel_diagnostico"], "DIAGNOSTICO_SESSAO_1")
        self.assertEqual(json.loads(ultimo["campos"]["rel_fotos_preservadas"])["fotos"][0]["legenda"], "MARCADOR_LEGENDA_PG")
        self.assertEqual([r["revision"] for r in revs], list(range(1, len(revs) + 1)))  # revisões lineares
        self.assertNotIn("_Rascunhos_RT", planilha.abas)  # a planilha de rascunhos não é usada
        self.assertFalse(self.contrato.conflitos(C1, self.alfa.condominium_id, self.alfa.period))

        # Sessão 2: outra sessão reabre a revisão do PostgreSQL e gera o relatório mensal.
        at2 = self._entrar_e_carregar()
        self._mes(at2, "09")
        chave = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", "pg:" + self.alfa.doc)
        next(b for b in at2.button if b.key == chave).click().run()
        at2.run()
        self.assertFalse(at2.exception, at2.exception)
        ss2 = at2.session_state
        self.assertEqual(ss2["rel_diagnostico"], "DIAGNOSTICO_SESSAO_1")
        fotos = ler_fotos(ss2.filtered_state, ss2["rel_origem_dados"])
        self.assertEqual(fotos[0]["legenda"], "MARCADOR_LEGENDA_PG")
        cabeca = len(self._revisoes())
        self.assertEqual(ss2["_rt_pg_base"][self.alfa.doc], cabeca)
        next(b for b in at2.button if b.label == "🚀 Gerar relatório mensal").click().run()
        self.assertFalse(at2.exception, at2.exception)
        docs = [p for p in (self.pasta / "app" / "generated").glob("**/*Alfa*RELATORIO_RT.*") if "PREVIA" not in p.name]
        self.assertEqual({p.suffix for p in docs}, {".docx", ".pdf"})
        self._evidencia(docs)
        for p in docs:
            texto = self._texto(p)
            for esperado in ("ACQUA GESTÃO", "Condomínio Alfa Fictício", "MARCADOR_LEGENDA_PG",
                             "DIAGNOSTICO_SESSAO_1", "Piscina Adulto"):
                self.assertIn(esperado, texto, (p.name, esperado))
            self.assertNotIn("Bem Star", texto)
        from pypdf import PdfReader
        pdf = PdfReader(next(p for p in docs if p.suffix == ".pdf"))
        self.assertIn((23, 29), [im.image.size for pg_ in pdf.pages for im in pg_.images])

        # Duas pessoas editam o mesmo rascunho: a primeira grava; a segunda vira conflito e nada é sobrescrito.
        self.assertEqual(at1.session_state["_rt_pg_base"][self.alfa.doc], cabeca)
        at1.text_area(key="rel_diagnostico").input("EDICAO_SESSAO_1").run()
        self._salvar(at1)
        revs = self._revisoes()
        self.assertEqual(revs[-1]["payload"]["campos"]["rel_diagnostico"], "EDICAO_SESSAO_1")
        cabeca_nova = len(revs)
        at2.text_area(key="rel_diagnostico").input("EDICAO_SESSAO_2").run()
        self._salvar(at2)
        self.assertTrue(any("Nada foi sobrescrito" in w.value for w in at2.warning)
                        or any("divergente" in e.value for e in at2.error),
                        [w.value for w in at2.warning] + [e.value for e in at2.error])
        revs = self._revisoes()
        self.assertEqual(len(revs), cabeca_nova)
        self.assertEqual(revs[-1]["payload"]["campos"]["rel_diagnostico"], "EDICAO_SESSAO_1")
        conflitos = self.contrato.conflitos(C1, self.alfa.condominium_id, self.alfa.period)
        self.assertIn("EDICAO_SESSAO_2", {c["payload"]["campos"]["rel_diagnostico"] for c in conflitos})
        # A cópia da sessão 2 continua preservada localmente para restauração explícita.
        pasta = self.pasta / "app" / "generated" / "_rascunhos_rt_sync" / "aqua_gestao"
        estados = {json.loads(p.read_text(encoding="utf-8"))["dados"]["campos"].get("rel_diagnostico"):
                   json.loads(p.read_text(encoding="utf-8"))["estado"] for p in pasta.glob("*.json")}
        self.assertEqual(estados.get("EDICAO_SESSAO_2"), "conflito")

        # Isolamento: outro condomínio sem escopo não grava nem toca o documento de Alfa.
        at3 = self._entrar_e_carregar("Residencial Beta Fictício")
        self._mes(at3, "09")
        at3.text_area(key="rel_diagnostico").input("DIAGNOSTICO_BETA").run()
        self._salvar(at3)
        self.assertFalse(any("PostgreSQL de homologação" in s.value for s in at3.success))
        documentos = set(self.contrato.revisoes)
        self.assertEqual(documentos, {(C1, self.alfa.condominium_id, self.alfa.period)})
        self.assertNotIn("DIAGNOSTICO_BETA", json.dumps([r["payload"] for r in self._revisoes()]))
        self.assertNotIn("_Rascunhos_RT", planilha.abas)

    def test_pendencia_de_periodo_sem_escopo_nao_bloqueia_periodo_autorizado(self):
        """Achado da homologação real: o mês aberto sem escopo ficava pendente e travava setembro."""
        self._preparar_alfa()
        at = self._entrar_e_carregar()
        self.assertNotEqual(at.session_state["rel_origem_dados"], ALFA)  # mês padrão, sem escopo
        at.text_area(key="rel_diagnostico").input("DIAGNOSTICO_MES_SEM_ESCOPO").run()
        self._salvar(at)
        self.assertFalse(self.contrato.revisoes)  # negado pelo escopo; nada gravado no banco
        self._mes(at, "09")
        at.text_area(key="rel_diagnostico").input("DIAGNOSTICO_SETEMBRO").run()
        self._salvar(at)
        self.assertTrue(any("no PostgreSQL de homologação" in s.value for s in at.success),
                        [e.value for e in at.error] + [w.value for w in at.warning])
        self.assertEqual(self._revisoes()[-1]["payload"]["campos"]["rel_diagnostico"], "DIAGNOSTICO_SETEMBRO")
        # A cópia do mês sem escopo continua preservada localmente e listada como pendente.
        pasta = self.pasta / "app" / "generated" / "_rascunhos_rt_sync" / "aqua_gestao"
        pendentes = [json.loads(p.read_text(encoding="utf-8")) for p in pasta.glob("*.json")]
        self.assertTrue(any(r["estado"] == "pendente" and r["dados"]["campos"].get("rel_diagnostico") == "DIAGNOSTICO_MES_SEM_ESCOPO"
                            for r in pendentes))


if __name__ == "__main__":
    unittest.main(verbosity=2)
