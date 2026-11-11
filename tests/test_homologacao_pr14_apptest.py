"""Regressões da homologação do PR #14 no app real (AppTest, planilha falsa, sem rede).

- Rascunho RT de outra sessão: não é listado com outro condomínio/período na tela;
  com a identidade dele carregada, a restauração traz dados, período e seletor de
  cliente do rascunho, sem mistura, e a edição do outro condomínio fica guardada.
- "Preencher parecer automático": o status calculado aparece no seletor
  "Status geral da água" na mesma execução e não é revertido depois.
- Relatório mensal Bem Star do condomínio fictício Épsilon: a visita V00005
  (cópia local sem ID + Sheets) aparece uma vez, a piscina é a cadastrada
  ("Piscina Adulto"), sem "Piscina adulta", e a geração não lança erro csr_sel.
"""
import json
import os
import re
import sys
import unittest
from unittest.mock import patch

import fitz

from test_relatorio_rt_campos import AppTest, PIN_FICTICIO, _AppReal, TestRelatorioRTCamposNoAppReal as _Aux
from utils.formatacao import slugify_nome

EPSILON = "Condomínio Épsilon Fictício"


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class RascunhoRTRestauracao(_AppReal):
    _entrar_e_carregar = _Aux._entrar_e_carregar

    def setUp(self):
        super().setUp()
        p = patch.dict(os.environ, {"CODESPACES": ""})  # rascunho na planilha (falsa), como no Cloud
        p.start()
        self.addCleanup(p.stop)

    def _rascunho_alfa_setembro(self):
        a = self._entrar_e_carregar("Condomínio Alfa Fictício")
        a.text_input(key="rel_mes_referencia").input("09").run()
        a.text_area(key="rel_diagnostico").input("DIAGNOSTICO_ALFA_SETEMBRO").run()
        self.assertFalse(a.exception, a.exception)
        self.assertEqual(a.session_state["rel_origem_dados"], "aqua_gestao|id:C1|09/2026")
        return a.session_state["_rt_rasc_chave"]

    def test_restaura_outro_condominio_e_periodo_sem_misturar(self):
        chave = self._rascunho_alfa_setembro()
        b = self._entrar_e_carregar("Residencial Beta Fictício")  # mês corrente, outro cliente
        b.text_area(key="rel_diagnostico").input("DIAGNOSTICO_BETA_ATUAL").run()
        prefixo = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", chave)
        # Rascunho de Alfa 09/2026 não é listado com Beta; aparece ao carregar Alfa 09/2026.
        self.assertFalse(any((x.key or "").startswith(prefixo) for x in b.button))
        b.selectbox(key="sel_cliente_rel").select("Condomínio Alfa Fictício").run()
        next(x for x in b.button if x.label == "⬇️ Carregar").click().run()
        b.text_input(key="rel_mes_referencia").input("09").run()
        next(x for x in b.button if (x.key or "").startswith(prefixo)).click().run()
        self.assertFalse(b.exception, b.exception)
        self.assertNotIn("_relatorio_rt_restore_erro", b.session_state)
        b.run()
        ss = b.session_state
        self.assertEqual(ss["rel_diagnostico"], "DIAGNOSTICO_ALFA_SETEMBRO")
        self.assertEqual((ss["rel_nome_condominio"], ss["rel_mes_referencia"]), ("Condomínio Alfa Fictício", "09"))
        self.assertEqual(ss["rel_origem_dados"], "aqua_gestao|id:C1|09/2026")
        # O seletor de cliente acompanha o rascunho restaurado.
        self.assertEqual(ss["sel_cliente_rel"], "Condomínio Alfa Fictício")
        # A edição de Beta não se perdeu: foi guardada como outra versão.
        sheet = sys.modules["gspread"].PLANILHA.worksheet("_Rascunhos_RT")
        self.assertTrue(any("DIAGNOSTICO_BETA_ATUAL" in "".join(row) for row in sheet.linhas))
        self.assertFalse([e.value for e in b.error])

    def test_parecer_automatico_atualiza_status_no_seletor(self):
        at = self._entrar_e_carregar("Condomínio Alfa Fictício")
        at.text_input(key="rel_mes_referencia").input("09").run()
        at.text_input(key="rel_analise_ph_0").input("6,5").run()  # fora do limite NBR 10818
        self.assertEqual(at.session_state["rel_status_agua_select"], "CONFORME")
        next(b for b in at.button if b.label == "Preencher parecer automático").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(at.session_state["rel_status_agua"], "NÃO CONFORME")
        self.assertEqual(at.selectbox(key="rel_status_agua_select").value, "NÃO CONFORME")
        at.run()
        self.assertEqual(at.session_state["rel_status_agua"], "NÃO CONFORME")
        # Escolha manual continua possível.
        at.selectbox(key="rel_status_agua_select").select("EM CORREÇÃO").run()
        self.assertEqual(at.session_state["rel_status_agua"], "EM CORREÇÃO")


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class RelatorioMensalBemStarEpsilon(_AppReal):
    def setUp(self):
        super().setUp()
        planilha = sys.modules["gspread"].PLANILHA
        clientes = planilha.worksheet("👥 Clientes")
        clientes.linhas.append(["", "C3", EPSILON, "45", "", "Síndico Épsilon", "Rua E, 5", "", "Ativo",
                                "45", "", "", "Bem Star Piscinas", "", ""])
        clientes.linhas.append(["", "C4", "Condomínio Gama Fictício", "60", "", "Síndico Gama", "Rua C, 3", "",
                                "Ativo", "60", "", "", "Bem Star Piscinas", "", ""])
        visitas = planilha.worksheet("🔬 Visitas")
        self.local = {"data": "27/10/2026", "operador": "OPERADOR BEM STAR FICTÍCIO", "condominio": EPSILON,
                      "salvo_em": "27/10/2026 10:15:00", "ph": "7,4", "cloro_livre": "2,0",
                      "piscinas": [{"nome": "Piscina Adulto", "ph": "7,4", "cloro_livre": "2,0", "cloro_total": "2,2",
                                    "alcalinidade": "100", "dureza": "220", "cianurico": "40"}],
                      "fotos_antes_b64": [], "dosagens": [], "sheets_sync": "ok"}
        payload = {k: v for k, v in self.local.items() if not k.endswith("_b64") and k != "sheets_sync"}
        payload.update(id_visita="V00005", status="Concluída")
        visitas.linhas.append(["", "V00005", "27/10/2026", "C3", EPSILON, "7,4", "2,0", "2,2", "100", "220", "40",
                               "", "", "", "", "", "", "", "", "Concluída", "OPERADOR BEM STAR FICTÍCIO", "",
                               json.dumps(payload, ensure_ascii=False)])
        # Cópia local gravada pelo modo campo antes de o Sheets gerar o ID (como V00005).
        pasta = self.pasta / "app" / "generated" / slugify_nome(EPSILON)
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "dados_condominio.json").write_text(json.dumps(
            {"nome_condominio": EPSILON, "lancamentos_campo": [self.local]}, ensure_ascii=False), encoding="utf-8")
        # Pasta local do Gama sem empresa (criada por visita/rascunho).
        pasta_gama = self.pasta / "app" / "generated" / slugify_nome("Condomínio Gama Fictício")
        pasta_gama.mkdir(parents=True, exist_ok=True)
        (pasta_gama / "dados_condominio.json").write_text(json.dumps(
            {"nome_condominio": "Condomínio Gama Fictício", "lancamentos_campo": []}, ensure_ascii=False),
            encoding="utf-8")

    def _entrar(self, empresa_botao):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao(empresa_botao).click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        self.assertFalse(at.exception, at.exception)
        return at

    def test_gama_nao_aparece_na_aqua(self):
        at = self._entrar("🔵 Aqua Gestão")
        self.assertNotIn("Condomínio Gama Fictício", at.selectbox(key="sel_cliente_rel").options)
        self.assertIn("Condomínio Alfa Fictício", at.selectbox(key="sel_cliente_rel").options)

    def test_epsilon_mensal_sem_duplicidade_sem_piscina_extra_e_sem_erro_csr_sel(self):
        at = self._entrar("⭐ Bem Star Piscinas")
        self.assertIn(EPSILON, at.selectbox(key="csr_sel_relatorio").options)
        at.selectbox(key="csr_sel_relatorio").select(EPSILON)
        at.text_input(key="csr_mes_rel").input("10")
        at.text_input(key="csr_ano_rel").input("2026").run()
        self.assertFalse(at.exception, at.exception)
        resumo = [m.value for m in at.markdown if "lançamento(s) encontrado(s)" in m.value]
        self.assertTrue(resumo and "<strong>📱 1 lançamento(s)" in resumo[0], resumo)
        piscinas = at.multiselect(key="bs_mensal_piscinas")
        self.assertEqual(piscinas.value, ["Piscina Adulto"])
        self.assertNotIn("Piscina adulta", piscinas.options)
        next(b for b in at.button if b.label == "📄 Gerar relatório Bem Star (PDF)").click().run()
        self.assertFalse(at.exception, at.exception)
        erros = [e.value for e in at.error]
        self.assertFalse(erros, erros)
        self.assertFalse([c.value for c in at.code if "csr_sel" in c.value])
        pdfs = sorted((self.pasta / "app" / "generated" / slugify_nome(EPSILON)).glob("*RELATORIO_MENSAL_BS*.pdf"))
        self.assertTrue(pdfs)
        with fitz.open(str(pdfs[-1])) as doc:
            texto = "\n".join(p.get_text() for p in doc)
        self.assertIn("Piscina Adulto", texto)
        self.assertNotIn("Piscina adulta", texto)
        # Os dados da homologação permanecem: cópia local e linha V00005 intactas.
        dados = json.loads((self.pasta / "app" / "generated" / slugify_nome(EPSILON) / "dados_condominio.json")
                           .read_text(encoding="utf-8"))
        self.assertEqual(len(dados["lancamentos_campo"]), 1)
        self.assertEqual(sum(1 for r in sys.modules["gspread"].PLANILHA.worksheet("🔬 Visitas").linhas
                             if len(r) > 1 and r[1] == "V00005"), 1)


del _Aux
