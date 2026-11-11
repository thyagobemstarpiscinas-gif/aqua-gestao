"""Bloqueios finais do PR #14 no app real (AppTest, planilha e Drive falsos, sem rede).

1. Sequência da homologação: restaurar rascunho de Alfa, trocar novembro → outubro
   com a planilha falhando, carregar Beta e restaurar o rascunho dele (rascunho de outro
   condomínio/período não é listado). Antes: "Troca bloqueada…",
   o mês voltava para novembro e a restauração era recusada.
2. Relatório mensal Bem Star do Épsilon com a V00005 lida só do Sheets (servidor
   sem pasta local) e a foto no Drive: antes o PDF saía sem registro fotográfico.
"""
import io
import json
import os
import re
import sys
import unittest
from unittest.mock import patch

import fitz
from PIL import Image

from test_relatorio_rt_campos import AppTest, PIN_FICTICIO, _AppReal, TestRelatorioRTCamposNoAppReal as _Aux
from utils.formatacao import slugify_nome

EPSILON = "Condomínio Épsilon Fictício"
FOTO_ID = "DRIVE_FOTO_EPSILON_V00005"


def _foto() -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", (640, 480), (200, 30, 30)).save(buf, format="JPEG", quality=90)
    return buf.getvalue()


class _DriveFalso:
    """Somente leitura (get_media) dos arquivos fictícios informados."""

    def __init__(self, arquivos):
        self.arquivos = arquivos

    def files(self):
        return self

    def get_media(self, fileId):
        arquivos = self.arquivos

        class _Exec:
            def execute(self):
                if fileId not in arquivos:
                    raise PermissionError("arquivo fora do alcance")
                return arquivos[fileId]
        return _Exec()


def _restaurar(at, chave):
    prefixo = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", chave)
    next(b for b in at.button if (b.key or "").startswith(prefixo)).click().run()


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class SequenciaRestauracaoRT(_AppReal):
    _entrar_e_carregar = _Aux._entrar_e_carregar

    def setUp(self):
        super().setUp()
        p = patch.dict(os.environ, {"CODESPACES": ""})  # rascunhos na planilha (falsa), como no Cloud
        p.start()
        self.addCleanup(p.stop)

    def test_alfa_outubro_beta_com_planilha_instavel(self):
        beta = self._entrar_e_carregar("Residencial Beta Fictício")
        beta.text_area(key="rel_diagnostico").input("DIAG_BETA").run()
        chave_beta = beta.session_state["_rt_rasc_chave"]
        alfa = self._entrar_e_carregar("Condomínio Alfa Fictício")
        alfa.text_area(key="rel_diagnostico").input("DIAG_ALFA").run()
        chave_alfa = alfa.session_state["_rt_rasc_chave"]

        at = self._entrar_e_carregar("Condomínio Alfa Fictício")
        _restaurar(at, chave_alfa)
        self.assertEqual(at.session_state["rel_diagnostico"], "DIAG_ALFA")
        aba = sys.modules["gspread"].PLANILHA.worksheet("_Rascunhos_RT")
        with patch.object(type(aba), "append_row", side_effect=ConnectionError("COTA_FICTICIA")), \
                patch.object(type(aba), "batch_update", side_effect=ConnectionError("COTA_FICTICIA")):
            at.text_area(key="rel_diagnostico").input("ALFA_NOVEMBRO_EDITADO").run()
            at.text_input(key="rel_mes_referencia").input("10").run()
            self.assertFalse(at.exception, at.exception)
            erros = [e.value for e in at.error]
            self.assertFalse([e for e in erros if "Troca bloqueada" in e], erros)
            self.assertEqual(at.session_state["rel_mes_referencia"], "10")  # não volta para novembro
            self.assertEqual(at.session_state["rel_origem_dados"], "aqua_gestao|id:C1|10/2026")
            self.assertNotEqual(at.session_state["rel_diagnostico"], "ALFA_NOVEMBRO_EDITADO")  # sem mistura

            # Rascunho de Beta/novembro só é listado com Beta/novembro carregado.
            prefixo_beta = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", chave_beta)
            self.assertFalse(any((b.key or "").startswith(prefixo_beta) for b in at.button))
            at.selectbox(key="sel_cliente_rel").select("Residencial Beta Fictício").run()
            next(b for b in at.button if b.label == "⬇️ Carregar").click().run()
            at.text_input(key="rel_mes_referencia").input("11").run()
            _restaurar(at, chave_beta)
            self.assertFalse(at.exception, at.exception)
            self.assertNotIn("_relatorio_rt_restore_erro", at.session_state)
            at.run()
            ss = at.session_state
            self.assertEqual((ss["rel_diagnostico"], ss["rel_nome_condominio"]), ("DIAG_BETA", "Residencial Beta Fictício"))
            self.assertEqual(ss["rel_origem_dados"], "aqua_gestao|id:C2|11/2026")

        # Planilha de volta: a edição de Alfa/novembro chega sem nova ação.
        at.session_state["_rt_rasc_proxima_tentativa"] = 0
        at.run()
        self.assertTrue(any("ALFA_NOVEMBRO_EDITADO" in "".join(row) for row in aba.linhas))
        self.assertEqual(at.session_state["rel_diagnostico"], "DIAG_BETA")


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class FotoMensalBemStarEpsilon(_AppReal):
    def setUp(self):
        super().setUp()
        planilha = sys.modules["gspread"].PLANILHA
        planilha.worksheet("👥 Clientes").linhas.append(
            ["", "C3", EPSILON, "45", "", "Síndico Épsilon", "Rua E, 5", "", "Ativo", "45", "", "",
             "Bem Star Piscinas", "", ""])
        payload = {"id_visita": "V00005", "data": "27/10/2026", "condominio": EPSILON, "status": "Concluída",
                   "operador": "OPERADOR BEM STAR FICTÍCIO", "salvo_em": "27/10/2026 10:15:00",
                   "piscinas": [{"nome": "Piscina Adulto", "ph": "7,4", "cloro_livre": "2,0", "cloro_total": "2,2",
                                 "alcalinidade": "100", "dureza": "220", "cianurico": "40"}],
                   "fotos_antes": ["antes_20261027_101500_1_epsilon.jpg"], "fotos_antes_ids": [FOTO_ID],
                   "dosagens": []}
        self.visitas = planilha.worksheet("🔬 Visitas")
        self.visitas.linhas.append(["", "V00005", "27/10/2026", "C3", EPSILON, "7,4", "2,0", "2,2", "100", "220",
                                    "40", "", "", "", "", "", "", "", "", "Concluída",
                                    "OPERADOR BEM STAR FICTÍCIO", "", json.dumps(payload, ensure_ascii=False)])
        self.drive = _DriveFalso({FOTO_ID: _foto()})
        p = patch("googleapiclient.discovery.build", lambda *a, **k: self.drive)
        p.start()
        self.addCleanup(p.stop)

    def _entrar_bem_star(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao("⭐ Bem Star Piscinas").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        at.selectbox(key="csr_sel_relatorio").select(EPSILON)
        at.text_input(key="csr_mes_rel").input("10")
        at.text_input(key="csr_ano_rel").input("2026").run()
        self.assertFalse(at.exception, at.exception)
        return at

    def test_pdf_mensal_inclui_foto_da_v00005_vinda_do_drive(self):
        pasta = self.pasta / "app" / "generated" / slugify_nome(EPSILON)
        self.assertFalse((pasta / "fotos_campo").exists())  # servidor sem cópia local
        at = self._entrar_bem_star()
        next(b for b in at.button if b.label == "📄 Gerar relatório Bem Star (PDF)").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertFalse([e.value for e in at.error])
        self.assertFalse([w.value for w in at.warning if "Registro fotográfico incompleto" in w.value])
        pdf = sorted(pasta.glob("*RELATORIO_MENSAL_BS*.pdf"))[-1]
        with fitz.open(str(pdf)) as doc:
            texto = "\n".join(p.get_text() for p in doc)
            tamanhos = [(i["width"], i["height"]) for page in doc
                        for i in (doc.extract_image(x[0]) for x in page.get_images())]
        self.assertIn("Registro fotográfico", texto)
        self.assertIn((640, 480), tamanhos)
        # V00005 não foi duplicada nem alterada.
        linhas = [r for r in self.visitas.linhas if len(r) > 1 and r[1] == "V00005"]
        self.assertEqual(len(linhas), 1)
        self.assertEqual(json.loads(linhas[0][22])["fotos_antes_ids"], [FOTO_ID])

    def test_foto_indisponivel_no_drive_avisa_na_tela(self):
        self.drive.arquivos.clear()
        at = self._entrar_bem_star()
        next(b for b in at.button if b.label == "📄 Gerar relatório Bem Star (PDF)").click().run()
        self.assertFalse(at.exception, at.exception)
        avisos = [w.value for w in at.warning if "Registro fotográfico incompleto" in w.value]
        self.assertTrue(avisos and "V00005" in avisos[0], avisos)


del _Aux
