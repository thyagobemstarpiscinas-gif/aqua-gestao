"""PDF mensal RT completo: parecer coerente com as medições, turbidez/ORP/TDS/temperatura
nos gráficos e na avaliação, e legenda/título junto do conteúdo.

App real (AppTest), planilha falsa em memória, visitas e foto fictícias, sem rede.
RT_TEST_EVIDENCE=<pasta> copia os PDFs gerados.
"""
import io
import json
import os
import shutil
import sys
import unittest
from pathlib import Path

import test_rt_relatorio_apptest as _base
from utils.cadastro_piscinas import serializar

AppTest = _base.AppTest
ALFA = "Condomínio Alfa Fictício"


CAMPOS_TOPO = ("ph", "cloro_livre", "cloro_total", "alcalinidade", "dureza", "cianurico",
               "turbidez", "orp", "tds", "temperatura")


def _visita(i, data, piscinas):
    """Mesmo formato do formulário RT real (app.py, salvar lançamento + sheets): o topo do
    payload e as colunas F–K repetem a primeira piscina; todas ficam em "piscinas"."""
    base = piscinas[0] if piscinas else {}
    payload = {"data": data, "condominio": ALFA, "operador": "Responsável Técnico Fictício", "status": "Concluída",
               "visita_rt_semanal": True, "tipo_visita": "RT semanal",
               **{k: base.get(k, "") for k in CAMPOS_TOPO}, "piscinas": piscinas}
    colunas = [base.get(k, "") for k in ("ph", "cloro_livre", "cloro_total", "alcalinidade", "dureza", "cianurico")]
    return ["", f"VA{i}", data, "C1", ALFA] + colunas + [""] * 8 + \
        ["Concluída", "Responsável Técnico Fictício", "", json.dumps(payload, ensure_ascii=False)]


def _p(nome, ph, cl, ct, turb, orp, tds, temp):
    return {"nome": nome, "ph": ph, "cloro_livre": cl, "cloro_total": ct, "alcalinidade": "100", "dureza": "200",
            "cianurico": "40", "turbidez": turb, "orp": orp, "tds": tds, "temperatura": temp}


# Quatro terças de 09/2026, duas piscinas. Fora do padrão, de propósito:
# turbidez 0,8 NTU (Adulto, 15/09) e ORP 600 mV (Infantil, 22/09).
VISITAS = [
    ("01/09/2026", [_p("Piscina Adulto", "7,4", "2,1", "2,4", "0,2", "720", "850", "27"),
                    _p("Piscina Infantil", "7,5", "2,5", "2,8", "0,3", "700", "900", "28")]),
    ("08/09/2026", [_p("Piscina Adulto", "7,5", "2,3", "2,6", "0,3", "710", "860", "26"),
                    _p("Piscina Infantil", "7,4", "2,2", "2,5", "0,2", "690", "910", "27")]),
    ("15/09/2026", [_p("Piscina Adulto", "7,6", "2,0", "2,3", "0,8", "705", "870", "28"),
                    _p("Piscina Infantil", "7,5", "2,4", "2,7", "0,4", "680", "905", "29")]),
    ("22/09/2026", [_p("Piscina Adulto", "7,4", "2,2", "2,5", "0,3", "715", "880", "27"),
                    _p("Piscina Infantil", "7,6", "2,1", "2,4", "0,3", "600", "920", "28")]),
]


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestRelatorioRTAvaliacaoCompleta(unittest.TestCase):
    setUp = _base.TestRelatorioRTProfissional.setUp
    _entrar_e_carregar = _base.TestRelatorioRTProfissional._entrar_e_carregar
    _mes = _base.TestRelatorioRTProfissional._mes
    _carregar = _base.TestRelatorioRTProfissional._carregar
    _texto = staticmethod(_base.TestRelatorioRTProfissional._texto)
    _evidencia = _base.TestRelatorioRTProfissional._evidencia

    def _preparar(self, visitas):
        import streamlit as st_real
        cli = sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes")
        meta = serializar({"adulto": "Adulto", "infantil": "Infantil"}, [])
        cli.linhas[1] = ["", "C1", ALFA, "60", "", "Síndico Alfa", "Rua A, 1", "", "Ativo",
                         "50", "10", "", "Aqua Gestão", "", "", meta]
        vis = sys.modules["gspread"].PLANILHA.worksheet("🔬 Visitas")
        vis.linhas = vis.linhas[:1] + [_visita(i, d, ps) for i, (d, ps) in enumerate(visitas, start=1)]
        st_real.cache_data.clear()

    def _guardar(self, pdf, nome):
        destino = os.environ.get("RT_TEST_EVIDENCE")
        if destino:
            Path(destino).mkdir(parents=True, exist_ok=True)
            shutil.copy2(pdf, Path(destino) / nome)

    def _gerar(self, at):
        next(b for b in at.button if b.label == "🚀 Gerar relatório mensal").click().run()
        self.assertFalse(at.exception, at.exception)
        pdfs = [p for p in (self.pasta / "app" / "generated").glob("**/*Alfa*RELATORIO_RT.pdf") if "PREVIA" not in p.name]
        self.assertEqual(len(pdfs), 1, [e.value for e in at.error])
        return pdfs[0]

    @staticmethod
    def _paginas(pdf):
        from pypdf import PdfReader
        r = PdfReader(pdf)
        return r, [pg.extract_text() or "" for pg in r.pages]

    def test_sem_medicoes_nao_avaliado(self):
        self._preparar([])
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        self.assertEqual(at.selectbox(key="rel_status_agua_select").value, "NÃO AVALIADO")
        pdf = self._gerar(at)
        self._guardar(pdf, "relatorio_rt_sem_medicoes_alfa_09_2026.pdf")
        texto = self._texto(pdf)
        self.assertIn("Não avaliado — sem dados no período", texto)
        self.assertNotIn("EM CORREÇÃO", texto)
        self.assertNotIn("sem desvios críticos", texto)
        self.assertNotIn("fase operacional de correção", texto)
        self.assertIn("Registrar as medições físico-químicas do período", texto)

    def test_pdf_completo_duas_piscinas_varias_datas(self):
        from PIL import Image
        self._preparar(VISITAS)
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        ss = at.session_state
        linhas = [(ss[f"rel_analise_piscina_{i}"], ss[f"rel_analise_data_{i}"], ss[f"rel_analise_turbidez_{i}"],
                   ss[f"rel_analise_orp_{i}"], ss[f"rel_analise_tds_{i}"], ss[f"rel_analise_temperatura_{i}"])
                  for i in range(8)]
        self.assertEqual(len({(p, d) for p, d, *_ in linhas}), 8, linhas)  # 4 datas x 2 piscinas
        self.assertIn(("Piscina Adulto", "15/09/2026", "0,8", "705", "870", "28"), linhas)
        # Foto com legenda.
        buf = io.BytesIO()
        Image.new("RGB", (640, 480), (20, 110, 170)).save(buf, format="PNG")
        next(w for w in at.file_uploader if w.label == "Upload de fotos do relatório").upload(
            "casa_maquinas.png", buf.getvalue(), "image/png").run()
        next(w for w in at.text_input if w.label == "Legenda — foto 1").input("LEGENDA_FICTICIA_CMAQ").run()
        self.assertEqual(at.selectbox(key="rel_status_agua_select").value, "NÃO CONFORME")
        pdf = self._gerar(at)
        self._guardar(pdf, "relatorio_rt_completo_alfa_09_2026_duas_piscinas.pdf")
        leitor, paginas = self._paginas(pdf)
        texto = " ".join(paginas)
        # Medições nas tabelas por piscina.
        for esperado in ("Piscina: Piscina Adulto", "Piscina: Piscina Infantil", "15/09/2026", "22/09/2026", "0,8", "600", "870"):
            self.assertIn(esperado, texto, esperado)
        # Parecer coerente: turbidez fora do limite, ORP como atenção, TDS/temperatura informativos.
        self.assertIn("Status geral da água: NÃO CONFORME", texto)
        self.assertIn("turbidez=0,8 fora do limite < 0,5 (NBR 10818:2026) — 15/09/2026, Piscina Adulto", texto)
        self.assertIn("0,30", texto)  # cloro combinado com vírgula, como as demais colunas
        self.assertIn("ORP=600 mV abaixo da referência", texto)
        self.assertIn("TDS e temperatura registrados para acompanhamento", texto)
        self.assertIn("Verificar a filtração", texto)
        for proibido in ("EM CORREÇÃO", "sem desvios críticos", "Não avaliado"):
            self.assertNotIn(proibido, texto, proibido)
        # Um gráfico por piscina, com 10 painéis (pH, CRL, CT, alc., dureza, CYA, turbidez, ORP, TDS, temp.).
        graficos = [im.image.size for pg in leitor.pages for im in pg.images if im.image.size[0] == 1600]
        self.assertEqual(graficos, [(1600, 150 + 5 * 340 + 60)] * 2)
        # Título da turbidez com o texto; legenda com a foto (mesma página).
        pg_turb = next(i for i, t in enumerate(paginas) if "IMPORTÂNCIA DA ANÁLISE DE TURBIDEZ" in t)
        self.assertIn("A turbidez é medida", paginas[pg_turb])
        pg_leg = next(i for i, t in enumerate(paginas) if "LEGENDA_FICTICIA_CMAQ" in t)
        self.assertIn((640, 480), [im.image.size for im in leitor.pages[pg_leg].images])
        pg_41 = next(i for i, t in enumerate(paginas) if "4.1 Registro de Análises" in t)
        self.assertIn("Piscina: Piscina Adulto", paginas[pg_41])  # título não fica sozinho no fim da página
        pg_5 = next(i for i, t in enumerate(paginas) if "5. PARECER TÉCNICO" in t)
        self.assertIn("Status geral da água", paginas[pg_5])
        pg_graf = next(i for i, t in enumerate(paginas) if "GRÁFICOS TÉCNICOS POR PISCINA" in t)
        self.assertIn(1600, [im.image.size[0] for im in leitor.pages[pg_graf].images])


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestPDFModoRTCampo(unittest.TestCase):
    """PDF do modo RT de campo (build_monthly_rt_pdf_bytes): todas as piscinas de cada visita,
    parecer pelas leituras e foto enviada nesta tela com a legenda."""
    setUp = TestRelatorioRTAvaliacaoCompleta.setUp
    _preparar = TestRelatorioRTAvaliacaoCompleta._preparar
    _guardar = TestRelatorioRTAvaliacaoCompleta._guardar
    _paginas = staticmethod(TestRelatorioRTAvaliacaoCompleta._paginas)

    def _entrar_rt(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["rt"] = {"pin": "rt-ficticio"}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        next(b for b in at.button if b.label == "📱 Acessar como Operador").click().run()
        at.text_input(key="op_pin_input").input("rt-ficticio")
        next(b for b in at.button if b.label == "Entrar").click().run()
        at.selectbox(key="rt_op_sel_cond").select(ALFA).run()
        at.selectbox(key="rt_periodo_mes").select("09").run()
        at.selectbox(key="rt_periodo_ano").select("2026").run()
        self.assertFalse(at.exception, at.exception)
        return at

    def _gerar_campo(self, at, nome):
        """O PDF do campo só vai para o botão de download: captura os bytes dele."""
        from unittest import mock
        import streamlit as st_mod
        original = getattr(st_mod.download_button, "_original_aqua", st_mod.download_button)
        capturado = {}

        def captura(label, data=None, *a, **k):  # função simples: o app procura _original_aqua
            capturado.setdefault(label, data)
            return original(label, data, *a, **k)
        with mock.patch("streamlit.download_button", new=captura):
            next(b for b in at.button if b.label == "📄 Gerar Relatório RT em PDF").click().run()
        self.assertFalse(at.exception, at.exception)
        dados = capturado.get("⬇️ Baixar Relatório RT em PDF")
        self.assertTrue(dados, [e.value for e in at.error] + [w.value for w in at.warning])
        pdf = self.pasta / nome
        pdf.write_bytes(dados)
        self._guardar(pdf, nome)
        return pdf

    def test_campo_duas_piscinas_turbidez_e_foto_com_legenda(self):
        from PIL import Image
        self._preparar(VISITAS)
        at = self._entrar_rt()
        buf = io.BytesIO()
        Image.new("RGB", (640, 480), (30, 140, 90)).save(buf, format="PNG")
        next(w for w in at.file_uploader if w.label == "Fotos do relatório RT").upload(
            "campo.png", buf.getvalue(), "image/png").run()
        at.text_input(key="rel_foto_legenda_cmaq").input("LEGENDA_CAMPO_CASA_MAQUINAS").run()
        pdf = self._gerar_campo(at, "relatorio_rt_modo_campo_alfa_09_2026_corrigido.pdf")
        leitor, paginas = self._paginas(pdf)
        texto = " ".join(paginas)
        # Oito linhas (4 datas x 2 piscinas), identificadas, com as medições de cada piscina.
        for esperado in ("Piscina Adulto", "Piscina Infantil", "0,8", "600", "920", "29"):
            self.assertIn(esperado, texto, esperado)
        self.assertNotIn("—\n—\n—", texto)
        self.assertIn("ATENÇÃO — HÁ DESVIOS NO PERÍODO", texto)
        self.assertIn("15/09/2026, Piscina Adulto — verificar: turbidez", texto)
        self.assertNotIn("SEM DESVIOS OPERACIONAIS", texto)
        self.assertIn("Visitas técnicas do RT 4", " ".join(texto.split()))
        self.assertIn("Leituras avaliadas 8", " ".join(texto.split()))
        # Foto enviada nesta tela, com a legenda digitada, na mesma página.
        pg = next(i for i, t in enumerate(paginas) if "LEGENDA_CAMPO_CASA_MAQUINAS" in t)
        self.assertIn((640, 480), [im.image.size for im in leitor.pages[pg].images])

    def test_campo_sem_leituras_nao_avaliado(self):
        vazias = [(d, [{"nome": "Piscina Adulto"}, {"nome": "Piscina Infantil"}]) for d, _ in VISITAS]
        self._preparar(vazias)
        at = self._entrar_rt()
        pdf = self._gerar_campo(at, "relatorio_rt_modo_campo_sem_leituras.pdf")
        texto = " ".join(self._paginas(pdf)[1])
        self.assertIn("Não avaliado — sem dados no período", texto)
        for proibido in ("SEM DESVIOS OPERACIONAIS", "Não foram identificados desvios", "EM CORREÇÃO", "sem desvios críticos"):
            self.assertNotIn(proibido, texto, proibido)


if __name__ == "__main__":
    unittest.main(verbosity=2)
