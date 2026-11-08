"""App real (AppTest) com planilha FALSA em memória: piscinas do cadastro, análises por
piscina, fotos com legenda, ART sem vigência, capa ACQUA GESTÃO e consolidado.

Nenhum acesso a Sheets/Drive/rede. RT_TEST_EVIDENCE=<pasta> copia os documentos gerados.
"""
import io
import json
import os
import shutil
import sys
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from test_relatorio_rt_campos import TestRelatorioRTCamposNoAppReal as Base, AppTest
from utils.cadastro_piscinas import serializar
from utils.rt_contexto import ler_fotos


def _visita(i, data, piscinas):
    payload = {"data": data, "condominio": "Condomínio Alfa Fictício", "operador": "Responsável Técnico Fictício",
               "status": "Concluída", "piscinas": piscinas,
               "visita_rt_semanal": True, "tipo_visita": "RT semanal"}  # visita técnica RT
    return ["", f"VP{i}", data, "C1", "Condomínio Alfa Fictício", "", "", "", "", "", ""] + [""] * 8 + \
        ["Concluída", "Responsável Técnico Fictício", "", json.dumps(payload, ensure_ascii=False)]


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestRelatorioRTProfissional(unittest.TestCase):
    setUp = Base.setUp
    _entrar_e_carregar = Base._entrar_e_carregar
    _mes = Base._mes
    _carregar = Base._carregar
    _cadastrar = Base._cadastrar

    def _preparar_alfa(self):
        import streamlit as st_real
        cli = sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes")
        meta = serializar({"adulto": "Adulto", "infantil": "Infantil", "extra1": "SPA"},
                          [{"nome": "SPA Cobertura", "vol": 4.5, "tipo": "SPA"}])
        cli.linhas[1] = ["", "C1", "Condomínio Alfa Fictício", "64.5", "", "Síndico Alfa", "Rua A, 1", "", "Ativo",
                         "50", "10", "", "Aqua Gestão", "", "", meta]
        vis = sys.modules["gspread"].PLANILHA.worksheet("🔬 Visitas")
        vis.linhas = vis.linhas[:1] + [
            _visita(1, "08/09/2026", [{"nome": "Piscina Adulto", "ph": "7,3", "cloro_livre": "1,4"},
                                      {"nome": "Piscina Infantil", "ph": "7,6", "cloro_livre": "2,1"}]),
            _visita(2, "15/09/2026", [{"nome": "Piscina Adulto", "ph": "7,5", "cloro_livre": "3,4"},
                                      {"nome": "Piscina Infantil", "ph": "7,4", "cloro_livre": "1,8"}]),
            _visita(3, "12/08/2026", [{"nome": "Piscina Adulto", "ph": "7,1", "cloro_livre": "1,0"}]),
        ]
        st_real.cache_data.clear()

    @staticmethod
    def _texto(p):
        if p.suffix == ".docx":
            with zipfile.ZipFile(p) as z:
                return " ".join(n.text or "" for n in ET.fromstring(z.read("word/document.xml")).iter() if n.tag.endswith("}t"))
        from pypdf import PdfReader
        return " ".join(pg.extract_text() or "" for pg in PdfReader(p).pages)

    def _evidencia(self, paths):
        destino = os.environ.get("RT_TEST_EVIDENCE")
        if destino:
            Path(destino).mkdir(parents=True, exist_ok=True)
            for p in paths:
                shutil.copy2(p, Path(destino) / p.name)

    def test_relatorio_completo_com_piscinas_fotos_art_e_consolidado(self):
        from PIL import Image
        self._preparar_alfa()
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        ss = at.session_state
        # Uma linha por piscina de cada visita, identificada.
        linhas = [(ss[f"rel_analise_piscina_{i}"], ss[f"rel_analise_data_{i}"], ss[f"rel_analise_ph_{i}"]) for i in range(4)]
        self.assertEqual(linhas, [("Piscina Adulto", "08/09/2026", "7,3"), ("Piscina Infantil", "08/09/2026", "7,6"),
                                  ("Piscina Adulto", "15/09/2026", "7,5"), ("Piscina Infantil", "15/09/2026", "7,4")])
        self.assertEqual(at.selectbox(key="rel_analise_piscina_0").options[1:4],
                         ["Piscina Adulto", "Piscina Infantil", "SPA Cobertura"])
        # Foto com legenda e piscina.
        buf = io.BytesIO()
        Image.new("RGB", (23, 29), (10, 120, 200)).save(buf, format="PNG")
        next(w for w in at.file_uploader if w.label == "Upload de fotos do relatório").upload(
            "alfa.png", buf.getvalue(), "image/png").run()
        leg = next(w for w in at.text_input if w.label == "Legenda — foto 1")
        leg.input("MARCADOR_LEGENDA_CASA_MAQUINAS").run()
        next(w for w in at.selectbox if w.label == "Piscina — foto 1").select("Piscina Infantil").run()
        fotos = ler_fotos(ss.filtered_state, ss["rel_origem_dados"])
        self.assertEqual((fotos[0]["legenda"], fotos[0]["piscina"]), ("MARCADOR_LEGENDA_CASA_MAQUINAS", "Piscina Infantil"))
        # ART emitida sem vigência: permitido; número aparece, vigência não é inventada.
        at.selectbox(key="rel_art_status_widget").select("Emitida").run()
        at.text_input(key="rel_art_numero").input("TESTE0001").run()
        self.assertFalse(at.exception, at.exception)
        next(b for b in at.button if b.label == "🚀 Gerar relatório mensal").click().run()
        self.assertFalse(at.exception, at.exception)
        docs = [p for p in (self.pasta / "app" / "generated").glob("**/*Alfa*RELATORIO_RT.*") if "PREVIA" not in p.name]
        self.assertEqual({p.suffix for p in docs}, {".docx", ".pdf"})
        self._evidencia(docs)
        for p in docs:
            texto = self._texto(p)
            for esperado in ("ACQUA GESTÃO", "Piscina Adulto", "Piscina Infantil", "SPA Cobertura", "50 m³",
                             "MARCADOR_LEGENDA_CASA_MAQUINAS", "TESTE0001", "Vigência não informada"):
                self.assertIn(esperado, texto, (p.name, esperado))
            self.assertNotIn("Aqua Gestão – Controle", texto)
        from pypdf import PdfReader
        pdf = PdfReader(next(p for p in docs if p.suffix == ".pdf"))
        primeira = pdf.pages[0].extract_text()
        self.assertIn("ACQUA GESTÃO", primeira)
        self.assertIn("Condomínio Alfa Fictício", primeira)
        self.assertIn("Piscina: Piscina Adulto", self._texto(next(p for p in docs if p.suffix == ".pdf")))
        imagens = [im.image.size for pg in pdf.pages for im in pg.images]
        self.assertIn((23, 29), imagens)
        self.assertGreaterEqual(sum(1 for s in imagens if s[0] == 1600), 2)  # um gráfico por piscina
        # Troca de cliente e volta: legenda e foto preservadas no contexto de Alfa.
        self._carregar(at, "Residencial Beta Fictício")
        self.assertEqual(ler_fotos(at.session_state.filtered_state, at.session_state["rel_origem_dados"]), [])
        self._carregar(at, "Condomínio Alfa Fictício")
        fotos = ler_fotos(at.session_state.filtered_state, at.session_state["rel_origem_dados"])
        self.assertEqual(fotos[0]["legenda"], "MARCADOR_LEGENDA_CASA_MAQUINAS")
        # Nova sessão: restauração do rascunho devolve foto, legenda e piscina.
        at2 = self._entrar_e_carregar()
        self._mes(at2, "09")
        import re
        rascunhos = [json.loads(p.read_text()) for p in (self.pasta / "app" / "generated").glob("_rascunho_relatorio_rt*.json")]
        alvo = max((d for d in rascunhos if d.get("campos", {}).get("rel_origem_dados") == "aqua_gestao|id:C1|09/2026"
                    and "MARCADOR_LEGENDA" in str(d["campos"].get("rel_fotos_preservadas"))),
                   key=lambda d: str(d.get("salvo_em") or d.get("_chave")))
        chave = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", alvo["_chave"])
        next(b for b in at2.button if b.key == chave).click().run()
        at2.run()
        fotos2 = ler_fotos(at2.session_state.filtered_state, at2.session_state["rel_origem_dados"])
        self.assertEqual((fotos2[0]["legenda"], fotos2[0]["piscina"]), ("MARCADOR_LEGENDA_CASA_MAQUINAS", "Piscina Infantil"))
        self.assertEqual(next(w for w in at2.text_input if w.label == "Legenda — foto 1").value, "MARCADOR_LEGENDA_CASA_MAQUINAS")
        # Consolidado trimestral (jul–set) a partir das visitas identificadas.
        at2.selectbox(key="rtc_tipo").select("Trimestral").run()
        at2.selectbox(key="rtc_trimestre").select(3).run()
        at2.number_input(key="rtc_ano").set_value(2026).run()
        next(b for b in at2.button if b.key == "btn_rtc_gerar").click().run()
        self.assertFalse(at2.exception, at2.exception)
        self.assertTrue(any("3º trimestre de 2026" in s.value for s in at2.success), [e.value for e in at2.error])
        consolidados = list((self.pasta / "app" / "generated" / "_rt_consolidados").glob("**/*.pdf"))
        self.assertEqual(len(consolidados), 1)
        self._evidencia(consolidados)
        texto = self._texto(consolidados[0])
        for esperado in ("ACQUA GESTÃO", "3º trimestre de 2026", "Piscina Adulto", "Piscina Infantil",
                         "processo de emissão", "08/2026", "09/2026"):
            self.assertIn(esperado, texto)
        self.assertNotIn("TESTE0001", texto)  # ART mensal não é herdada pelo consolidado
        self.assertNotIn("MARCADOR_LEGENDA", texto)

    def test_consolidado_bloqueia_art_emitida_sem_numero_e_relatorio_avulso(self):
        self._preparar_alfa()
        at = self._entrar_e_carregar()
        at.selectbox(key="rtc_art_status").select("Emitida").run()
        next(b for b in at.button if b.key == "btn_rtc_gerar").click().run()
        self.assertTrue(any("número" in e.value for e in at.error))
        self.assertFalse(list((self.pasta / "app" / "generated").glob("_rt_consolidados/**/*.pdf")))


del Base

if __name__ == "__main__":
    unittest.main()
