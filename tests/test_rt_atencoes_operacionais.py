"""Regressão do mensal: limite aceitável, ideal e meta interna são critérios distintos.

Dados de teste reproduzem as faixas do caso relatado, sem acesso a clientes/Sheets.
"""
import os
import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image

from test_mensal_rt_app_pdf import _carregar, ROOT
from utils.normas_piscinas import classificacao_operacional, avaliar_leitura
from utils.rt_relatorio import pontos_atencao, interpretacao_resultados, gerar_grafico, PARAMETROS_MENSAIS


def analises():
    valores = [
        ("08/10/2026", "7,7", "2,85", "2,93", "85", "160", "38", "0,16", "541", "654", "29"),
        ("15/10/2026", "7,6", "3,88", "4,01", "75", "150", "42", "0", "645", "587", "28,8"),
        ("22/10/2026", "7,5", "2,0", "2,6", "88", "150", "48", "0,28", "612", "784", "31"),
        ("29/10/2026", "7,8", "3,08", "3,12", "95", "160", "55", "0", "547", "654", "31,6"),
    ]
    campos = ("data", "ph", "cloro_livre", "cloro_total", "alcalinidade", "dureza", "cianurico", "turbidez", "orp", "tds", "temperatura")
    return [{**dict(zip(campos, linha)), "piscina": "Adulto", "operador": "RT de teste"} for linha in valores]


class AtencoesOperacionais(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _carregar()

    def avaliar(self, linhas=None):
        linhas = analises() if linhas is None else linhas
        avaliacao = self.ns["avaliar_conformidade_analises"](linhas)
        textos = self.ns["gerar_textos_automaticos_relatorio"](linhas, avaliacao)
        return linhas, avaliacao, textos

    def test_limites_ideais_e_meta_interna(self):
        for campo, valor, esperado in [
            ("cya", 50, "conforme"), ("cya", 55, "atencao"), ("cya", 100, "atencao"), ("cya", 101, "fora_limite"),
            ("cc", .49, "conforme"), ("cc", .50, "atencao"), ("cc", .60, "atencao"), ("cc", .61, "fora_limite"),
            ("ph", 7.6, "conforme"), ("ph", 7.8, "atencao"), ("ph", 7.9, "fora_limite"),
            ("alcalinidade", 80, "conforme"), ("alcalinidade", 75, "atencao"), ("dureza", 160, "atencao"),
            ("orp", 649, "atencao"), ("orp", 650, "conforme"), ("turbidez", .5, "fora_limite"),
            ("tds", 654, "informativo"), ("cc", None, "ausente"), ("ph", float("nan"), "ausente"),
        ]:
            with self.subTest(campo=campo, valor=valor):
                self.assertEqual(classificacao_operacional(campo, valor), esperado)
        self.assertEqual(avaliar_leitura("cc", .60), "dentro_ideal")  # sem mudar o limite normativo

    def test_caso_mensal_preserva_todas_ocorrencias_e_recomendacoes(self):
        linhas, avaliacao, textos = self.avaliar()
        self.assertEqual(avaliacao["status"], "CONFORME")
        self.assertEqual(avaliacao["detalhes"], [])
        self.assertIn("com pontos de atenção operacional", self.ns["rotulo_status_agua"]("CONFORME", avaliacao))
        grupos = {g["campo"]: g for g in pontos_atencao(linhas)}
        self.assertEqual(set(grupos), {"ph", "alcalinidade", "dureza", "cianurico", "cloro_combinado", "orp"})
        obs = " ".join(textos["observacoes"])
        recs = " ".join(textos["recomendacoes"])
        self.assertLessEqual(len(textos["observacoes"]), 5)
        self.assertLessEqual(len(textos["recomendacoes"]), 5)
        for grupo in grupos.values():
            self.assertIn(grupo["texto"], obs)
            self.assertIn(grupo["recomendacao"], recs)
        for dia in ("08", "15", "22", "29"):
            self.assertIn(f"{dia}/10/2026", grupos["orp"]["texto"])
            self.assertIn(f"{dia}/10/2026", grupos["dureza"]["texto"])
        self.assertIn("29/10/2026, Adulto (55 mg/L)", grupos["cianurico"]["texto"])
        self.assertIn("22/10/2026, Adulto (0,60 mg/L, no limite aceitável", grupos["cloro_combinado"]["texto"])
        self.assertNotIn("acima do limite", grupos["cloro_combinado"]["texto"])
        self.assertIn("pontos de atenção", textos["diagnostico"])

    def test_nao_conformidade_e_ausencia_mantem_prioridade(self):
        _, aval, _ = self.avaliar([{"cloro_livre": "2", "cloro_total": "2,61", "cianurico": "101"}])
        self.assertEqual(aval["status"], "NÃO CONFORME")
        self.assertEqual(len(aval["detalhes"]), 2)
        for linhas in ([], [{"tds": "654"}], [{"orp": "500"}]):
            _, aval, _ = self.avaliar(linhas)
            self.assertEqual(aval["status"], "NÃO AVALIADO")
        _, aval, _ = self.avaliar([{"cloro_livre": "2", "cloro_total": "1,8"}])
        self.assertEqual(aval["status"], "NÃO CONFORME")
        self.assertFalse(aval["grupos_atencao"])

    def test_piscinas_separadas_e_orp_sem_recomendacao_indevida(self):
        linhas = [{"data": "29/10/2026", "piscina": "Adulto", "cianurico": "55"},
                  {"data": "29/10/2026", "piscina": "Infantil", "cianurico": "40"}]
        _, aval, textos = self.avaliar(linhas)
        self.assertNotIn("Infantil", " ".join(aval["atencoes"]))
        self.assertNotIn("ORP", " ".join(textos["recomendacoes"]))
        for linhas in ([{"cloro_livre": "2"}], [{"cloro_total": "2,5"}]):
            self.assertFalse(pontos_atencao(linhas))  # não inventa CC sem o par

    def test_interpretacao_inclui_meta_e_datas(self):
        grupos = dict(interpretacao_resultados(analises()))
        texto = grupos["Estabilizante e cloro combinado"]
        for trecho in ("fora da faixa ideal 30–50", "29/10/2026", "CC < 0,50", "no limite aceitável de 0,60", "22/10/2026"):
            self.assertIn(trecho, texto)

    def test_rascunho_antigo_preserva_edicoes_e_recebe_alertas(self):
        ns = _carregar()
        linhas = analises()
        ns["st"].session_state = {"rel_diagnostico": "No período avaliado, as leituras registradas foram conferidas. NOTA MANUAL DO RT.",
                                  "rel_status_agua": "CONFORME", "rel_data_emissao": "17/11/2026"}
        ns.update({"obter_snapshot_relatorio_independente": lambda: {}, "coletar_analises_relatorio": lambda: linhas,
                   "coletar_recomendacoes_relatorio": lambda: [{"recomendacao": "ORIENTAÇÃO MANUAL", "prazo": "20/11", "responsavel": "RT"}],
                   "coletar_observacoes_relatorio": lambda: ["OBSERVAÇÃO MANUAL"], "_empresa_ativa_nome": lambda: "Aqua Gestão",
                   "_piscinas_relatorio_rt": lambda: {}, "_metadados_fotos_contexto_rt": lambda: {},
                   "coletar_dosagens_relatorio": lambda: [], "coletar_conformidades_relatorio": lambda: {}, "coletar_epis_relatorio": lambda: {}})
        dados = ns["montar_dados_relatorio"]()
        self.assertIn("NOTA MANUAL DO RT.", dados["diagnostico"])
        self.assertIn("pontos de atenção operacional", dados["diagnostico"])
        self.assertEqual(dados["observacoes"][0], "OBSERVAÇÃO MANUAL")
        self.assertEqual(dados["recomendacoes"][0]["prazo"], "20/11")
        for g in pontos_atencao(linhas):
            self.assertIn(g["texto"], dados["observacoes"])
            self.assertIn(g["recomendacao"], [r["recomendacao"] for r in dados["recomendacoes"]])
        self.assertNotIn("pontos de atenção", ns["st"].session_state["rel_diagnostico"])  # sem regravar o rascunho

    def test_grafico_distingue_atencao_de_fora_limite(self):
        with tempfile.TemporaryDirectory() as pasta:
            for valor, cor in ((55, (201, 107, 44)), (101, (180, 35, 24))):
                p = gerar_grafico([("cianurico", [valor])], ["29/10"], "Teste", "Adulto", Path(pasta) / "grafico.png", PARAMETROS_MENSAIS)
                with Image.open(p) as im:
                    self.assertIn(cor, {c for _n, c in im.getcolors(im.width * im.height)})

    def test_pdf_real_status_alertas_e_destaques(self):
        linhas, aval, textos = self.avaliar()
        dados = {
            "nome_condominio": "CONDOMÍNIO TESTE — ATENÇÕES", "mes_referencia": "10", "ano_referencia": "2026",
            "data_emissao": "17/11/2026", "analises": linhas, "status_agua": aval["status"], "avaliacao_automatica": aval,
            "diagnostico": textos["diagnostico"], "observacoes": textos["observacoes"],
            "recomendacoes": [{"recomendacao": t, "prazo": "Próxima rotina", "responsavel": "RT"} for t in textos["recomendacoes"]],
            "art_status": "Em processo de emissão", "piscinas_cadastro": {"piscinas": [{"nome": "Adulto", "tipo": "Adulto", "volume": 173}]},
        }
        with tempfile.TemporaryDirectory() as pasta:
            pdf = Path(pasta) / "mensal_atencoes_teste.pdf"
            self.ns["encontrar_logo"] = lambda: ROOT / "acqua_gestao_logo_rt.png"
            self.ns["_log_sheets_erro"] = lambda *args: None
            ok, erro = self.ns["gerar_pdf_relatorio_rt_premium_reportlab"](dados, [], pdf)
            self.assertTrue(ok, erro)
            with fitz.open(pdf) as doc:
                texto = " ".join(" ".join(p.get_text() for p in doc).split())
                for trecho in ("com pontos de atenção operacional", "fora da faixa ideal 30–50", "no limite aceitável de 0,60", "CC < 0,50"):
                    self.assertIn(trecho, texto)
                for dia in ("08", "15", "22", "29"):
                    self.assertIn(f"{dia}/10/2026, Adulto", texto)
                spans = [s for p in doc for b in p.get_text("dict")["blocks"] if "lines" in b for l in b["lines"] for s in l["spans"]]
                valores_ambar = [s["text"].strip() for s in spans if s["color"] == 0x8F4D00]
                self.assertIn("55", valores_ambar)
                self.assertIn("0,60", valores_ambar)
                destino = os.environ.get("RT_TEST_EVIDENCE")
                if destino:
                    Path(destino).mkdir(parents=True, exist_ok=True)
                    doc.save(Path(destino) / pdf.name)


if __name__ == "__main__":
    unittest.main()
