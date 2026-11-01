"""Relatório RT ACQUA GESTÃO: regras puras, sem Streamlit, rede ou dados reais."""
import io
import json
import tempfile
import unittest
from pathlib import Path

from utils import rt_relatorio as rr
from utils.cadastro_piscinas import serializar
from utils.rt_contexto import (atualizar_metadados_fotos, campos, ler_fotos, limpar, materializar_fotos,
                               remover_foto, sincronizar_fotos)


class Foto(io.BytesIO):
    def __init__(self, dados, nome="foto.png"):
        super().__init__(dados)
        self.name = nome


class ART(unittest.TestCase):
    def test_legado_e_vazio_nunca_viram_emitida(self):
        for valor in ("", None, "Em processo de formalização", "Em tramitação", "em tramitacao administrativa", "xyz"):
            self.assertEqual(rr.normalizar_status_art(valor), rr.ART_EM_EMISSAO, valor)
        self.assertEqual(rr.normalizar_status_art("Não emitida"), rr.ART_NAO_EMITIDA)
        self.assertEqual(rr.normalizar_status_art("Emitida"), rr.ART_EMITIDA)
        self.assertEqual(rr.ART_OPCOES, ("Não emitida", "Em processo de emissão", "Emitida"))

    def test_texto_so_mostra_numero_e_vigencia_informados(self):
        self.assertIn("processo de emissão", rr.texto_art("Em processo de emissão", "123", "01/01/2026", "31/12/2026"))
        self.assertNotIn("123", rr.texto_art("Não emitida", "123"))
        self.assertEqual(rr.texto_art("Emitida", "TESTE-1"), "ART nº TESTE-1 | Vigência não informada")
        self.assertEqual(rr.texto_art("Emitida", "TESTE-1", "01/01/2026", "31/12/2026"),
                         "ART nº TESTE-1 | Vigência: 01/01/2026 a 31/12/2026")

    def test_validacao(self):
        self.assertEqual(rr.validar_art("Em processo de emissão"), [])
        self.assertEqual(rr.validar_art("Não emitida", "", "lixo"), [])  # campos ignorados sem emissão
        self.assertTrue(rr.validar_art("Emitida", ""))
        self.assertTrue(rr.validar_art("Emitida", "N/A"))
        self.assertEqual(rr.validar_art("Emitida", "TESTE-1"), [])
        self.assertTrue(rr.validar_art("Emitida", "TESTE-1", "01/01/2026", ""))
        self.assertTrue(rr.validar_art("Emitida", "TESTE-1", "31/12/2026", "01/01/2026"))
        self.assertTrue(rr.validar_art("Emitida", "TESTE-1", "32/01/2026", "01/02/2026"))
        self.assertEqual(rr.validar_art("Emitida", "TESTE-1", "01/01/2026", "31/12/2026"), [])


class Piscinas(unittest.TestCase):
    def test_cadastro_novo_com_tipos_e_extras(self):
        meta = json.loads(serializar({"adulto": "Adulto", "infantil": "Infantil", "extra1": "SPA"},
                                     [{"nome": "SPA Cobertura", "vol": 4.5, "tipo": "SPA"}]))
        res = rr.piscinas_do_cliente({"vol_adulto": 50, "vol_infantil": 0, "vol_total": 54.5,
                                      "cadastro_piscinas": meta})
        self.assertEqual([(p["nome"], p["tipo"], p["volume"]) for p in res["piscinas"]],
                         [("Piscina Adulto", "Adulto", 50.0), ("Piscina Infantil", "Infantil", None),
                          ("SPA Cobertura", "SPA", 4.5)])
        self.assertFalse(res["sem_distribuicao"])

    def test_cadastro_antigo_total_sem_distribuicao_nao_e_atribuido(self):
        res = rr.piscinas_do_cliente({"vol_total": 80, "vol_adulto": 0})
        self.assertEqual(res["piscinas"], [])
        self.assertTrue(res["sem_distribuicao"])
        self.assertEqual(res["volume_total"], 80)

    def test_cadastro_antigo_com_volumes_e_extras_locais(self):
        res = rr.piscinas_do_cliente({"vol_adulto": "30", "vol_family": "12,5",
                                      "piscinas_extras": [{"nome": "Raia", "vol": 20}, {"nome": ""}]})
        self.assertEqual([(p["nome"], p["tipo"]) for p in res["piscinas"]],
                         [("Piscina Adulto", "Não informado"), ("Piscina Family", "Não informado"), ("Raia", "Não informado")])
        self.assertEqual(rr.formatar_volume(12.5), "12,5 m³")
        self.assertEqual(rr.formatar_volume(None), "Não informado")

    def test_metadados_corrompidos_nao_quebram(self):
        res = rr.piscinas_do_cliente({"vol_adulto": 10, "cadastro_piscinas": "{lixo"})
        self.assertEqual(len(res["piscinas"]), 1)


class Analises(unittest.TestCase):
    def test_leitura_inicial_e_agrupamento(self):
        self.assertEqual(rr.leitura("7,2 / 7,6"), 7.2)
        self.assertIsNone(rr.leitura("—"))
        grupos = rr.agrupar_por_piscina([
            {"piscina": "Adulto", "ph": "7,4"}, {"piscina": "", "ph": "7,0"},
            {"piscina": "Infantil", "ph": ""}, {"piscina": "Adulto", "cloro_livre": "1,5"}])
        self.assertEqual([(n, len(i)) for n, i in grupos], [("Adulto", 2), (rr.SEM_PISCINA, 1)])

    def test_visita_com_varias_piscinas_gera_uma_linha_por_piscina(self):
        linhas = rr.linhas_de_lancamentos([
            {"data": "01/09/2026", "piscinas": [{"nome": "Adulto", "ph": "7,4"}, {"nome": "Infantil", "ph": "7,6"}]},
            {"data": "08/09/2026", "ph": "7,2"}])
        self.assertEqual([(l["piscina"], l["dados"]["ph"]) for l in linhas],
                         [("Adulto", "7,4"), ("Infantil", "7,6"), ("", "7,2")])


class Periodos(unittest.TestCase):
    def test_meses(self):
        self.assertEqual(rr.meses_do_periodo("Trimestral", 3, 2026), [(7, 2026), (8, 2026), (9, 2026)])
        self.assertEqual(len(rr.meses_do_periodo("Semestral", 2, 2026)), 6)
        self.assertEqual(rr.meses_do_periodo("Anual", 9, 2026)[0], (1, 2026))
        for args in (("Trimestral", 5, 2026), ("Semestral", 0, 2026), ("Mensal", 1, 2026), ("Anual", 1, 1999)):
            with self.assertRaises(ValueError):
                rr.meses_do_periodo(*args)
        self.assertEqual(rr.rotulo_periodo("Semestral", 1, 2026), "1º semestre de 2026")

    def test_consolidacao_por_piscina_e_periodo(self):
        meses = rr.meses_do_periodo("Trimestral", 3, 2026)
        linhas = [
            # Limites NBR 10818:2026: pH 7,0–7,8 e CRL 1,0–5,0 (6,9 e 0,8 ficam fora).
            {"data": "01/07/2026", "piscina": "Adulto", "ph": "6,9", "cloro_livre": "0,8"},
            {"data": "15/08/2026", "piscina": "Adulto", "ph": "7,6", "cloro_livre": "4,0"},
            {"data": "15/08/2026", "piscina": "Infantil", "ph": "7,4"},
            {"data": "01/10/2026", "piscina": "Adulto", "ph": "9,9"},   # fora do período
            {"data": "sem data", "piscina": "Adulto", "ph": "9,9"},
        ]
        c = rr.consolidar(linhas, meses)
        self.assertEqual(set(c["piscinas"]), {"Adulto", "Infantil"})
        ph = c["piscinas"]["Adulto"]["parametros"]["ph"]
        self.assertEqual((ph["n"], ph["min"], ph["max"]), (2, 6.9, 7.6))
        self.assertAlmostEqual(ph["conformidade"], 50.0)
        self.assertAlmostEqual(c["piscinas"]["Adulto"]["parametros"]["cloro_livre"]["conformidade"], 50.0)
        self.assertEqual(set(c["piscinas"]["Adulto"]["mensal"]), {"07/2026", "08/2026"})
        self.assertEqual(c["descartadas"], 2)

    def test_graficos(self):
        with tempfile.TemporaryDirectory() as tmp:
            analises = [{"piscina": "Adulto", "data": "0%d/09/2026" % d, "ph": "7,%d" % d, "cloro_livre": "1,%d" % d}
                        for d in range(1, 5)] + [{"piscina": "SPA", "data": "02/09/2026", "ph": "8,1"}]
            graficos = rr.graficos_por_piscina(analises, Path(tmp), "Cond Ficticio", "09/2026")
            self.assertEqual([n for n, _ in graficos], ["Adulto", "SPA"])
            from PIL import Image
            self.assertEqual(Image.open(graficos[0][1]).size[0], 1600)
            self.assertEqual(rr.graficos_por_piscina([{"ph": ""}], Path(tmp), "x", "y"), [])


class FotosComLegenda(unittest.TestCase):
    CH = "aqua_gestao|id:C1|09/2026"

    def test_legenda_e_piscina_sobrevivem_ao_rascunho_e_ao_reenvio(self):
        ss = {}
        foto = Foto(b"imagem-ficticia")
        fotos = sincronizar_fotos(ss, self.CH, [foto])
        digest = fotos[0]["sha256"]
        self.assertTrue(atualizar_metadados_fotos(ss, self.CH, {digest: ("Casa de máquinas", "Piscina Adulto")}))
        self.assertFalse(atualizar_metadados_fotos(ss, self.CH, {digest: ("Casa de máquinas", "Piscina Adulto")}))
        sincronizar_fotos(ss, self.CH, [Foto(b"imagem-ficticia")])  # mesmo upload no rerun
        self.assertEqual(rr.metadados_foto(ler_fotos(ss, self.CH)[0]), ("Casa de máquinas", "Piscina Adulto"))
        salvo = json.loads(json.dumps(campos(ss)))  # ida e volta como no rascunho
        limpar(ss)
        self.assertEqual(ler_fotos(ss, self.CH), [])
        ss.update(salvo)
        f = ler_fotos(ss, self.CH)[0]
        self.assertEqual(rr.legenda_foto(2, f), "Foto 2 — Piscina Adulto — Casa de máquinas")
        with tempfile.TemporaryDirectory() as tmp:
            p = materializar_fotos(ss, self.CH, tmp)[0]
            self.assertEqual(p.stem, digest)  # o PDF/DOCX recupera a legenda pelo SHA-256
        with self.assertRaises(ValueError):
            atualizar_metadados_fotos(ss, "aqua_gestao|id:C2|09/2026", {digest: ("x", "y")})

    def test_remocao_explicita_e_legenda_padrao(self):
        ss = {}
        a, b = sincronizar_fotos(ss, self.CH, [Foto(b"a"), Foto(b"b")])
        nonce = ss.get("_rt_fotos_nonce", 0)
        remover_foto(ss, self.CH, a["sha256"])
        self.assertEqual([f["sha256"] for f in ler_fotos(ss, self.CH)], [b["sha256"]])
        self.assertEqual(ss["_rt_fotos_nonce"], nonce + 1)
        sincronizar_fotos(ss, self.CH, [])  # uploader recriado vazio não apaga a foto restante
        self.assertEqual(len(ler_fotos(ss, self.CH)), 1)
        self.assertEqual(rr.legenda_foto(1, {}), "Foto 1 — Vista geral da piscina no período de referência")
        self.assertEqual(rr.legenda_foto(1, {"legenda": "x" * 500})[:9], "Foto 1 — ")


class PDFConsolidado(unittest.TestCase):
    def test_pdf_com_capa_piscinas_art_e_graficos(self):
        from pypdf import PdfReader
        from utils.pdf_rt_consolidado import gerar_pdf_consolidado
        meses = rr.meses_do_periodo("Trimestral", 3, 2026)
        cons = rr.consolidar([{"data": "0%d/0%d/2026" % (d, m), "piscina": "Piscina Adulto", "ph": "7,%d" % d,
                               "cloro_livre": "1,2"} for m in (7, 8, 9) for d in (1, 5)], meses)
        with tempfile.TemporaryDirectory() as tmp:
            graficos = rr.graficos_consolidados(cons, Path(tmp), "Condomínio Fictício", "3º trimestre de 2026")
            ctx = {"tipo": "Trimestral", "periodo_rotulo": "3º trimestre de 2026", "meses": cons["meses"],
                   "nome_condominio": "Condomínio Fictício", "consolidado": cons, "graficos": graficos,
                   "piscinas_cadastro": rr.piscinas_do_cliente({"vol_adulto": 50}),
                   "art": {"status": "Em processo de formalização"}, "responsavel": "RT Fictício",
                   "crq": "CRQ 000", "data_emissao": "31/10/2026"}
            ok, erro = gerar_pdf_consolidado(ctx, Path(tmp) / "c.pdf", None)
            self.assertTrue(ok, erro)
            pdf = PdfReader(Path(tmp) / "c.pdf")
            texto = " ".join(p.extract_text() or "" for p in pdf.pages)
            for esperado in ("ACQUA GESTÃO", "3º trimestre de 2026", "Piscina Adulto", "50 m³",
                             "processo de emissão", "Médias mensais"):
                self.assertIn(esperado, texto)
            self.assertNotIn("ART nº", texto)
            self.assertTrue(any(pg.images for pg in pdf.pages))
            import fitz
            with fitz.open(Path(tmp) / "c.pdf") as rendered:
                fontes = [f[3] for pg in rendered for f in pg.get_fonts()]
                self.assertTrue(any("BitstreamVeraSans" in fonte for fonte in fontes))


if __name__ == "__main__":
    unittest.main()
