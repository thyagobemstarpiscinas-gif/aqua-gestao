"""Rascunhos do relatório RT por identidade e origem RT; fotos preservadas nas atualizações.

1. A seção do relatório RT lista e restaura só rascunhos da identidade carregada
   (empresa + ID do condomínio + período) gravados por acesso RT autenticado, com
   lançamentos de origem RT. Operador, outro condomínio/período e registros antigos sem
   origem não aparecem (continuam preservados).
2. Lista vazia do uploader não apaga fotos/legendas; "Carregar dados do cadastro no
   relatório" não interrompe a execução (antes: StreamlitAPIException em rel_tipo_atendimento).

App real (AppTest), planilha falsa em memória, dados fictícios, sem rede.
"""
import io
import json
import unittest

import test_rt_relatorio_avaliacao as T
from rt_monthly_report import identify_record_origin
from utils.rt_contexto import ler_fotos, motivo_rascunho_oculto, sincronizar_fotos

AppTest = T.AppTest
ALFA_09 = "aqua_gestao|id:C1|09/2026"
ALFA_08 = "aqua_gestao|id:C1|08/2026"
BETA_09 = "aqua_gestao|id:C2|09/2026"
RT_OK = json.dumps([{"id": "VA1", "rt": True}])


def _rascunho(chave, origem_dados, nome, origem_acesso="escritorio_rt", lancamentos=RT_OK, marcador="X"):
    d = {"_chave": chave, "salvo_em": "08/11/2026 10:00:00", "_salvo_iso": "2026-11-08T10:00:00",
         "empresa_ativa": "aqua_gestao", "rel_analises_total": 4,
         "campos": {"rel_nome_condominio": nome, "rel_origem_dados": origem_dados,
                    "rel_mes_referencia": origem_dados[-7:-5], "rel_ano_referencia": "2026",
                    "rel_analise_data_0": "01/09/2026", "rel_analise_ph_0": "7,4", "rel_diagnostico": marcador,
                    "rel_lancamentos_origem": lancamentos}}
    if origem_acesso is not None:
        d["origem_acesso"] = origem_acesso
    return d


class Foto:
    def __init__(self, conteudo, nome="f.png"):
        self.conteudo, self.name = conteudo, nome

    def getvalue(self):
        return self.conteudo


class TestRegrasOrigem(unittest.TestCase):
    def test_origem_so_por_marca_do_acesso_autenticado(self):
        self.assertTrue(identify_record_origin({"visita_rt_semanal": True})["is_rt"])
        self.assertTrue(identify_record_origin({"tipo_visita": "RT semanal"})["is_rt"])
        self.assertTrue(identify_record_origin({"is_rt": True, "tipo_visita": "Relatório mensal RT"})["is_rt"])
        # Nome da pessoa / campo operador não fazem um lançamento virar RT.
        nome = identify_record_origin({"operador": "Responsável Técnico Fictício"})
        self.assertEqual((nome["is_rt"], nome["reason"]), (False, "unknown"))
        self.assertTrue(identify_record_origin({"visita_rt_semanal": False, "tipo_visita": "Operador",
                                                "operador": "Thyago"})["is_operator"])

    def test_motivo_rascunho_oculto(self):
        ok = _rascunho("thyago#s#a1", ALFA_09, "Alfa")
        self.assertEqual(motivo_rascunho_oculto(ok, "aqua_gestao", ALFA_09), "")
        self.assertEqual(motivo_rascunho_oculto(ok, "aqua_gestao", ""), "sem_identidade")
        self.assertEqual(motivo_rascunho_oculto(ok, "aqua_gestao", BETA_09), "outra_identidade")
        self.assertEqual(motivo_rascunho_oculto(ok, "bem_star", ALFA_09), "outra_identidade")
        for ruim in (_rascunho("thyago#s#a2", ALFA_09, "Alfa", origem_acesso="operador"),
                     _rascunho("thyago#s#a3", ALFA_09, "Alfa", origem_acesso=None),  # antigo: sem origem
                     _rascunho("thyago#s#a4", ALFA_09, "Alfa", lancamentos=json.dumps([{"id": "V9", "rt": False}]))):
            self.assertEqual(motivo_rascunho_oculto(ruim, "aqua_gestao", ALFA_09), "origem_nao_rt")

    def test_uploader_vazio_nao_apaga_fotos_e_remocao_intencional(self):
        from utils.rt_contexto import atualizar_metadados_fotos
        ss = {}
        a, b = sincronizar_fotos(ss, ALFA_09, [Foto(b"a", "a.png"), Foto(b"b", "b.png")])
        atualizar_metadados_fotos(ss, ALFA_09, {a["sha256"]: ("LEG A", ""), b["sha256"]: ("LEG B", "")})
        sincronizar_fotos(ss, ALFA_09, [])  # execução anterior parou antes do uploader
        self.assertEqual([f.get("legenda") for f in ler_fotos(ss, ALFA_09)], ["LEG A", "LEG B"])
        sincronizar_fotos(ss, ALFA_09, [Foto(b"a", "a.png"), Foto(b"b", "b.png")])  # o navegador reenvia
        self.assertEqual(len(ler_fotos(ss, ALFA_09)), 2)
        sincronizar_fotos(ss, ALFA_09, [Foto(b"a", "a.png")])  # removeu "b" do upload: remoção intencional
        self.assertEqual([f.get("legenda") for f in ler_fotos(ss, ALFA_09)], ["LEG A"])


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestRascunhosEFotosNoApp(unittest.TestCase):
    setUp = T.TestRelatorioRTAvaliacaoCompleta.setUp
    _preparar = T.TestRelatorioRTAvaliacaoCompleta._preparar
    _entrar_e_carregar = T.TestRelatorioRTAvaliacaoCompleta._entrar_e_carregar
    _mes = T.TestRelatorioRTAvaliacaoCompleta._mes
    _carregar = T.TestRelatorioRTAvaliacaoCompleta._carregar
    _gerar = T.TestRelatorioRTAvaliacaoCompleta._gerar
    _paginas = staticmethod(T.TestRelatorioRTAvaliacaoCompleta._paginas)

    RASCUNHOS = {
        "thyago#s#rtalfa09": (ALFA_09, "Condomínio Alfa Fictício", "escritorio_rt"),
        "thyago#s#opalfa09": (ALFA_09, "Condomínio Alfa Fictício", "operador"),
        "thyago#s#antigoalfa09": (ALFA_09, "Condomínio Alfa Fictício", None),
        "thyago#s#rtalfa08": (ALFA_08, "Condomínio Alfa Fictício", "rt_campo"),
        "thyago#s#rtbeta09": (BETA_09, "Residencial Beta Fictício", "escritorio_rt"),
    }

    def _gravar_rascunhos(self):
        pasta = self.pasta / "app" / "generated"
        pasta.mkdir(parents=True, exist_ok=True)
        for chave, (ident, nome, origem) in self.RASCUNHOS.items():
            (pasta / f"_rascunho_relatorio_rt_{chave.replace('#', '_')}.json").write_text(
                json.dumps(_rascunho(chave, ident, nome, origem_acesso=origem, marcador=chave), ensure_ascii=False),
                encoding="utf-8")

    @staticmethod
    def _listados(at):
        return {b.key.replace("btn_restaurar_rascunho_rt_", "") for b in at.button
                if (b.key or "").startswith("btn_restaurar_rascunho_rt_")}

    def test_lista_por_identidade_e_origem_rt(self):
        self._preparar(T.VISITAS)
        self._gravar_rascunhos()
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        alfa = self._listados(at)
        self.assertIn("thyago_s_rtalfa09", alfa)  # RT do condomínio e período atuais
        for oculto in ("thyago_s_opalfa09", "thyago_s_antigoalfa09", "thyago_s_rtalfa08", "thyago_s_rtbeta09"):
            self.assertNotIn(oculto, alfa, oculto)
        self.assertTrue(any("2 rascunho(s) deste condomínio e período sem origem RT comprovada" in c.value
                            for c in at.caption), [c.value for c in at.caption])
        # Troca de condomínio: a lista muda na hora; voltar preserva os rascunhos.
        self._carregar(at, "Residencial Beta Fictício")
        self._mes(at, "09")
        beta = self._listados(at)
        self.assertIn("thyago_s_rtbeta09", beta)
        self.assertNotIn("thyago_s_rtalfa09", beta)
        self._carregar(at, "Condomínio Alfa Fictício")
        self._mes(at, "09")
        self.assertIn("thyago_s_rtalfa09", self._listados(at))
        for chave in self.RASCUNHOS:  # nenhum arquivo foi apagado para esconder a lista
            self.assertTrue((self.pasta / "app" / "generated" / f"_rascunho_relatorio_rt_{chave.replace('#', '_')}.json").exists())
        # Restaurar o rascunho RT listado funciona; o salvo nesta sessão grava a origem do acesso.
        next(b for b in at.button if b.key == "btn_restaurar_rascunho_rt_thyago_s_rtalfa09").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(at.session_state["rel_diagnostico"], "thyago#s#rtalfa09")

    def test_status_e_cadastro_preservam_duas_fotos_e_legendas_no_pdf(self):
        from PIL import Image
        self._preparar(T.VISITAS)
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        ss = at.session_state
        arquivos = []
        for nome, cor in (("um.png", (200, 30, 30)), ("dois.png", (30, 30, 200))):
            buf = io.BytesIO()
            Image.new("RGB", (64, 48), cor).save(buf, format="PNG")
            arquivos.append((nome, buf.getvalue(), "image/png"))
        next(w for w in at.file_uploader if w.label == "Upload de fotos do relatório").set_value(arquivos).run()
        for n, leg in ((1, "LEGENDA_UM"), (2, "LEGENDA_DOIS")):
            next(w for w in at.text_input if w.label == f"Legenda — foto {n}").input(leg).run()
        at.selectbox(key="rel_status_agua_select").select("CONFORME").run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual([f.get("legenda") for f in ler_fotos(ss.filtered_state, ss["rel_origem_dados"])],
                         ["LEGENDA_UM", "LEGENDA_DOIS"])
        pdf = self._gerar(at)
        leitor, paginas = self._paginas(pdf)
        texto = " ".join(paginas)
        self.assertIn("LEGENDA_UM", texto)
        self.assertIn("LEGENDA_DOIS", texto)
        self.assertEqual(sum(1 for pg in leitor.pages for im in pg.images if im.image.size == (64, 48)), 2)
        # "Carregar dados do cadastro no relatório": sem exceção; ao voltar a Alfa, fotos e legendas continuam.
        next(b for b in at.button if b.label == "Carregar dados do cadastro no relatório").click().run()
        self.assertFalse(at.exception, at.exception)
        self._carregar(at, "Condomínio Alfa Fictício")
        self._mes(at, "09")
        self.assertEqual([f.get("legenda") for f in ler_fotos(ss.filtered_state, ss["rel_origem_dados"])],
                         ["LEGENDA_UM", "LEGENDA_DOIS"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
