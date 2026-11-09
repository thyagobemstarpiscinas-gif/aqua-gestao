"""Relatório RT do escritório: o que acontece na tela depois de selecionar fotos.

Achado em uso: as fotos selecionadas "não apareciam" e não havia mensagem. Sem condomínio
carregado nada era incorporado, e com condomínio as miniaturas ficavam numa seção fechada.
App real (AppTest), planilha falsa em memória, imagens fictícias geradas no teste.
"""
import io
import unittest

import test_rt_relatorio_apptest as _base
from test_relatorio_rt_campos import PIN_FICTICIO
from utils.rt_contexto import ler_fotos

AppTest = _base.AppTest
UPLOADER = "Upload de fotos do relatório"
PAINEL = "📷 Legendas e piscinas das fotografias"


def _png(cor, tamanho=(17, 13)):
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", tamanho, cor).save(buf, format="PNG")
    return buf.getvalue()


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class TestFotosRTVisibilidade(unittest.TestCase):
    setUp = _base.TestRelatorioRTProfissional.setUp
    _entrar_e_carregar = _base.TestRelatorioRTProfissional._entrar_e_carregar
    _mes = _base.TestRelatorioRTProfissional._mes
    _carregar = _base.TestRelatorioRTProfissional._carregar
    _preparar_alfa = _base.TestRelatorioRTProfissional._preparar_alfa

    def _entrar_sem_carregar(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao("🔵 Aqua Gestão").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        self.assertFalse(at.exception, at.exception)
        return at

    @staticmethod
    def _uploader(at):
        return next(w for w in at.file_uploader if w.label == UPLOADER)

    @staticmethod
    def _painel(at):
        return [e for e in at.expander if e.label.startswith(PAINEL)]

    @staticmethod
    def _fotos(at):
        ss = at.session_state
        return ler_fotos(ss.filtered_state, ss["rel_origem_dados"])

    def test_selecao_antes_de_carregar_avisa_e_nao_incorpora(self):
        at = self._entrar_sem_carregar()
        self.assertFalse(at.session_state["rel_origem_dados"] if "rel_origem_dados" in at.session_state else "")
        self._uploader(at).upload("antes.png", _png((200, 10, 10)), "image/png").run()
        self.assertFalse(at.exception, at.exception)
        avisos = [w.value for w in at.warning]
        self.assertTrue(any("Carregue o condomínio e o período para anexar as fotos" in a
                            and "1 arquivo(s) selecionado(s), nenhum incorporado" in a for a in avisos), avisos)
        self.assertFalse(self._painel(at))
        self.assertFalse(at.session_state["rel_fotos_preservadas"] if "rel_fotos_preservadas" in at.session_state else "")
        # Ao carregar, o Streamlit descarta o upload (rerun antes do uploader); a foto
        # guardada entra no contexto carregado, com o painel aberto e a contagem certa.
        at.selectbox(key="sel_cliente_rel").select("Condomínio Alfa Fictício").run()
        next(b for b in at.button if b.label == "⬇️ Carregar").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(len(self._fotos(at)), 1)
        painel = self._painel(at)
        self.assertEqual(len(painel), 1)
        self.assertTrue(painel[0].proto.expanded)
        self.assertTrue(any("1 foto(s) incorporada(s) ao relatório" in i.value for i in at.info),
                        [i.value for i in at.info])
        self.assertFalse(any("Carregue o condomínio" in w.value for w in at.warning))
        at.run()  # rerun não duplica
        self.assertEqual(len(self._fotos(at)), 1)
        # Beta não recebe a foto que entrou em Alfa.
        self._carregar(at, "Residencial Beta Fictício")
        self.assertEqual(self._fotos(at), [])

    def test_miniaturas_abrem_na_hora_legenda_ilegivel_e_troca_de_condominio(self):
        self._preparar_alfa()
        at = self._entrar_e_carregar()
        self._mes(at, "09")
        self.assertFalse(self._painel(at))  # sem foto, sem painel
        self._uploader(at).upload("casa_maquinas.png", _png((10, 120, 200)), "image/png").run()
        self.assertFalse(at.exception, at.exception)
        painel = self._painel(at)
        self.assertEqual(len(painel), 1)
        self.assertEqual(painel[0].label, f"{PAINEL} (1)")
        self.assertTrue(painel[0].proto.expanded, "painel deve abrir sozinho quando chega foto nova")
        self.assertTrue(any("1 arquivo(s) selecionado(s) no upload · 1 foto(s) incorporada(s)" in i.value
                            for i in at.info), [i.value for i in at.info])
        self.assertTrue(any(True for _ in painel[0].get("imgs")), "miniatura dentro do painel")
        # Legenda: o painel continua aberto na execução seguinte e a legenda é gravada.
        next(w for w in at.text_input if w.label == "Legenda — foto 1").input("LEGENDA_FICTICIA_CMAQ").run()
        self.assertTrue(self._painel(at)[0].proto.expanded, "painel não pode fechar enquanto se digita a legenda")
        self.assertEqual(self._fotos(at)[0]["legenda"], "LEGENDA_FICTICIA_CMAQ")

        # Arquivo que não é imagem: aviso na tela, log sem nome do arquivo, nada incorporado.
        with self.assertLogs("aqua.rt_fotos", level="WARNING") as log:
            # set_value substitui a seleção: a foto válida continua selecionada junto.
            self._uploader(at).set_value([("casa_maquinas.png", _png((10, 120, 200)), "image/png"),
                                          ("documento_sigiloso.jpg", b"nao e uma imagem", "image/jpeg")]).run()
        self.assertFalse(at.exception, at.exception)
        self.assertTrue(any("Não foi possível ler a foto “documento_sigiloso.jpg”" in w.value for w in at.warning),
                        [w.value for w in at.warning])
        self.assertEqual(len(log.output), 1, log.output)
        self.assertIn("extensao=.jpg", log.output[0])
        self.assertNotIn("documento_sigiloso", log.output[0])
        self.assertEqual(len(self._fotos(at)), 1)
        self.assertTrue(any("2 arquivo(s) selecionado(s) no upload · 1 foto(s) incorporada(s)" in i.value
                            for i in at.info), [i.value for i in at.info])
        at.run()  # rerun não repete o registro no log nem duplica a foto
        self.assertEqual(len(self._fotos(at)), 1)

        # Troca de condomínio e de período: sem mistura nem duplicação.
        self._carregar(at, "Residencial Beta Fictício")
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(self._fotos(at), [])
        self.assertFalse(self._painel(at))
        self._carregar(at, "Condomínio Alfa Fictício")
        self._mes(at, "09")
        fotos = self._fotos(at)
        self.assertEqual(len(fotos), 1)
        self.assertEqual(fotos[0]["legenda"], "LEGENDA_FICTICIA_CMAQ")
        self._mes(at, "08")
        self.assertEqual(self._fotos(at), [], "foto de 09/2026 não pode aparecer em 08/2026")
        self._mes(at, "09")
        self.assertEqual(len(self._fotos(at)), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
