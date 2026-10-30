"""F1/F2/F3 no painel real via AppTest; somente cópia temporária e serviços falsos.

Executar com tests/rodar_isolado.py, que bloqueia rede e a fábrica Drive.
Não utiliza navegador real nem valida o Streamlit Cloud.
"""
import copy
import io
import json
import re
import unittest

from test_relatorio_rt_campos import AppTest, _AppReal, TestRelatorioRTCamposNoAppReal as _Aux
from utils.formatacao import slugify_nome
from utils.rt_contexto import caminho_ultimos, registrar_ultimos, ler_fotos


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestUltimosNoAppReal(_AppReal):
    _entrar_e_carregar = _Aux._entrar_e_carregar
    _mes = _Aux._mes
    _cadastrar = _Aux._cadastrar
    _visitas_alfa = _Aux._visitas_alfa
    _carregar = _Aux._carregar

    PARAMETROS = "Carregar parâmetros usados pela última vez"
    DOSAGENS = "Carregar usados pela última vez"

    @property
    def generated(self):
        return self.pasta / "app" / "generated"

    def botao(self, at, label):
        next(b for b in at.button if b.label == label).click().run()
        self.assertFalse(at.exception, at.exception)

    def snapshot(self, ph="7,31", produto="PRODUTO_FICTICIO", quantidade="2,75"):
        return {"parametros_ultimos": [{"data": "15/09/2026", "ph": ph, "cl": "1,23", "orp": "701", "temperatura": "28"}],
                "dosagens_ultimas": [{"produto": produto, "quantidade": quantidade}]}

    def campos(self, at):
        ss = at.session_state.filtered_state
        return {k: copy.deepcopy(v) for k, v in ss.items()
                if k.startswith(("rel_analise_", "rel_dos_")) or k in ("rel_origem_dados", "rel_nome_condominio", "rel_mes_referencia", "rel_ano_referencia")}

    def editar(self, at):
        at.text_input(key="rel_analise_ph_0").input("8,21").run()
        at.text_input(key="rel_dos_produto_0").input("ATUAL_FICTICIO").run()
        at.text_input(key="rel_dos_qtd_0").input("3,5").run()
        self.assertFalse(at.exception, at.exception)

    def test_f1_ambos_botoes_isolam_ids_empresas_e_slug_colidente(self):
        prefixo = "Residencial " + "A" * 115
        nome_a, nome_b = prefixo + " Alameda", prefixo + " Bosque"
        self.assertEqual(slugify_nome(nome_a), slugify_nome(nome_b))
        self._cadastrar("C31", nome_a)
        self._cadastrar("C32", nome_b)
        at = self._entrar_e_carregar(nome_b)
        self._mes(at, "10", "2026")
        self.editar(at)
        origem = at.session_state["rel_origem_dados"]
        self.assertEqual(origem, "aqua_gestao|id:C32|10/2026")
        registrar_ultimos(self.generated, "aqua_gestao|id:C31|09/2026", self.snapshot("9,91", "OUTRO_ID"))
        registrar_ultimos(self.generated, "bem_star|id:C32|09/2026", self.snapshot("9,92", "OUTRA_EMPRESA"))
        registrar_ultimos(self.generated, "aqua_gestao|id:C32|09/2026", self.snapshot())
        legado = self.generated / slugify_nome(nome_a)
        legado.mkdir(parents=True, exist_ok=True)
        (legado / "dados_condominio.json").write_text(json.dumps(self.snapshot("9,99", "LEGADO_NAO_USAR")), encoding="utf-8")
        datas = {k: v for k, v in self.campos(at).items() if k.startswith("rel_analise_data_")}
        self.botao(at, self.PARAMETROS)
        self.assertEqual(at.session_state["rel_analise_ph_0"], "7,31")
        self.assertEqual({k: at.session_state[k] for k in datas}, datas)
        self.assertTrue(any("09/2026" in w.value and "recuperados" in w.value for w in at.info))
        self.botao(at, self.DOSAGENS)
        self.assertEqual(at.session_state["rel_dos_produto_0"], "PRODUTO_FICTICIO")
        self.assertEqual(at.session_state["rel_dos_qtd_0"], "2,75")
        self.assertTrue(any("09/2026" in w.value and "recuperados" in w.value for w in at.info))
        # Um arquivo copiado para o caminho correto também precisa ter identidade válida.
        destino = caminho_ultimos(self.generated, origem)
        for chave_errada in ("bem_star|id:C32|09/2026", "aqua_gestao|id:C31|09/2026"):
            destino.write_bytes(caminho_ultimos(self.generated, chave_errada).read_bytes())
            for label in (self.PARAMETROS, self.DOSAGENS):
                with self.subTest(chave_errada=chave_errada, botao=label):
                    antes = self.campos(at)
                    self.botao(at, label)
                    self.assertEqual(self.campos(at), antes)
                    self.assertTrue(any("Não foi possível recuperar" in w.value for w in at.info))

    def test_f1_nomes_identicos_entre_empresas_bloqueiam_ambos_botoes(self):
        self._cadastrar("C41", "Residencial Sol Fictício", "Aqua Gestão")
        self._cadastrar("C42", "Residencial Sol Fictício", "Bem Star Piscinas")
        at = self._entrar_e_carregar("Residencial Sol Fictício")
        self.editar(at)
        registrar_ultimos(self.generated, "bem_star|id:C42|09/2026", self.snapshot("9,9", "BEM_NAO_USAR"))
        registrar_ultimos(self.generated, "aqua_gestao|id:C41|09/2026", self.snapshot("9,8", "AQUA_NAO_VALIDADA"))
        for label in (self.PARAMETROS, self.DOSAGENS):
            antes = self.campos(at)
            self.botao(at, label)
            self.assertEqual(self.campos(at), antes)
            self.assertTrue(any("Não foi possível recuperar" in w.value for w in at.info))

    def test_f1_ausente_invalido_e_legado_preservam_ambos_botoes(self):
        at = self._entrar_e_carregar()
        self._mes(at, "10", "2026")
        self.editar(at)
        origem = at.session_state["rel_origem_dados"]
        legado = self.generated / slugify_nome(at.session_state["rel_nome_condominio"])
        legado.mkdir(parents=True, exist_ok=True)
        (legado / "dados_condominio.json").write_text(json.dumps(self.snapshot()), encoding="utf-8")
        destino = caminho_ultimos(self.generated, origem)
        for caso in ("ausente", "json_invalido", "estrutura_invalida"):
            if caso != "ausente":
                destino.parent.mkdir(parents=True, exist_ok=True)
                destino.write_text("{" if caso == "json_invalido" else "{}", encoding="utf-8")
            for label in (self.PARAMETROS, self.DOSAGENS):
                with self.subTest(caso=caso, botao=label):
                    antes = self.campos(at)
                    self.botao(at, label)
                    self.assertEqual(self.campos(at), antes)
                    self.assertTrue(any("Não foi possível recuperar" in w.value for w in at.info))

    def test_f2_painel_rejeita_metadado_sem_perder_estado_apos_reruns(self):
        from PIL import Image
        at = self._entrar_e_carregar()
        self._mes(at, "09", "2026")
        at.text_area(key="rel_diagnostico").input("DIAGNOSTICO_ORIGINAL_FICTICIO").run()
        self.editar(at)
        foto = io.BytesIO()
        Image.new("RGB", (13, 17), "blue").save(foto, format="PNG")
        next(w for w in at.file_uploader if w.label == "Upload de fotos do relatório").upload("ficticia.png", foto.getvalue(), "image/png").run()
        self.assertFalse(at.exception, at.exception)
        origem = at.session_state["rel_origem_dados"]
        self.assertEqual(len(ler_fotos(at.session_state.filtered_state, origem)), 1)
        chave = "thyago#s#f123-invalido"
        dados = {"_chave": chave, "salvo_em": "28/10/2026 12:00:00", "_salvo_iso": "2099-01-01T00:00:00",
                 "empresa_ativa": "aqua_gestao", "rel_analises_total": "abc",
                 "campos": {"rel_nome_condominio": "OUTRO_FICTICIO", "rel_diagnostico": "NAO_APLICAR", "rel_origem_dados": origem}}
        (self.generated / "_rascunho_relatorio_rt_f123_invalido.json").write_text(json.dumps(dados), encoding="utf-8")
        if "_rt_rasc_lista_cache" in at.session_state:
            del at.session_state["_rt_rasc_lista_cache"]
        at.run()
        at.run()
        chave_botao = "btn_restaurar_rascunho_rt_" + re.sub(r"\W", "_", chave)
        self.assertTrue(any(b.key == chave_botao for b in at.button))
        def estado():
            ss = at.session_state.filtered_state
            return {k: copy.deepcopy(v) for k, v in ss.items()
                    if ((k.startswith("rel_") and k != "rel_fotos_upload")
                        or k.startswith(("_rt_rasc_", "_rt_fotos_"))
                        or k in ("_rt_upload_hashes", "_rel_fotos_origem", "empresa_ativa", "admin_empresa_fixa"))
                    and k != "_rt_rasc_lista_cache"}
        for pendente in (None, {"dados": {"rel_analises_total": "pendente-invalido", "campos": {"rel_diagnostico": "PENDENTE_ANTERIOR"}}}):
            with self.subTest(pendente=pendente):
                if pendente is not None:
                    at.session_state["_rt_rasc_pendente"] = copy.deepcopy(pendente)
                    at.run()
                antes = estado()
                cache_antes = copy.deepcopy(at.session_state["_rt_rasc_lista_cache"][1])
                next(b for b in at.button if b.key == chave_botao).click().run()
                for _ in range(3):
                    self.assertFalse(at.exception, at.exception)
                    self.assertEqual(estado(), antes)
                    self.assertEqual(at.session_state["_rt_rasc_lista_cache"][1], cache_antes)
                    self.assertIn("Total de análises inválido", at.session_state["_relatorio_rt_restore_erro"])
                    self.assertTrue(any("Não foi possível restaurar" in e.value for e in at.error))
                    self.assertEqual(len(ler_fotos(at.session_state.filtered_state, origem)), 1)
                    at.run()

    def test_f3_final_outro_periodo_ambos_botoes_datas_e_origem(self):
        at = self._entrar_e_carregar()
        self._mes(at, "09", "2026")
        at.text_input(key="rel_analise_ph_0").input("7,31").run()
        at.text_input(key="rel_dos_produto_0").input("PRODUTO_FINAL_FICTICIO").run()
        at.text_input(key="rel_dos_qtd_0").input("2,75").run()
        self.botao(at, "🚀 Gerar relatório mensal")
        origem = at.session_state["rel_origem_dados"]
        registro = caminho_ultimos(self.generated, origem)
        self.assertTrue(registro.exists(), [w.value for w in at.error] + [w.value for w in at.warning])
        publicado = registro.read_bytes()
        self.assertTrue(list(self.generated.glob("**/*RELATORIO_RT.docx")))
        self.assertTrue(list(self.generated.glob("**/*RELATORIO_RT.pdf")))
        self._mes(at, "10", "2026")
        at.text_input(key="rel_analise_data_0").input("06/10/2026").run()
        self.editar(at)
        antes = self.campos(at)
        datas = {k: v for k, v in antes.items() if k.startswith("rel_analise_data_")}
        self.assertEqual(registro.read_bytes(), publicado, "Prévia não pode substituir o último final")
        self.botao(at, self.PARAMETROS)
        self.assertEqual(at.session_state["rel_analise_ph_0"], "7,31")
        self.assertEqual({k: at.session_state[k] for k in datas}, datas)
        self.assertNotIn("29/09/2026", [at.session_state[k] for k in datas])
        self.assertTrue(any("Valores recuperados do relatório de 09/2026" in w.value for w in at.info))
        self.botao(at, self.DOSAGENS)
        self.assertEqual(at.session_state["rel_dos_produto_0"], "PRODUTO_FINAL_FICTICIO")
        self.assertEqual(at.session_state["rel_dos_qtd_0"], "2,75")
        self.assertEqual({k: at.session_state[k] for k in datas}, datas)
        self.assertTrue(any("Valores recuperados do relatório de 09/2026" in w.value for w in at.info))
        self.assertEqual(registro.read_bytes(), publicado)
        at.run()
        self.assertEqual(at.session_state["rel_analise_ph_0"], "7,31")
        self.assertEqual(at.session_state["rel_dos_produto_0"], "PRODUTO_FINAL_FICTICIO")
        self.assertFalse(at.exception, at.exception)


del _Aux
