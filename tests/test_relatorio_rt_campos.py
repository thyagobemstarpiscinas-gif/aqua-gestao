"""Regressão: dados do condomínio no relatório RT do escritório não podem sumir
após carregar um cliente e enviar fotos.

Defeito (vídeo de produção): carregar o cliente e enviar uma foto apagava nome,
endereço e representante. Cadeia comprovada:
1. O mês de referência é gravado antes do widget ser desenhado; o Streamlit não
   envia esse valor ao navegador, que mostra o campo vazio.
2. No upload, o navegador devolve o mês vazio; sem filtro de mês, a importação
   automática encontra visitas e chama st.rerun() acima dos campos.
3. Um st.rerun() no meio do script conta como execução concluída, e o Streamlit
   descarta o estado dos widgets ainda não desenhados (nome, endereço, ...).

O teste AppTest roda o app real com planilha FALSA em memória (tests/e2e/fakes),
credenciais fictícias e proxy inválido: nenhum acesso a Google Sheets, Drive ou
rede. A parte do navegador (passo 1) é coberta por tests/e2e/rt_upload_campos.js.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import json
import importlib.util
import os
import shutil
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

RAIZ = Path(__file__).resolve().parents[1]
APP = RAIZ / "app.py"
FAKES = RAIZ / "tests" / "e2e" / "fakes"
TREE = ast.parse(APP.read_text(encoding="utf-8"))

try:
    from streamlit.testing.v1 import AppTest
except Exception:  # pragma: no cover - ambiente sem streamlit
    AppTest = None

PIN_FICTICIO = "pin-ficticio-teste"
CAMPOS = ("rel_nome_condominio", "rel_endereco_condominio", "rel_representante",
          "rel_cnpj_condominio", "rel_mes_referencia", "rel_ano_referencia")


def _carregar_modulo(nome, caminho):
    spec = importlib.util.spec_from_file_location(nome, caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _modulos_falsos():
    """gspread e google.oauth2.service_account falsos, só durante o teste."""
    sa = _carregar_modulo("google.oauth2.service_account", FAKES / "google" / "oauth2" / "service_account.py")
    mods = {"gspread": _carregar_modulo("gspread", FAKES / "gspread" / "__init__.py"),
            "google.oauth2.service_account": sa}
    if "google" not in sys.modules:
        try:
            import google.oauth2  # noqa: F401
        except Exception:
            g, o = types.ModuleType("google"), types.ModuleType("google.oauth2")
            g.__path__, o.__path__ = [], []
            g.oauth2, o.service_account = o, sa
            mods.update({"google": g, "google.oauth2": o})
    return mods


class _AppReal(unittest.TestCase):
    """Cópia do app em pasta temporária, planilha falsa, credenciais fictícias, sem rede."""

    def setUp(self):
        if AppTest is not None:
            # Conexão e abas ficam em cache de recurso do processo: cada teste começa
            # com a planilha falsa dele, nunca com a de um teste anterior.
            import streamlit as st_real
            st_real.cache_data.clear()
            st_real.cache_resource.clear()
        self.pasta = Path(tempfile.mkdtemp(prefix="rt_campos_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        # Cópia do app: o que ele gravar (generated/, json locais) fica na pasta temporária.
        shutil.copytree(RAIZ, self.pasta / "app", ignore=shutil.ignore_patterns(
            ".git", "generated", "tests", "__pycache__", ".streamlit", "*.json", "backups_dev"))
        sem_rede = {k: "http://127.0.0.1:9" for k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy")}
        sem_rede.update(CODESPACES="true", NO_PROXY="localhost,127.0.0.1", no_proxy="localhost,127.0.0.1")
        p = mock.patch.dict(os.environ, sem_rede)
        p.start()
        self.addCleanup(p.stop)
        # Só as chaves falsas são trocadas e devolvidas: restaurar sys.modules inteiro
        # descarregaria módulos que o app importou (numpy etc.) e quebraria o próximo teste.
        falsos = _modulos_falsos()
        originais = {k: sys.modules.get(k) for k in falsos}
        sys.modules.update(falsos)

        def _restaurar():
            for k, v in originais.items():
                if v is None:
                    sys.modules.pop(k, None)
                else:
                    sys.modules[k] = v
        self.addCleanup(_restaurar)


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestRelatorioRTCamposNoAppReal(_AppReal):
    def _entrar_e_carregar(self, cliente="Condomínio Alfa Fictício"):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao("🔵 Aqua Gestão").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        at.selectbox(key="sel_cliente_rel").select(cliente).run()
        botao("⬇️ Carregar").click().run()
        self.assertFalse(at.exception, at.exception)
        return at

    @staticmethod
    def _valores(at):
        return {k: at.session_state[k] if k in at.session_state else None for k in CAMPOS}

    def test_upload_com_mes_vazio_nao_apaga_dados_do_condominio(self):
        at = self._entrar_e_carregar()
        carregado = self._valores(at)
        self.assertEqual(carregado["rel_nome_condominio"], "Condomínio Alfa Fictício")
        self.assertEqual(carregado["rel_endereco_condominio"], "Rua A, 1")
        self.assertEqual(carregado["rel_representante"], "Síndico Alfa")

        # O que o navegador enviava no upload: o mês que ele mostrava vazio.
        at.text_input(key="rel_mes_referencia").input("").run()
        self.assertFalse(at.exception, at.exception)
        depois = self._valores(at)
        for k in ("rel_nome_condominio", "rel_endereco_condominio", "rel_representante", "rel_cnpj_condominio"):
            self.assertEqual(depois[k], carregado[k], k)
        self.assertEqual(depois["rel_mes_referencia"], "")  # o valor enviado é respeitado, não trocado pelo padrão
        self.assertEqual(at.session_state["rel_analise_data_0"], "")  # mês vazio não importa a visita de setembro
        self.assertTrue(any("Preencha o **mês" in w.value for w in at.warning))

        # Reruns extras e edição de outro campo não alteram nada.
        at.run()
        at.text_input(key="rel_data_emissao").input("26/10/2026").run()
        at.run()
        self.assertEqual(self._valores(at), depois)


    def _mes(self, at, mes, ano=None):
        at.text_input(key="rel_mes_referencia").input(mes)
        if ano is not None:
            at.text_input(key="rel_ano_referencia").input(ano)
        at.run()
        self.assertFalse(at.exception, at.exception)
        return at.session_state["rel_analise_data_0"], at.session_state["rel_analise_ph_0"]

    def test_troca_de_mes_e_periodo_invalido(self):
        at = self._entrar_e_carregar()  # Alfa: visita só em 29/09/2026
        dados = {k: at.session_state[k] for k in CAMPOS[:4]}
        self.assertEqual(at.session_state["rel_analise_data_0"], "")  # outubro: nada importado
        self.assertEqual(self._mes(at, "09"), ("29/09/2026", "7,4"))  # setembro importado
        self.assertEqual(self._mes(at, "10"), ("", ""))  # de volta a outubro: setembro sai da tabela
        self.assertTrue(any("retiradas" in i.value for i in at.info))
        self.assertEqual(self._mes(at, "09"), ("29/09/2026", "7,4"))
        for mes, ano in (("", None), ("13", None), ("ab", None), ("09", "20"), ("09", "")):
            self.assertEqual(self._mes(at, mes, ano), ("29/09/2026", "7,4"), (mes, ano))  # nada importado nem apagado
            self.assertTrue(any("Preencha o **mês" in w.value for w in at.warning), (mes, ano))
        self._mes(at, "09", "2026")
        at.text_input(key="rel_analise_ph_0").input("7,9").run()  # edição manual
        self.assertEqual(self._mes(at, "10"), ("29/09/2026", "7,9"))  # editada: mantida, com aviso
        self.assertTrue(any("editados" in w.value for w in at.warning))
        self.assertEqual({k: at.session_state[k] for k in CAMPOS[:4]}, dados)  # dados do condomínio intactos

    def test_identificacao_duvidosa_interrompe_importacao_com_aviso(self):
        """Nomes idênticos em cadastros diferentes e cadastro indisponível: nada é importado."""
        import streamlit as st_real
        gspread = sys.modules["gspread"]
        clientes = gspread.PLANILHA.worksheet("👥 Clientes")
        linha = [""] * 15
        linha[1], linha[2], linha[8], linha[12] = "C9", "Condomínio Alfa Fictício", "Ativo", "Bem Star Piscinas"
        clientes.linhas.append(linha)  # mesmo nome, outra empresa
        st_real.cache_data.clear()
        at = self._entrar_e_carregar()  # Alfa: visita de setembro na planilha falsa
        self.assertEqual(self._mes(at, "09"), ("", ""))  # não importa
        avisos = [w.value for w in at.warning if "cadastros com o nome" in w.value]
        self.assertTrue(avisos and "C1" in avisos[0] and "C9" in avisos[0])
        self.assertEqual(at.session_state["rel_nome_condominio"], "Condomínio Alfa Fictício")  # campos intactos

        clientes.linhas.pop()
        def _falha():
            raise ConnectionError("planilha indisponível (simulado)")
        clientes.get_all_values = _falha
        st_real.cache_data.clear()
        self.assertEqual(self._mes(at, "10"), ("", ""))
        self.assertEqual(self._mes(at, "09"), ("", ""))  # setembro não entra sem confirmar o cadastro
        self.assertTrue(any("cadastro de clientes" in w.value for w in at.warning))
        del clientes.get_all_values
        st_real.cache_data.clear()
        self.assertEqual(self._mes(at, "10"), ("", ""))
        self.assertEqual(self._mes(at, "09"), ("29/09/2026", "7,4"))  # cadastro de volta: importa normalmente

    # ---- Defeito A: falha de consulta não pode apagar nem gerar com dados incompletos
    def _tabela(self, at):
        ss = at.session_state
        return {k: ss[k] for k in ("rel_analise_data_0", "rel_analise_ph_0", "rel_analise_cl_0", "rel_diagnostico",
                                   "rel_endereco_condominio", "rel_nome_condominio")}

    def _rascunho_rt_local(self):
        arquivos = sorted((self.pasta / "app" / "generated").glob("_rascunho_relatorio_rt*.json"),
                          key=lambda p: p.stat().st_mtime)
        return json.loads(arquivos[-1].read_text(encoding="utf-8"))["campos"] if arquivos else {}

    def _gerar_desabilitado(self, at):
        return next(b for b in at.button if b.label == "🚀 Gerar relatório mensal").disabled

    def _preparar_setembro_com_edicoes(self):
        at = self._entrar_e_carregar()
        self.assertEqual(self._mes(at, "09"), ("29/09/2026", "7,4"))
        at.text_area(key="rel_diagnostico").input("Diagnóstico digitado pelo RT").run()  # valor manual
        at.text_area(key="rel_endereco_condominio").input("Rua A, 1 — bloco B").run()
        self.assertFalse(at.exception, at.exception)
        return at

    def _verificar_preservado(self, at, esperado, passo):
        self.assertEqual(self._tabela(at), esperado, passo)
        self.assertFalse(at.exception, at.exception)

    def test_falha_de_leitura_da_aba_visitas_preserva_tudo_e_suspende_a_geracao(self):
        import streamlit as st_real
        at = self._preparar_setembro_com_edicoes()
        esperado = self._tabela(at)
        rascunho_antes = self._rascunho_rt_local()
        visitas = sys.modules["gspread"].PLANILHA.worksheet("🔬 Visitas")

        def _falha():
            raise ConnectionError("timeout ao ler a aba (simulado)")
        visitas.get_all_values = _falha
        st_real.cache_data.clear()
        at.run()
        self._verificar_preservado(at, esperado, "rerun com a aba indisponível")
        self.assertTrue(any("🔬 Visitas" in w.value and "falha de consulta" in w.value for w in at.warning))
        self.assertFalse(any("Nenhum lançamento de campo encontrado" in i.value for i in at.info))
        self.assertTrue(self._gerar_desabilitado(at))
        self._mes(at, "10")
        self._mes(at, "09")  # troca de mês durante a falha
        self._verificar_preservado(at, {**esperado}, "troca de mês com a aba indisponível")
        # O salvamento automático não troca a versão válida por uma vazia.
        rascunho = self._rascunho_rt_local()
        self.assertEqual((rascunho.get("rel_analise_data_0"), rascunho.get("rel_diagnostico")),
                         ("29/09/2026", "Diagnóstico digitado pelo RT"))
        self.assertEqual(rascunho_antes.get("rel_analise_ph_0"), rascunho.get("rel_analise_ph_0"))

        del visitas.get_all_values  # conexão volta
        st_real.cache_data.clear()
        at.run()
        self._verificar_preservado(at, esperado, "após a recuperação")
        self.assertFalse(self._gerar_desabilitado(at))
        self.assertFalse(any("falha de consulta" in w.value for w in at.warning))

        # Nova visita depois da recuperação, com a tabela editada pelo usuário: nada é sobrescrito.
        at.text_input(key="rel_analise_ph_0").input("7,9").run()
        linha = ["", "V2", "22/09/2026", "", "Condomínio Alfa Fictício", "7,0", "1,2", "1,4", "80", "", "0"] + [""] * 15
        visitas.linhas.append(linha)
        st_real.cache_data.clear()
        at.run()
        self.assertEqual((at.session_state["rel_analise_data_0"], at.session_state["rel_analise_ph_0"]), ("29/09/2026", "7,9"))
        self.assertTrue(any("editados" in w.value for w in at.warning))

    def test_queda_de_conexao_preserva_tudo_e_recupera(self):
        import streamlit as st_real
        at = self._preparar_setembro_com_edicoes()
        esperado = self._tabela(at)
        gspread = sys.modules["gspread"]
        autorizar = gspread.authorize

        def _sem_conexao(creds):
            raise ConnectionError("sem rede (simulado)")
        gspread.authorize = _sem_conexao
        st_real.cache_data.clear()
        st_real.cache_resource.clear()
        try:
            at.run()
            self._verificar_preservado(at, esperado, "sem conexão")
            self.assertTrue(any("cadastro de clientes" in w.value for w in at.warning))
            self.assertTrue(self._gerar_desabilitado(at))
        finally:
            gspread.authorize = autorizar
        st_real.cache_data.clear()
        st_real.cache_resource.clear()
        at.run()
        self._verificar_preservado(at, esperado, "conexão de volta")
        self.assertFalse(self._gerar_desabilitado(at))

    def test_bloqueio_por_cadastro_duplicado_preserva_analises_importadas(self):
        import streamlit as st_real
        at = self._preparar_setembro_com_edicoes()
        esperado = self._tabela(at)
        linha = [""] * 15
        linha[1], linha[2], linha[8], linha[12] = "C9", "Condomínio Alfa Fictício", "Ativo", "Bem Star Piscinas"
        sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes").linhas.append(linha)
        st_real.cache_data.clear()
        at.run()
        self._verificar_preservado(at, esperado, "nome duplicado no cadastro")
        self.assertTrue(any("cadastros com o nome" in w.value for w in at.warning))
        self.assertFalse(self._gerar_desabilitado(at))  # identificação duvidosa não impede relatório manual

    def test_nome_digitado_semelhante_a_cadastro_nao_importa_e_mantem_formulario(self):
        at = self._entrar_e_carregar()
        at.text_input(key="rel_nome_condominio").input("Alfa").run()  # abreviação digitada à mão
        self.assertEqual(self._mes(at, "09"), ("", ""))  # setembro de Alfa não entra
        self.assertTrue(any("não corresponde exatamente a um cadastro" in w.value for w in at.warning))
        self.assertEqual((at.session_state["rel_nome_condominio"], at.session_state["rel_endereco_condominio"]),
                         ("Alfa", "Rua A, 1"))

    def test_visita_do_mes_importa_e_sai_ao_trocar_de_mes(self):
        at = self._entrar_e_carregar("Residencial Beta Fictício")  # Beta: visita numa terça do mês corrente
        data0 = at.session_state["rel_analise_data_0"]
        self.assertTrue(data0.endswith(at.session_state["rel_mes_referencia"] + "/" + at.session_state["rel_ano_referencia"]))
        self.assertEqual(at.session_state["rel_analise_ph_0"], "7,2")
        mes_atual = at.session_state["rel_mes_referencia"]
        outro = "01" if mes_atual != "01" else "02"
        self.assertEqual(self._mes(at, outro), ("", ""))  # mês sem visitas: dados do mês atual saem
        self.assertEqual(self._mes(at, mes_atual), (data0, "7,2"))


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestTrocaDeCondominioNoAppReal(_AppReal):
    """Troca de condomínio com falha ao salvar o rascunho do anterior: nada é limpo."""

    def _entrar_operador(self, pin):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["operador"] = {"pin_mestre": "mestre-ficticio"}
        at.secrets["rt"] = {"pin": "rt-ficticio"}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        next(b for b in at.button if b.label == "📱 Acessar como Operador").click().run()
        at.text_input(key="op_pin_input").input(pin)
        next(b for b in at.button if b.label == "Entrar").click().run()
        return at

    def _bloquear_gravacao(self, cond):
        sys.path.insert(0, str(RAIZ))
        from utils.formatacao import slugify_nome
        caminho = self.pasta / "app" / "generated" / slugify_nome(cond)
        shutil.rmtree(caminho, ignore_errors=True)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        caminho.write_text("arquivo no lugar da pasta: a gravação do rascunho falha")
        return caminho

    @staticmethod
    def _avisos(at):
        return [w.value for w in at.warning if "foi cancelada" in w.value]

    def test_operador_troca_cancelada_mantem_campos_e_volta_a_selecao(self):
        at = self._entrar_operador("mestre-ficticio")
        a, b = at.selectbox(key="op_sel_cond").options[:2]
        at.selectbox(key="op_sel_cond").select(a).run()
        k_ph = next(t.key for t in at.text_input if t.key and t.key.startswith("op_") and t.key.endswith("_ph"))
        at.text_input(key=k_ph).input("7.4").run()
        bloqueio = self._bloquear_gravacao(a)
        at.selectbox(key="op_sel_cond").select(b).run()
        self.assertEqual(len(self._avisos(at)), 1)
        self.assertEqual(at.session_state[k_ph], "7.4")
        at.run()
        self.assertEqual((at.session_state["op_sel_cond"], at.session_state["_op_cond_form"]), (a, a))
        self.assertEqual(at.session_state[k_ph], "7.4")
        self.assertFalse(at.exception, at.exception)
        bloqueio.unlink()
        at.selectbox(key="op_sel_cond").select(b).run()
        self.assertEqual((self._avisos(at), at.session_state[k_ph], at.session_state["_op_cond_form"]), ([], "", b))

    def test_operador_com_outras_fotos_salva_rascunho_e_troca_de_condominio(self):
        """KeyError 'extras': com foto em "Outras fotos" o salvamento automático falhava
        em silêncio e a troca de condomínio ficava bloqueada."""
        sys.path.insert(0, str(RAIZ))
        from utils.formatacao import slugify_nome
        at = self._entrar_operador("mestre-ficticio")
        a, b = at.selectbox(key="op_sel_cond").options[:2]
        at.selectbox(key="op_sel_cond").select(a).run()
        pasta = self.pasta / "app" / "generated" / slugify_nome(a)
        (pasta / "fotos_rascunho").mkdir(parents=True, exist_ok=True)
        import base64  # PNG 1x1 fictício, gravado como se fosse um upload anterior
        (pasta / "fotos_rascunho" / "rasc_extras_ralo.png").write_bytes(base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=="))
        k_ph = next(t.key for t in at.text_input if t.key and t.key.startswith("op_") and t.key.endswith("_ph"))
        at.text_input(key=k_ph).input("7.4").run()  # dispara o salvamento automático
        self.assertFalse(at.exception, at.exception)
        rascunho = json.loads((pasta / "_rascunho_operador.json").read_text(encoding="utf-8"))
        self.assertEqual(rascunho["fotos_rascunho"]["extras"], ["rasc_extras_ralo.png"])
        self.assertEqual(rascunho["piscinas"][0]["ph"], "7.4")
        at.selectbox(key="op_sel_cond").select(b).run()
        self.assertEqual(self._avisos(at), [])
        self.assertEqual(at.session_state["_op_cond_form"], b)
        self.assertTrue((pasta / "fotos_rascunho" / "rasc_extras_ralo.png").exists())

    def test_rt_troca_cancelada_mantem_dados_e_rascunho_volta_depois(self):
        at = self._entrar_operador("rt-ficticio")
        a, b = at.selectbox(key="rt_op_sel_cond").options[:2]
        at.text_area(key="rel_parecer_complemento_rt").input("Parecer fictício").run()
        bloqueio = self._bloquear_gravacao(a)
        at.selectbox(key="rt_op_sel_cond").select(b).run()
        self.assertEqual(len(self._avisos(at)), 1)
        at.run()
        self.assertEqual(at.session_state["rt_op_sel_cond"], a)
        self.assertEqual(at.session_state["rel_parecer_complemento_rt"], "Parecer fictício")
        bloqueio.unlink()
        at.selectbox(key="rt_op_sel_cond").select(b).run()
        self.assertEqual(at.session_state["rel_parecer_complemento_rt"], "")
        at.selectbox(key="rt_op_sel_cond").select(a).run()
        self.assertEqual(at.session_state["rel_parecer_complemento_rt"], "Parecer fictício")
        self.assertFalse(at.exception, at.exception)

    def test_rt_formulario_completo_volta_intacto_apos_trocar_e_retornar(self):
        at = self._entrar_operador("rt-ficticio")
        opcoes = [o for o in at.selectbox(key="rt_op_sel_cond").options if o.upper() != "TODOS"]
        a, b = opcoes[:2]
        at.selectbox(key="rt_op_sel_cond").select(a).run()
        # Pela tela: identificação, ART, análises, parecer e legenda.
        at.text_area(key="rel_endereco_condominio").input("Rua Fictícia, 10").run()
        at.selectbox(key="rel_art_modo_rt2025").select("Informar número da ART").run()
        at.text_input(key="rel_art_numero_rt2025").input("12345").run()
        for chave, valor in (("rel_analise_data_0", "06/10/2026"), ("rel_analise_ph_0", "7,4"),
                             ("rel_analise_cl_0", "1,5"), ("rel_analise_cya_0", "0"), ("rel_foto_legenda_cmaq", "Casa A")):
            at.text_input(key=chave).input(valor)
        at.text_area(key="rel_parecer_complemento_rt").input("Parecer fictício").run()
        # Preenchidos por importação/escritório (sem widget nesta tela), também vão para o PDF.
        for chave, valor in (("rel_dos_produto_0", "Cloro granulado"), ("rel_dos_qtd_0", "2"),
                             ("rel_observacoes_gerais", "Observação geral"), ("rel_diagnostico", "Diagnóstico")):
            at.session_state[chave] = valor
        at.run()
        chaves = ("rel_endereco_condominio", "rel_cnpj_condominio", "rel_representante", "rel_art_numero",
                  "rel_art_status", "rel_analise_data_0", "rel_analise_ph_0", "rel_analise_cl_0", "rel_analise_cya_0",
                  "rel_foto_legenda_cmaq", "rel_parecer_complemento_rt", "rel_dos_produto_0", "rel_dos_qtd_0",
                  "rel_observacoes_gerais", "rel_diagnostico", "rel_mes_referencia", "rel_ano_referencia")
        antes = {k: at.session_state[k] for k in chaves}
        self.assertEqual((antes["rel_art_status"], antes["rel_dos_produto_0"]), ("Emitida", "Cloro granulado"))
        at.selectbox(key="rt_op_sel_cond").select(b).run()
        self.assertEqual(self._avisos(at), [])
        for k in ("rel_analise_ph_0", "rel_parecer_complemento_rt", "rel_dos_produto_0", "rel_observacoes_gerais",
                  "rel_diagnostico", "rel_foto_legenda_cmaq"):
            self.assertEqual(at.session_state[k] if k in at.session_state else "", "", k)  # nada de A em B
        self.assertNotEqual(at.session_state["rel_endereco_condominio"], "Rua Fictícia, 10")
        at.selectbox(key="rt_op_sel_cond").select(a).run()
        at.run()
        self.assertEqual({k: at.session_state[k] for k in chaves}, antes)
        self.assertFalse(at.exception, at.exception)


class TestPeriodoEImportacao(unittest.TestCase):
    def _ns(self, ss):
        nomes = ("periodo_relatorio_valido", "_valores_importaveis_rt", "_preenchidos", "tabela_rt_livre_para_importar",
                 "registrar_importacao_rt", "retirar_importacao_rt", "_filtrar_mes")
        nos = [n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name in nomes]
        ns = {"st": types.SimpleNamespace(session_state=ss),
              "lancamento_pertence_mes_ano": lambda data, mes, ano: data.endswith(f"/{int(mes):02d}/{ano}")}
        exec(compile(ast.Module(body=nos, type_ignores=[]), str(APP), "exec"), ns)
        return ns

    def test_periodo_valido(self):
        f = self._ns({})["periodo_relatorio_valido"]
        for mes, ano in (("10", "2026"), ("1", "2026"), ("09", "2025"), (" 12 ", "2026")):
            self.assertTrue(f(mes, ano), (mes, ano))
        for mes, ano in (("", "2026"), ("10", ""), (None, None), ("0", "2026"), ("13", "2026"), ("1a", "2026"),
                         ("10", "26"), ("10", "1999"), ("10", "abcd"), ("-1", "2026")):
            self.assertFalse(f(mes, ano), (mes, ano))

    def test_mes_vazio_nao_significa_todos_os_meses(self):
        filtrar = self._ns({})["_filtrar_mes"]
        visitas = [{"data": "29/09/2026"}, {"data": "06/10/2026"}]
        self.assertEqual(filtrar(visitas, "", "2026"), [])
        self.assertEqual(filtrar(visitas, "10", ""), [])
        self.assertEqual(filtrar(visitas, "10", "2026"), [{"data": "06/10/2026"}])

    def test_tabela_livre_so_vazia_ou_igual_a_ultima_importacao(self):
        ss = {"rel_analise_ph_0": "", "rel_dos_produto_0": "", "rel_observacoes_gerais": ""}
        ns = self._ns(ss)
        self.assertTrue(ns["tabela_rt_livre_para_importar"]())
        ss.update(rel_analise_ph_0="7,4", rel_analise_data_0="29/09/2026")
        self.assertFalse(ns["tabela_rt_livre_para_importar"]())  # digitado, sem importação registrada
        ns["registrar_importacao_rt"]("09/2026")
        self.assertTrue(ns["tabela_rt_livre_para_importar"]())
        ss["rel_dos_produto_0"] = "Cloro"
        self.assertFalse(ns["tabela_rt_livre_para_importar"]())  # editado depois da importação
        ss["rel_dos_produto_0"] = ""
        self.assertEqual(ns["retirar_importacao_rt"](), "09/2026")
        self.assertEqual((ss["rel_analise_ph_0"], ss["rel_analise_data_0"]), ("", ""))
        self.assertNotIn("_rel_autoimport_snapshot", ss)


class TestPreservarCamposRT(unittest.TestCase):
    def _ns(self, ss):
        nos = [n for n in TREE.body if isinstance(n, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in ("_RT_CAMPOS_NAO_REAFIRMAR", "_RASCUNHO_RT_NAO_RESTAURAR")
            for t in n.targets)]
        nos.append(next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == "preservar_campos_rt"))
        ns = {"st": types.SimpleNamespace(session_state=ss)}
        exec(compile(ast.Module(body=nos, type_ignores=[]), str(APP), "exec"), ns)
        return ns

    def test_regrava_campos_rel_e_pula_uploads_e_widgets_com_indice(self):
        class _SS(dict):
            def __init__(self, *a):
                super().__init__(*a)
                self.gravados = []

            def __setitem__(self, k, v):
                if k == "rel_ja_desenhado":
                    raise RuntimeError("widget já criado nesta execução")
                self.gravados.append(k)
                super().__setitem__(k, v)

        ss = _SS({"rel_nome_condominio": "Alfa", "rel_mes_referencia": "10", "rel_analise_ph_0": "0",
                  "rel_fotos_upload": ["foto"], "rel_pdf_visita_importador": None,
                  "rel_status_agua_select": "CONFORME", "rel_ja_desenhado": "x",
                  "nome_condominio": "fora do relatório", "op_adulto_ph": "7.4", "csr_nome": "Bem Star"})
        antes = dict(ss)
        self._ns(ss)["preservar_campos_rt"]()
        self.assertEqual(sorted(ss.gravados), ["rel_analise_ph_0", "rel_mes_referencia", "rel_nome_condominio"])
        self.assertEqual(dict(ss), antes)  # nenhum valor muda


class TestLigacaoNoApp(unittest.TestCase):
    def test_preservacao_antes_dos_widgets_e_sem_rerun_na_importacao_automatica(self):
        src = APP.read_text(encoding="utf-8")
        i_sub = src.index('st.subheader("Relatório mensal de responsabilidade técnica")')
        i_pres = src.index("preservar_campos_rt()", i_sub)
        self.assertLess(i_pres, src.index('key="sel_cliente_rel"', i_sub))
        self.assertLess(i_pres, src.index('key="rel_nome_condominio"', i_sub))
        i_auto = src.index('if st.session_state.get("_rel_autoimport_assinatura") != _assinatura_auto:')
        self.assertNotIn("st.rerun()", src[i_auto:src.index("except Exception", i_auto)])
        self.assertIn("if _periodo_rel_ok and nome_rel_atual:", src)
        self.assertIn("elif lancamentos_disponiveis:\n                _importar_lancamentos", src)
        i_cli = src.index('if st.session_state.pop("_rel_auto_importar_cliente", False) and _periodo_rel_ok:')
        self.assertNotIn("st.rerun()", src[i_cli:src.index('st.markdown("**Dados do condomínio / local atendido**")', i_cli)])


if __name__ == "__main__":
    unittest.main()
