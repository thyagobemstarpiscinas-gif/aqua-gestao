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


@unittest.skipIf(AppTest is None, "streamlit.testing indisponível")
class TestRelatorioRTCamposNoAppReal(unittest.TestCase):
    def setUp(self):
        self.pasta = Path(tempfile.mkdtemp(prefix="rt_campos_"))
        self.addCleanup(shutil.rmtree, self.pasta, ignore_errors=True)
        # Cópia do app: o que ele gravar (generated/, json locais) fica na pasta temporária.
        shutil.copytree(RAIZ, self.pasta / "app", ignore=shutil.ignore_patterns(
            ".git", "generated", "tests", "__pycache__", ".streamlit", "*.json", "backups_dev"))
        sem_rede = {k: "http://127.0.0.1:9" for k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy")}
        sem_rede.update(CODESPACES="true", NO_PROXY="localhost,127.0.0.1", no_proxy="localhost,127.0.0.1")
        for p in (mock.patch.dict(os.environ, sem_rede), mock.patch.dict(sys.modules, _modulos_falsos())):
            p.start()
            self.addCleanup(p.stop)

    def _entrar_e_carregar(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)  # noqa: E731
        botao("🔐 Acesso administrativo").click().run()
        botao("🔵 Aqua Gestão").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        at.selectbox(key="sel_cliente_rel").select("Condomínio Alfa Fictício").run()
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

        # Reruns extras e edição de outro campo não alteram nada.
        at.run()
        at.text_input(key="rel_data_emissao").input("26/10/2026").run()
        at.run()
        self.assertEqual(self._valores(at), depois)


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
        i_cli = src.index('if st.session_state.pop("_rel_auto_importar_cliente", False):')
        self.assertNotIn("st.rerun()", src[i_cli:src.index('st.markdown("**Dados do condomínio / local atendido**")', i_cli)])


if __name__ == "__main__":
    unittest.main()
