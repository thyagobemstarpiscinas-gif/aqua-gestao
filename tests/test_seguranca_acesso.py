"""Testes de segurança do acesso por PIN (operador, RT e administrativo).

Carregam somente as funções de autenticação do app.py, sem executar o app,
com Secrets e planilha simulados. Não acessam Google Sheets, Drive nem rede.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import hmac
import re
import unittest
import unicodedata
from datetime import datetime
from pathlib import Path

APP = Path(__file__).resolve().parents[1] / "app.py"
SRC = APP.read_text(encoding="utf-8")
TREE = ast.parse(SRC)

FUNCOES = [
    "_normalizar_chave_acesso", "_condominios_organizar", "_resolver_condominios_permitidos_exatos",
    "_secret_pin", "_pin_confere", "_pin_mestre_operador", "_pin_reservado", "validar_pin_operador",
    "sheets_listar_operadores", "_op_tem_acesso_total",
    "_admin_pin_configurado", "_admin_pin_valido", "_admin_nome_empresa", "_admin_icone_empresa",
    "_admin_entrar_empresa", "_admin_sessao_valida",
]


def _funcao(nome):
    for n in TREE.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            return n
    raise AssertionError(f"Função {nome} não encontrada no app.py")


class _Secrets(dict):
    """Imita st.secrets: .get() de seções que se comportam como dict."""


class _FakeSt:
    def __init__(self):
        self.secrets = _Secrets()
        self.session_state = {}


class _Aba:
    def __init__(self, linhas):
        self._linhas = linhas

    def get_all_values(self):
        return self._linhas


def _carregar(secrets=None, linhas_operadores=None, operadores_locais=None):
    """Monta um namespace isolado com as funções reais do app.py."""
    st = _FakeSt()
    st.secrets.update(secrets or {})
    ns = {
        "st": st, "re": re, "hmac": hmac, "unicodedata": unicodedata, "datetime": datetime,
        "conectar_sheets": lambda: object(),
        "obter_aba_sheets": lambda nome: _Aba(linhas_operadores or []),
        "_log_sheets_erro": lambda *a, **k: None,
        "carregar_operadores": lambda: list(operadores_locais or []),
    }
    try:
        from utils.normalizacao import nomes_condominio_equivalentes, normalizar_texto_busca
        ns.update(nomes_condominio_equivalentes=nomes_condominio_equivalentes,
                  normalizar_texto_busca=normalizar_texto_busca)
    except Exception:
        pass
    for nome in FUNCOES:
        f = _funcao(nome)
        f.decorator_list = []  # remove @st.cache_data
        exec(compile(ast.Module(body=[f], type_ignores=[]), str(APP), "exec"), ns)
    return ns


CABECALHO = ["Nome", "PIN", "Condomínios", "Ativo"]
OPERADORES = [
    CABECALHO,
    ["Ana Campo", "4821", "Condomínio Alfa", "Sim"],
    ["Beto Total", "7350", "TODOS", "Sim"],
    ["Caio Vazio", "6194", "", "Sim"],
    ["Dora Inativa", "2587", "Condomínio Alfa", "Não"],
    ["Sem PIN", "", "TODOS", "Sim"],
]
SECRETS_OK = {
    "operador": {"pin_mestre": "mestre-teste-1"},
    "rt": {"pin": "rt-teste-2"},
    "admin": {"pin": "admin-teste-3"},
}


class TestSemSecrets(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(secrets={}, linhas_operadores=OPERADORES)

    def test_pin_mestre_desativado(self):
        self.assertEqual(self.ns["_pin_mestre_operador"](), "")
        for tentativa in ["", " ", "0000", "1234", "admin", "PIN_MESTRE", "mestre-teste-1"]:
            op = self.ns["validar_pin_operador"](tentativa)
            self.assertFalse(op and op.get("nome") == "Operador" and op.get("acesso_total"),
                             f"PIN mestre aceito sem Secret: {tentativa!r}")

    def test_admin_desativado(self):
        self.assertEqual(self.ns["_admin_pin_configurado"](), "")
        for tentativa in ["", " ", "admin", "1234", "admin-teste-3"]:
            self.assertFalse(self.ns["_admin_pin_valido"](tentativa))

    def test_rt_desativado(self):
        self.assertIsNone(self.ns["validar_pin_operador"]("rt-teste-2"))

    def test_secrets_indisponivel_nao_quebra(self):
        ns = _carregar()

        class _Quebrado:
            def get(self, *a, **k):
                raise FileNotFoundError("sem secrets.toml")

        ns["st"].secrets = _Quebrado()
        self.assertEqual(ns["_secret_pin"]("admin"), "")
        self.assertFalse(ns["_admin_pin_valido"]("qualquer"))


class TestComSecrets(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(secrets=SECRETS_OK, linhas_operadores=OPERADORES)

    def test_pin_mestre_valido(self):
        op = self.ns["validar_pin_operador"]("mestre-teste-1")
        self.assertTrue(op["acesso_total"])
        self.assertEqual(op["condomínios"], ["TODOS"])
        self.assertNotIn("is_rt", op)

    def test_rt_valido(self):
        op = self.ns["validar_pin_operador"]("  rt-teste-2  ")
        self.assertTrue(op["is_rt"])
        self.assertTrue(op["acesso_total"])

    def test_admin_valido(self):
        self.assertTrue(self.ns["_admin_pin_valido"]("admin-teste-3"))
        self.assertFalse(self.ns["_admin_pin_valido"]("admin-teste-"))
        self.assertFalse(self.ns["_admin_pin_valido"]("ADMIN-TESTE-3"))

    def test_pin_admin_nao_entra_no_modo_operador(self):
        self.assertIsNone(self.ns["validar_pin_operador"]("admin-teste-3"))

    def test_pin_reservado(self):
        for pin in ("mestre-teste-1", "rt-teste-2", "admin-teste-3"):
            self.assertTrue(self.ns["_pin_reservado"](pin))
        self.assertFalse(self.ns["_pin_reservado"]("4821"))
        self.assertFalse(self.ns["_pin_reservado"](""))


class TestOperadores(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(secrets={}, linhas_operadores=OPERADORES)
        self.validar = self.ns["validar_pin_operador"]

    def test_pin_invalido(self):
        for pin in ["", "   ", "0000", "48210", "482", "4821 x"]:
            self.assertIsNone(self.validar(pin), pin)

    def test_pin_vazio_nao_casa_com_operador_sem_pin(self):
        self.assertIsNone(self.validar(""))

    def test_operador_restrito_ve_apenas_seus_condominios(self):
        op = self.validar("4821")
        self.assertEqual(op["nome"], "Ana Campo")
        self.assertFalse(op["acesso_total"])
        todos = ["Condomínio Alfa", "Condomínio Beta", "Residencial Gama"]
        liberados = self.ns["_resolver_condominios_permitidos_exatos"](op["condomínios"], todos)
        self.assertEqual(liberados, ["Condomínio Alfa"])

    def test_operador_todos_explicito(self):
        self.assertTrue(self.validar("7350")["acesso_total"])

    def test_lista_vazia_nao_libera_acesso_total(self):
        op = self.validar("6194")
        self.assertIsNotNone(op)
        self.assertFalse(op["acesso_total"])
        self.assertEqual(op["condomínios"], [])
        self.assertFalse(self.ns["_op_tem_acesso_total"](op))

    def test_operador_inativo_negado(self):
        self.assertIsNone(self.validar("2587"))

    def test_fallback_json_local_respeita_regras(self):
        ns = _carregar(secrets={}, linhas_operadores=[], operadores_locais=[
            {"nome": "Local", "pin": "9911", "condomínios": [], "ativo": True},
            {"nome": "Local Sem PIN", "pin": "", "condomínios": ["TODOS"], "ativo": True},
        ])
        op = ns["validar_pin_operador"]("9911")
        self.assertFalse(op["acesso_total"])
        self.assertIsNone(ns["validar_pin_operador"](""))


class TestSeparacaoEmpresas(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(secrets=SECRETS_OK)
        self.ss = self.ns["st"].session_state

    def test_escritorio_sem_login_volta_para_entrada(self):
        self.ss.update(modo_atual="escritorio", empresa_ativa="bem_star")
        self.assertFalse(self.ns["_admin_sessao_valida"]())
        self.assertEqual(self.ss["modo_atual"], "entrada")
        self.assertFalse(self.ss.get("admin_logado", False))

    def test_login_bem_star_fica_travado_em_bem_star(self):
        self.ns["_admin_entrar_empresa"]("bem_star")
        self.ss["empresa_ativa"] = "aqua_gestao"  # widget antigo tentando trocar
        self.assertTrue(self.ns["_admin_sessao_valida"]())
        self.assertEqual(self.ss["empresa_ativa"], "bem_star")

    def test_login_aqua_fica_travado_em_aqua(self):
        self.ns["_admin_entrar_empresa"]("aqua_gestao")
        self.ss["empresa_ativa"] = "bem_star"
        self.ns["_admin_sessao_valida"]()
        self.assertEqual(self.ss["empresa_ativa"], "aqua_gestao")

    def test_empresa_desconhecida_vira_aqua(self):
        self.ns["_admin_entrar_empresa"]("outra")
        self.assertEqual(self.ss["empresa_ativa"], "aqua_gestao")


class TestSemCredencialFixaNoCodigo(unittest.TestCase):
    """Garante que nenhum PIN/senha literal volte a autenticar."""

    def test_sem_constante_pin_operador(self):
        nomes = {t.id for n in TREE.body if isinstance(n, ast.Assign) for t in n.targets if isinstance(t, ast.Name)}
        self.assertNotIn("PIN_OPERADOR", nomes)

    def test_funcoes_de_login_sem_string_fixa_comparada_ou_retornada(self):
        for nome in ("validar_pin_operador", "_admin_pin_configurado", "_admin_pin_valido", "_pin_mestre_operador"):
            f = _funcao(nome)
            for no in ast.walk(f):
                if isinstance(no, ast.Return) and isinstance(no.value, ast.Constant) and isinstance(no.value.value, str):
                    self.assertEqual(no.value.value, "", f"{nome} retorna string fixa")
                if isinstance(no, ast.Compare):
                    for lado in [no.left, *no.comparators]:
                        # "todos" é a marcação de escopo de condomínios, não uma credencial
                        if (isinstance(lado, ast.Constant) and isinstance(lado.value, str)
                                and lado.value and lado.value != "todos"):
                            self.fail(f"{nome} compara com string fixa")


if __name__ == "__main__":
    unittest.main()
