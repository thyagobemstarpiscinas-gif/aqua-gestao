"""Testes do nome seguro dos downloads (st.download_button).

Carregam o bloco "NOME SEGURO PARA DOWNLOADS" do app.py com um st simulado,
sem executar o app nem acessar rede, Google Sheets ou Drive.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import re
import types
import unicodedata
import unittest
from pathlib import Path

APP = Path(__file__).resolve().parents[1] / "app.py"
TREE = ast.parse(APP.read_text(encoding="utf-8"))


def _bloco_download():
    """Nós de nível superior desde nome_download_seguro até `st.download_button = ...`."""
    nos, dentro = [], False
    for n in TREE.body:
        if isinstance(n, ast.FunctionDef) and n.name == "nome_download_seguro":
            dentro = True
        if dentro:
            nos.append(n)
            if isinstance(n, ast.Assign) and any(
                isinstance(t, ast.Attribute) and t.attr == "download_button" for t in n.targets
            ):
                return nos
    raise AssertionError("bloco do wrapper não encontrado")


BLOCO = compile(ast.Module(body=_bloco_download(), type_ignores=[]), str(APP), "exec")


def _st_falso():
    chamadas = []

    def download_button(*args, **kwargs):
        chamadas.append((args, kwargs))
        return False

    return types.SimpleNamespace(download_button=download_button), chamadas


def _carregar(st, vezes=1):
    ns = {"st": st, "re": re, "unicodedata": unicodedata}
    for _ in range(vezes):  # cada rerun do Streamlit reexecuta o script inteiro
        exec(BLOCO, ns)
    return ns


class TestNomeDownloadSeguro(unittest.TestCase):
    def setUp(self):
        self.f = _carregar(_st_falso()[0])["nome_download_seguro"]

    def test_acentos_preservam_extensao(self):
        casos = {
            "20261025_101637_Condomínio_Mozaic_PREVIA_EXATA_RELATORIO_RT.pdf":
                "20261025_101637_Condominio_Mozaic_PREVIA_EXATA_RELATORIO_RT.pdf",
            "20261025_101637_Condomínio_Mozaic_PREVIA_EXATA_RELATORIO_RT.docx":
                "20261025_101637_Condominio_Mozaic_PREVIA_EXATA_RELATORIO_RT.docx",
            "Residencial São João — Relatório.pdf": "Residencial_Sao_Joao_Relatorio.pdf",
            "Aditivo RT - Edifício Açaí.docx": "Aditivo_RT_-_Edificio_Acai.docx",
            "mockup_relatorio_Aqua_Gestão_tela.html": "mockup_relatorio_Aqua_Gestao_tela.html",
        }
        for entrada, esperado in casos.items():
            self.assertEqual(self.f(entrada), esperado)

    def test_espacos_e_caracteres_especiais(self):
        self.assertEqual(self.f("Contrato limpeza: Ed. Ômega/Bloco B.pdf"), "Contrato_limpeza_Ed._Omega_Bloco_B.pdf")
        self.assertEqual(self.f("BEM STAR – Visita 03/10.pdf"), "BEM_STAR_Visita_03_10.pdf")
        self.assertEqual(self.f('Termo "Síndico" <2026>?.docx'), "Termo_Sindico_2026_.docx")

    def test_nomes_ja_seguros_nao_mudam(self):
        for nome in ("Proposta_Comercial.pdf", "MOZAIC_8507_CONFERENCIA.pdf", "relatorio.final.v2.pdf",
                     "arquivo.PDF", "operadores_permissoes_20261025_1016.csv", "edificio_aurora_cadastro.json"):
            self.assertEqual(self.f(nome), nome)

    def test_nome_vazio_ou_ausente(self):
        for nome in ("", None, "   ", "___"):
            self.assertEqual(self.f(nome), "arquivo")

    def test_nome_que_viraria_so_extensao_ganha_nome_base(self):
        casos = {"日本.pdf": "arquivo.pdf", "Ωμέγα.docx": "arquivo.docx", "–.pdf": "arquivo.pdf",
                 ".pdf": "arquivo.pdf", "-_.docx": "arquivo.docx", "????.pdf": "arquivo.pdf", "日本": "arquivo"}
        for entrada, esperado in casos.items():
            self.assertEqual(self.f(entrada), esperado, entrada)

    def test_resultado_sempre_ascii_seguro_e_estavel(self):
        amostra = ["Condomínio.pdf", "日本.docx", "a b/c\\d.pdf", "ção", "", None, "x.y.z", "Ñandú – 2026.docx"]
        for nome in amostra:
            r = self.f(nome)
            r.encode("ascii")
            self.assertRegex(r, r"^[A-Za-z0-9._-]+$")
            self.assertFalse(r.startswith("."), r)
            self.assertEqual(self.f(r), r)  # aplicar de novo não muda


class TestWrapper(unittest.TestCase):
    def test_wrapper_sanitiza_so_o_nome_e_preserva_o_resto(self):
        st, chamadas = _st_falso()
        _carregar(st)
        dados = b"%PDF-1.4 conteudo ficticio"
        st.download_button("Baixar PDF", data=dados, file_name="Relatório Condomínio.pdf",
                           mime="application/pdf", key="k1", use_container_width=True)
        args, kwargs = chamadas[-1]
        self.assertEqual(args, ("Baixar PDF",))
        self.assertEqual(kwargs["file_name"], "Relatorio_Condominio.pdf")
        self.assertIs(kwargs["data"], dados)  # conteúdo intocado (mesmo objeto)
        self.assertEqual((kwargs["mime"], kwargs["key"], kwargs["use_container_width"]),
                         ("application/pdf", "k1", True))

    def test_sem_file_name_nao_altera_nada(self):
        st, chamadas = _st_falso()
        _carregar(st)
        st.download_button("x", data=b"1")
        st.download_button("y", data=b"2", file_name=None)
        self.assertNotIn("file_name", chamadas[0][1])
        self.assertIsNone(chamadas[1][1]["file_name"])

    def test_reruns_nao_empilham_wrapper(self):
        st, chamadas = _st_falso()
        original = st.download_button
        _carregar(st, vezes=10)
        self.assertIs(st.download_button._original_aqua, original)
        self.assertFalse(hasattr(original, "_original_aqua"))
        st.download_button("z", data=b"3", file_name="Ação.docx")
        self.assertEqual(len(chamadas), 1)  # uma única chamada ao original
        self.assertEqual(chamadas[0][1]["file_name"], "Acao.docx")


class TestCoberturaNoApp(unittest.TestCase):
    def _chamadas(self):
        return [n for n in ast.walk(TREE) if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                and n.func.attr == "download_button"]

    def test_todas_as_chamadas_passam_pelo_wrapper(self):
        chamadas = self._chamadas()
        self.assertGreaterEqual(len(chamadas), 40)
        for c in chamadas:
            self.assertTrue(isinstance(c.func.value, ast.Name) and c.func.value.id == "st",
                            f"linha {c.lineno}: download_button fora de st. não passa pelo wrapper")
            self.assertIn("file_name", {k.arg for k in c.keywords}, f"linha {c.lineno}: file_name não nomeado")
            self.assertLessEqual(len(c.args), 2, f"linha {c.lineno}: file_name posicional")

    def test_wrapper_instalado_antes_da_primeira_chamada(self):
        linha_wrapper = _bloco_download()[-1].lineno
        self.assertLess(linha_wrapper, min(c.lineno for c in self._chamadas()))


if __name__ == "__main__":
    unittest.main()
