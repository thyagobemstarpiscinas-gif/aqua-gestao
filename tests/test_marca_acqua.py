"""Regressão de identidade visual: ACQUA GESTÃO (dois C) e Bem Star separadas.

Gera os documentos reais com dados fictícios (sem rede, sem Sheets) e confere,
pelas imagens incorporadas, que:
- documentos ACQUA usam somente a logo aprovada acqua_gestao_logo_rt.png;
- nenhum documento usa a arte antiga "AQUA GESTÃO";
- documentos Bem Star usam somente a marca Bem Star.
Também verifica o texto do nome-marca (a razão social com CNPJ não muda).
"""
import ast
import re
import subprocess
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.marca import LOGOS_AQUA_ANTIGOS_SHA256, logo_acqua, sha256, substituir_logo_antigo_docx  # noqa: E402
from utils.verificar_marca import marcas_do_arquivo  # noqa: E402

ACQUA = [
    "02_visita_operador_acqua.pdf", "03_visita_supervisao_rt.pdf", "05_mensal_rt_app.pdf",
    "06_consolidado_rt.pdf", "08_relatorio_mensal_textos.docx", "09_mensal_rt_premium.pdf",
    "10_mensal_rt.docx", "11_consolidado_semestral.pdf", "12_consolidado_anual.pdf",
    "13_contrato_rt.pdf", "14_proposta_crm_acqua.pdf", "15_contrato_crm_acqua.pdf",
    "16_dossie_fds_ghs.pdf", "17_caderno_pops.pdf", "18_termo_ciencia_sindico.pdf",
    "20_contrato_rt.docx", "21_aditivo_rt.docx", "23_aditivo_rt_premium.pdf",
    "24_previa_relatorio_acqua.html", "25_previa_impressao_acqua.html", "27_visita_html_acqua.html",
    "28_contrato_limpeza_acqua.pdf", "29_termo_ciencia_operador.pdf",
]
BEM_STAR = [
    "01_visita_bem_star.pdf", "07_mensal_bem_star.pdf", "19_proposta_crm_bem_star.pdf",
    "22_contrato_bem_star.docx", "30_contrato_crm_bem_star.pdf",
]
# Nome-marca antigo, exceto a razão social registrada.
MARCA_ANTIGA = re.compile(r"(?<![A-Za-z])AQUA GEST[ÃA]O(?! CONTROLE T[ÉE]CNICO LTDA)")


def _texto(caminho: Path) -> str:
    if caminho.suffix == ".pdf":
        import fitz
        with fitz.open(caminho) as doc:
            return " ".join(" ".join(p.get_text() for p in doc).split())
    if caminho.suffix == ".docx":
        import zipfile
        with zipfile.ZipFile(caminho) as z:
            xml = " ".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist()
                           if n.startswith("word/") and n.endswith(".xml"))
        return re.sub(r"<[^>]+>", "", xml)
    return re.sub(r"<[^>]+>", " ", caminho.read_text(encoding="utf-8"))


class DocumentosGerados(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.saida = Path(cls._tmp.name)
        r = subprocess.run([sys.executable, str(ROOT / "tests" / "gerar_documentos_ficticios.py"), str(cls.saida)],
                           cwd=ROOT, capture_output=True, text=True, timeout=600)
        cls.log = r.stdout + r.stderr
        cls.codigo = r.returncode

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_todos_os_documentos_foram_gerados(self):
        self.assertEqual(self.codigo, 0, self.log[-3000:])
        for nome in ACQUA + BEM_STAR:
            self.assertTrue((self.saida / nome).is_file(), nome)

    def test_documentos_acqua_usam_somente_a_logo_aprovada(self):
        for nome in ACQUA:
            with self.subTest(documento=nome):
                marcas = Counter(marcas_do_arquivo(self.saida / nome))
                self.assertGreater(marcas["acqua"], 0, f"{nome}: logo ACQUA aprovada ausente ({dict(marcas)})")
                self.assertEqual(marcas["aqua_antiga"], 0, f"{nome}: arte antiga AQUA ({dict(marcas)})")
                self.assertEqual(marcas["bem_star"], 0, f"{nome}: marca Bem Star indevida ({dict(marcas)})")

    def test_documentos_bem_star_usam_somente_a_propria_marca(self):
        for nome in BEM_STAR:
            with self.subTest(documento=nome):
                marcas = Counter(marcas_do_arquivo(self.saida / nome))
                self.assertGreater(marcas["bem_star"], 0, f"{nome}: logo Bem Star ausente ({dict(marcas)})")
                self.assertEqual(marcas["acqua"] + marcas["aqua_antiga"], 0, f"{nome}: marca ACQUA indevida")

    def test_nome_marca_com_grafia_aprovada(self):
        for nome in ACQUA:
            with self.subTest(documento=nome):
                antigos = MARCA_ANTIGA.findall(_texto(self.saida / nome))
                self.assertEqual(antigos, [], f"{nome}: grafia antiga do nome-marca")
        for nome in BEM_STAR:
            with self.subTest(documento=nome):
                self.assertNotIn("ACQUA GEST", _texto(self.saida / nome))


class FonteSemArteAntiga(unittest.TestCase):
    SRC = (ROOT / "app.py").read_text(encoding="utf-8")

    def test_encontrar_logo_retorna_apenas_a_arte_aprovada(self):
        tree = ast.parse(self.SRC)
        ns = {"BASE_DIR": ROOT, "Path": Path}
        for node in tree.body:
            if (isinstance(node, ast.FunctionDef) and node.name == "encontrar_logo") or (
                    isinstance(node, ast.Assign) and any(getattr(t, "id", "") == "LOGO_CANDIDATOS" for t in node.targets)):
                exec(compile(ast.Module(body=[node], type_ignores=[]), "app.py", "exec"), ns)
        self.assertEqual(ns["encontrar_logo"](), ROOT / "acqua_gestao_logo_rt.png")
        self.assertEqual(ns["LOGO_CANDIDATOS"], [ROOT / "acqua_gestao_logo_rt.png"])

    def test_nenhum_gerador_aponta_para_arte_antiga(self):
        for arquivo in [ROOT / "app.py", ROOT / "rt_monthly_report.py", *sorted((ROOT / "utils").glob("*.py"))]:
            fonte = arquivo.read_text(encoding="utf-8")
            with self.subTest(arquivo=arquivo.name):
                self.assertNotRegex(fonte, r'/\s*"aqua_gestao_logo[^"]*"')
                self.assertNotRegex(fonte, r'Path\("[^"]*aqua_gestao_logo')

    def test_sem_logos_antigas_embutidas_em_base64(self):
        import base64
        for literal in re.findall(r'"""((?:/9j/|iVBORw0KGgo)[A-Za-z0-9+/=\s]{2000,})"""|"((?:/9j/|iVBORw0KGgo)[A-Za-z0-9+/=]{2000,})"', self.SRC):
            dados = base64.b64decode(re.sub(r"\s", "", literal[0] or literal[1]))
            self.assertNotIn(sha256(dados), LOGOS_AQUA_ANTIGOS_SHA256)
        for nome in ("LOGO_AQUA_OFICIAL_B64 =", "LOGO_AQUA_APROVADA_B64 =", "LOGO_DOSSIE_JPG_B64 =", "_LOGO_B64   ="):
            self.assertNotIn(nome, self.SRC)

    def test_modelos_docx_acqua_trocam_a_arte_antiga_e_preservam_bem_star(self):
        from docx import Document
        for modelo, esperado in [("relatorio_mensal.docx", 1), ("aditivo.docx", 1), ("template_rt_aqua_v2_relatorio.docx", 1),
                                 ("template_bem_star.docx", 0), ("template.docx", 0)]:
            with self.subTest(modelo=modelo):
                doc = Document(ROOT / modelo)
                self.assertGreaterEqual(substituir_logo_antigo_docx(doc, logo_acqua()), esperado)
                if not esperado:
                    self.assertEqual(substituir_logo_antigo_docx(Document(ROOT / modelo), logo_acqua()), 0)


if __name__ == "__main__":
    unittest.main()
