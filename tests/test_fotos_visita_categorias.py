"""Classificação das fotos no relatório de visita operacional (não no relatório mensal RT).

Fotos Depois e Casa de máquinas, recuperadas do Drive pelos IDs por categoria, devem sair
nas seções "Após o tratamento" e "Casa de máquinas". Visitas antigas, que só têm a lista
geral fotos_drive_ids, continuam com essas fotos em "Antes do tratamento".
"""
import __future__
import ast
import io
import re
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from PIL import Image
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]


def _jpeg(cor, tamanho):
    # Tamanhos diferentes, como fotos reais: a deduplicação do relatório compara o início
    # e o comprimento do base64, e imagens lisas do mesmo tamanho seriam tidas como iguais.
    buf = io.BytesIO()
    Image.new("RGB", tamanho, cor).save(buf, format="JPEG")
    return buf.getvalue()


FOTOS_DRIVE = {"id_depois": _jpeg((30, 140, 60), (48, 32)), "id_cmaq": _jpeg((30, 80, 170), (96, 64)),
               "id_antes": _jpeg((200, 60, 40), (160, 120))}


def ambiente(generated: Path):
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    ns = {"BASE_DIR": ROOT, "GENERATED_DIR": generated, "datetime": datetime, "date": date,
          "DANCING_SCRIPT_FONT_B64": "", "LOGO_AQUA_OFICIAL_B64": "",
          "CRQ": "CRQ-MG 2ª Região | CRQ 024025748",
          "slugify_nome": lambda s: re.sub(r"\W+", "_", str(s)).strip("_"),
          "drive_baixar_foto": FOTOS_DRIVE.get}
    constants = {"LOGO_CANDIDATOS", "NBR_LIM", "NBR_TXT", "NBR_CC_MAX"}
    names = {"_resolver_fotos_visita_para_relatorio", "_normalizar_lista_fotos_b64",
             "_imagem_bytes_para_b64_relatorio", "encontrar_logo", "gerar_pdf_relatorio_visita",
             "valor_float", "_normalizar_assinatura_b64", "normalizar_texto"}
    for node in tree.body:
        if (isinstance(node, ast.ImportFrom) and node.module == "utils.normas_piscinas" or
                isinstance(node, ast.FunctionDef) and node.name in names or
                isinstance(node, ast.Assign) and any(isinstance(x, ast.Name) and x.id in constants
                                                      for x in node.targets)):
            exec(compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / "app.py"), "exec",
                         flags=__future__.annotations.compiler_flag), ns)
    return ns


def _cor_media(b64):
    import base64
    with Image.open(io.BytesIO(base64.b64decode(b64))) as im:
        return im.convert("RGB").resize((1, 1)).getpixel((0, 0))


def _perto(cor, alvo, tol=25):
    return all(abs(a - b) <= tol for a, b in zip(cor, alvo))


class ClassificacaoFotosVisita(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ns = ambiente(Path(self.tmp.name))  # disco vazio: nada local para achar

    def tearDown(self):
        self.tmp.cleanup()

    def visita_atual(self):
        # Formato gravado pelo modo campo: IDs por categoria e a lista geral com todos.
        return {"condominio": "Condomínio Alfa Fictício", "data": "12/11/2026", "operador": "João",
                "fotos_antes_ids": [], "fotos_depois_ids": ["id_depois"], "fotos_cmaq_ids": ["id_cmaq"],
                "fotos_extras_ids": [], "fotos_drive_ids": ["id_depois", "id_cmaq"],
                "piscinas": [{"nome": "Adulto", "ph": "7.4", "cloro_livre": "2.0", "cloro_total": "2.1",
                              "alcalinidade": "100", "dureza": "200"}]}

    def test_depois_e_casa_de_maquinas_nas_proprias_categorias(self):
        r = self.ns["_resolver_fotos_visita_para_relatorio"](self.visita_atual())
        self.assertEqual(r["antes"], [])
        self.assertEqual(len(r["depois"]), 1)
        self.assertEqual(len(r["cmaq"]), 1)
        self.assertTrue(_perto(_cor_media(r["depois"][0]), (30, 140, 60)))
        self.assertTrue(_perto(_cor_media(r["cmaq"][0]), (30, 80, 170)))

    def test_visita_antiga_so_com_lista_geral_continua_em_antes(self):
        antiga = {"condominio": "Condomínio Alfa Fictício", "fotos_drive_ids": ["id_antes"]}
        r = self.ns["_resolver_fotos_visita_para_relatorio"](antiga)
        self.assertEqual(len(r["antes"]), 1)
        self.assertEqual((r["depois"], r["cmaq"]), ([], []))

    def test_antes_por_categoria_nao_e_duplicado_pela_lista_geral(self):
        v = self.visita_atual()
        v["fotos_antes_ids"] = ["id_antes"]
        v["fotos_drive_ids"] = ["id_antes", "id_depois", "id_cmaq"]
        r = self.ns["_resolver_fotos_visita_para_relatorio"](v)
        self.assertEqual([len(r[c]) for c in ("antes", "depois", "cmaq")], [1, 1, 1])
        self.assertTrue(_perto(_cor_media(r["antes"][0]), (200, 60, 40)))

    def test_pdf_da_visita_operacional_mostra_as_secoes_corretas(self):
        pdf = self.ns["gerar_pdf_relatorio_visita"](self.visita_atual(), "Condomínio Alfa Fictício")
        texto = "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(pdf)).pages)
        self.assertIn("Registro fotogr", texto)
        self.assertIn("Após o tratamento:", texto)
        self.assertIn("Casa de máquinas:", texto)
        self.assertNotIn("Antes do tratamento:", texto)
        self.assertLess(texto.index("Após o tratamento:"), texto.index("Casa de máquinas:"))


if __name__ == "__main__":
    unittest.main()
