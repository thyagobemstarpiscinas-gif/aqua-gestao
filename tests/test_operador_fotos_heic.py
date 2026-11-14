"""HEIC no upload/rascunho, preview resiliente e fotos no PDF operacional.

Imagens fictícias geradas em memória; sem planilha, Drive ou dados reais.
O AppTest executa o bloco de preview extraído do app, sem carregar o app inteiro.
"""
import ast
import base64
import io
import tempfile
import textwrap
import unittest
from pathlib import Path

from PIL import Image
from pillow_heif import from_pillow
from pypdf import PdfReader
from streamlit.testing.v1 import AppTest

from test_fotos_visita_categorias import ambiente

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)


def _source(nome):
    return ast.get_source_segment(SOURCE, next(
        n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == nome))


# Usa o registro do decoder feito pelo app, em vez de registrá-lo só no teste.
REGISTRO = ast.get_source_segment(SOURCE, next(
    n for n in TREE.body if isinstance(n, ast.Try)
    and any(isinstance(i, ast.ImportFrom) and i.module == "pillow_heif" for i in n.body)))
PREVIEW = textwrap.dedent(ast.get_source_segment(SOURCE, next(
    n for n in ast.walk(TREE) if isinstance(n, ast.For)
    and isinstance(n.target, ast.Tuple)
    and any(isinstance(t, ast.Name) and t.id == "_cat_label" for t in n.target.elts)),
    padded=True))


def _foto(formato="HEIF", cor=(20, 150, 60), tamanho=(64, 48), orientacao=None):
    imagem = Image.new("RGB", tamanho, cor)
    buf = io.BytesIO()
    if formato == "HEIF":
        from_pillow(imagem).save(buf, quality=90)
    else:
        exif = Image.Exif()
        if orientacao:
            exif[274] = orientacao
        imagem.save(buf, format=formato, exif=exif)
    return buf.getvalue()


class Upload(io.BytesIO):
    def __init__(self, nome, dados):
        super().__init__(dados)
        self.name = nome


class FotosHEICOperador(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.heic = _foto()
        cls.jpeg = _foto("JPEG", cor=(30, 70, 180), tamanho=(80, 60))
        ns = {"Path": Path}
        exec(REGISTRO + "\n" + _source("carregar_imagem_corrigida_orientacao"), ns)
        cls.carregar = staticmethod(ns["carregar_imagem_corrigida_orientacao"])

    def test_heic_upload_bytes_e_arquivo_salvo_sem_alterar_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "rasc_antes_ficticia.heic"
            path.write_bytes(self.heic)
            upload = Upload("ficticia.heic", self.heic)
            upload.seek(19)
            for origem in (upload, self.heic, path, str(path)):
                with self.subTest(tipo=type(origem).__name__):
                    imagem = self.carregar(origem)
                    self.assertIsInstance(imagem, Image.Image)
                    self.assertEqual(imagem.size, (64, 48))
                    imagem.save(io.BytesIO(), format="PNG")
            self.assertEqual(upload.tell(), 19)
            self.assertEqual(upload.getvalue(), self.heic)
            self.assertEqual(path.read_bytes(), self.heic)

    def test_formatos_existentes_e_orientacao_exif(self):
        for formato in ("JPEG", "PNG", "WEBP"):
            with self.subTest(formato=formato):
                data = _foto(formato, orientacao=6)
                imagem = self.carregar(data)
                self.assertIsInstance(imagem, Image.Image)
                self.assertEqual(imagem.size, (48, 64))
                self.assertNotEqual(imagem.getexif().get(274), 6)

    def test_ilegivel_nao_retorna_bytes_para_o_streamlit(self):
        for data in (b"arquivo ilegivel", self.heic[:50]):
            upload = Upload("ficticia.heic", data)
            self.assertIsNone(self.carregar(upload))
            self.assertEqual(upload.getvalue(), data)

    def test_preview_continua_apos_ilegivel_e_preserva_rascunho(self):
        codigo = "\n".join([
            "import streamlit as st\nfrom pathlib import Path\nfrom utils.formatacao import limpar_nome_arquivo",
            REGISTRO, _source("carregar_imagem_corrigida_orientacao"), _source("deduplicar_fotos"),
            "_todas_fotos_preview = st.session_state['fotos_teste']", PREVIEW,
            "st.text_input('Observações', key='observacoes')\nst.button('Continuar visita')",
        ])
        with tempfile.TemporaryDirectory() as tmp:
            pasta = Path(tmp)
            (pasta / "_rascunho_operador.json").write_text('{"ph":"7,4"}')
            fotos = [("antes", b"arquivo ilegivel", ".heic"),
                     ("depois", self.heic, ".heic"), ("cmaq", self.jpeg, ".jpg")]
            paths = []
            for cat, dados, ext in fotos:
                path = pasta / f"rasc_{cat}_ficticia{ext}"
                path.write_bytes(dados)
                paths.append(path)
            originais = {p.name: p.read_bytes() for p in pasta.iterdir()}
            for como_upload in (True, False):
                with self.subTest(upload=como_upload):
                    at = AppTest.from_string(codigo, default_timeout=30)
                    at.session_state["fotos_teste"] = [
                        (cat, cat, [Upload("ficticia" + ext, dados)] if como_upload else [], [p])
                        for (cat, dados, ext), p in zip(fotos, paths)]
                    at.session_state["observacoes"] = "Manter esta observação"
                    at.run()
                    self.assertFalse(at.exception, str(at.exception))
                    self.assertEqual(len(at.warning), 1)
                    self.assertIn("mantido no rascunho", at.warning[0].value)
                    self.assertEqual(len(at.image), 2)
                    self.assertEqual(at.text_input[0].value, "Manter esta observação")
                    at.button[0].click().run()
                    self.assertFalse(at.exception, str(at.exception))
                    self.assertEqual({p.name: p.read_bytes() for p in pasta.iterdir()}, originais)

    def test_heic_recuperado_por_id_entra_no_pdf_na_categoria_certa(self):
        with tempfile.TemporaryDirectory() as tmp:
            ns = ambiente(Path(tmp))
            ns["drive_baixar_foto"] = {"antes": self.heic, "depois": self.jpeg}.get
            visita = {"condominio": "Condomínio Fictício", "data": "14/11/2026",
                      "fotos_antes_ids": ["antes"], "fotos_depois_ids": ["depois"]}
            categorias = ns["_resolver_fotos_visita_para_relatorio"](visita)
            self.assertEqual([len(categorias[c]) for c in ("antes", "depois", "cmaq")], [1, 1, 0])
            with Image.open(io.BytesIO(base64.b64decode(categorias["antes"][0]))) as im:
                self.assertEqual(im.format, "JPEG")
                self.assertEqual(im.size, (64, 48))
            pdf = ns["gerar_pdf_relatorio_visita"](visita, visita["condominio"])
            reader = PdfReader(io.BytesIO(pdf))
            texto = "\n".join(p.extract_text() or "" for p in reader.pages)
            self.assertIn("Antes do tratamento:", texto)
            self.assertIn("Após o tratamento:", texto)
            tamanhos = [im.image.size for p in reader.pages for im in p.images]
            self.assertIn((64, 48), tamanhos)
            self.assertIn((80, 60), tamanhos)


if __name__ == "__main__":
    unittest.main()
