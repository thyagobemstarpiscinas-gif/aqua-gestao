"""Destinos de homologação: configuração pura, sem acesso ao Google."""
import ast
from pathlib import Path
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.ambiente import destinos_google

class TestAmbiente(unittest.TestCase):
    producao = ("planilha_producao_123", "pasta_producao_123")

    def selecionar(self, cfg):
        return destinos_google({"aqua_app": cfg}, *self.producao)

    def test_sem_configuracao_preserva_destinos_existentes(self):
        self.assertEqual(destinos_google({}, *self.producao), ("producao", *self.producao))

    def test_sem_arquivo_secrets_mantem_comportamento_local(self):
        class Ausente:
            def get(self, *_): raise FileNotFoundError("Secrets fictício ausente")
        self.assertEqual(destinos_google(Ausente(), *self.producao), ("producao", *self.producao))

    def test_homologacao_exige_os_dois_destinos(self):
        for cfg in ({"ambiente":"homologacao"},
                    {"ambiente":"homologacao", "sheets_id":"planilha_teste_123"},
                    {"ambiente":"homologacao", "drive_fotos_folder_id":"pasta_teste_123"}):
            with self.subTest(cfg=cfg), self.assertRaises(ValueError): self.selecionar(cfg)

    def test_homologacao_recusa_producao_em_qualquer_destino(self):
        for sheet, drive in ((self.producao[0], "pasta_teste_123"),
                             ("planilha_teste_123", self.producao[1])):
            with self.subTest(sheet=sheet, drive=drive), self.assertRaises(ValueError):
                self.selecionar({"ambiente":"homologacao", "sheets_id":sheet, "drive_fotos_folder_id":drive})

    def test_destinos_alternativos_nao_sao_ativados_sem_modo_explicito(self):
        with self.assertRaises(ValueError):
            self.selecionar({"sheets_id":"planilha_teste_123"})

    def test_configuracao_invalida_nao_recai_em_producao(self):
        for cfg in (None, [], {"ambiente":"homologação"},
                    {"ambiente":"homologacao", "sheets_id":123, "drive_fotos_folder_id":"pasta_teste_123"},
                    {"ambiente":"homologacao", "sheets_id":"https://exemplo.invalid", "drive_fotos_folder_id":"pasta_teste_123"}):
            with self.subTest(cfg=cfg), self.assertRaises(ValueError): self.selecionar(cfg)

    def test_modo_homologacao_retorna_so_destinos_de_teste(self):
        self.assertEqual(self.selecionar({"ambiente":"homologacao", "sheets_id":"planilha_teste_123",
            "drive_fotos_folder_id":"pasta_teste_123"}), ("homologacao", "planilha_teste_123", "pasta_teste_123"))

    def test_bloco_real_interrompe_configuracao_insegura(self):
        tree = ast.parse((Path(__file__).resolve().parents[1]/"app.py").read_text())
        node = next(n for n in tree.body if isinstance(n, ast.Try)
                    and any(isinstance(x, ast.Assign) and isinstance(x.value, ast.Call)
                            and getattr(x.value.func, 'id', '') == 'destinos_google' for x in n.body))
        class Parada(Exception): pass
        st = SimpleNamespace(secrets={"aqua_app":{"ambiente":"homologacao"}},
                             error=Mock(), stop=Mock(side_effect=Parada))
        ns = {"destinos_google":destinos_google, "st":st, "SHEETS_ID":self.producao[0],
              "DRIVE_FOTOS_FOLDER_ID":self.producao[1]}
        with self.assertRaises(Parada):
            exec(compile(ast.Module(body=[node], type_ignores=[]), "app.py", "exec"), ns)
        st.error.assert_called_once();st.stop.assert_called_once()

if __name__ == "__main__": unittest.main(verbosity=2)
