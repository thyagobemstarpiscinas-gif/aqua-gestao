"""Cadastro real com Sheets simulado; nenhuma credencial ou rede."""
import ast
import copy
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
import unittest

from utils.cadastro_piscinas import volume, ler_metadados, serializar


class CadastroPiscinas(unittest.TestCase):
    def setUp(self):
        self.row = ["", "C001", "Alfa", "100", "", "", "", "2026-01-01", "Ativo", "100", "", "", "Aqua Gestão", "", ""]
        self.rows = [self.row.copy(), ["", "C002", "Beta"]]
        self.updates = []
        self.inserts = []
        self.payload = serializar({"adulto": "Adulto", "extra1": "SPA"}, [{"nome": "SPA", "vol": 12.5}])
        aba = SimpleNamespace(get_all_values=lambda: copy.deepcopy(self.rows),
                              update=lambda **kw: self.updates.append(kw),
                              insert_row=lambda *a, **kw: self.inserts.append((a, kw)))
        self.ns = dict(conectar_sheets=lambda: object(), obter_aba_sheets=lambda _: aba,
                       formatar_telefone=lambda x: x, datetime=datetime,
                       ler_piscinas_cadastro=ler_metadados, volume_cadastro=volume,
                       _log_sheets_erro=lambda *a: None,
                       _enriquecer_cliente_com_dados_locais=lambda x: x,
                       st=SimpleNamespace(cache_data=SimpleNamespace(clear=lambda: None)))
        source = ast.parse((Path(__file__).resolve().parents[1] / "app.py").read_text())
        names = {"sheets_salvar_cliente", "sheets_editar_cliente", "sheets_listar_clientes_completo", "_enriquecer_cliente_com_dados_locais"}
        nodes = [n for n in source.body if isinstance(n, ast.FunctionDef) and n.name in names]
        for n in nodes:
            n.decorator_list = []
        exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), self.ns)
        self.ns.update(_carregar_dados_cliente_local=lambda _: {},
                       _normalizar_servicos_cliente=lambda x: {},
                       _servicos_cliente_resolvidos=lambda cliente, locais: {},
                       _servicos_para_empresa=lambda _: "Aqua Gestão",
                       _normalizar_lista_textos_unicos=lambda x: x)

    def editar(self, **kwargs):
        return self.ns["sheets_editar_cliente"]("C001", "Alfa", "", "", "", "", vol_adulto=100, **kwargs)

    def test_edicao_isolada_preserva_data_status_e_colunas_legadas(self):
        self.rows[0] += [self.payload, "CONTEÚDO Q"]
        original = copy.deepcopy(self.rows)
        self.assertTrue(self.editar(piscinas_cadastro=self.payload))
        update = self.updates[0]
        self.assertEqual(update["range_name"], "A1:P1")
        row = update["values"][0]
        self.assertEqual(row[3], "112.5")
        self.assertEqual(row[7:15], self.row[7:15])
        self.assertEqual(ler_metadados(row[15])["extras"][0]["nome"], "SPA")
        self.assertEqual(self.rows, original)

    def test_chamador_legado_nao_escreve_coluna_p(self):
        self.rows[0].append(self.payload)
        self.assertTrue(self.editar())
        self.assertEqual(self.updates[0]["range_name"], "A1:O1")

    def test_coluna_ocupada_bloqueia_escrita(self):
        for texto in ("Outro uso", '{"schema":"outra.v1"}', '{'):
            self.rows[0] = self.row + [texto]
            self.assertFalse(self.editar(piscinas_cadastro=self.payload))
        self.assertEqual(self.updates, [])

    def test_cliente_inexistente_nao_altera_outro(self):
        self.rows = [self.rows[1]]
        self.assertFalse(self.editar(piscinas_cadastro=self.payload))
        self.assertEqual(self.updates, [])

    def test_recupera_extras_e_tipos_sem_disco_local(self):
        self.rows[0].append(self.payload)
        clientes = self.ns["sheets_listar_clientes_completo"]()
        self.assertEqual(clientes[0]["tipos_piscinas"]["extra1"], "SPA")
        self.assertEqual(clientes[0]["piscinas_extras"][0]["vol"], 12.5)
        self.assertEqual(clientes[1]["piscinas_extras"], [])

    def test_legado_mantem_volumes_e_extras_locais(self):
        self.ns["_carregar_dados_cliente_local"] = lambda _: {"piscinas_extras": [{"nome": "Coberta", "vol": 20}]}
        clientes = self.ns["sheets_listar_clientes_completo"]()
        self.assertEqual(clientes[0]["vol_adulto"], 100)
        self.assertEqual(clientes[0]["piscinas_extras"][0]["nome"], "Coberta")

    def test_remoto_prevalece_sobre_copia_local_antiga(self):
        self.rows[0].append(self.payload)
        self.ns["_carregar_dados_cliente_local"] = lambda _: {"vol_adulto": 999, "piscinas_extras": []}
        cliente = self.ns["sheets_listar_clientes_completo"]()[0]
        self.assertEqual(cliente["vol_adulto"], 100)
        self.assertEqual(len(cliente["piscinas_extras"]), 1)

    def test_coluna_invalida_nao_oculta_demais_clientes(self):
        self.rows[0].append("Outro uso")
        clientes = self.ns["sheets_listar_clientes_completo"]()
        self.assertEqual(len(clientes), 2)
        self.assertIn("erro_cadastro_piscinas", clientes[0])

    def test_novo_cliente_grava_extras_junto_com_volumes(self):
        self.assertTrue(self.ns["sheets_salvar_cliente"]("Gama", "", "", "", "", vol_adulto=100, piscinas_cadastro=self.payload))
        row = self.inserts[0][0][0]
        self.assertEqual(row[1], "C003")
        self.assertEqual(row[3], "112.5")
        self.assertEqual(row[15], self.payload)

    def test_duplicado_nao_simula_sucesso(self):
        self.assertFalse(self.ns["sheets_salvar_cliente"]("Alfa", "", "", "", "", piscinas_cadastro=self.payload))
        self.assertEqual(self.inserts, [])

    def test_volumes_invalidos_recusados(self):
        for valor in ("abc", "nan", "inf", "-1"):
            with self.subTest(valor=valor), self.assertRaises(ValueError):
                volume(valor)
        self.assertEqual(volume("12,5"), 12.5)
        self.assertEqual(volume(""), 0)

    def test_payload_invalido_nao_grava(self):
        self.assertFalse(self.editar(piscinas_cadastro='{}'))
        self.assertEqual(self.updates, [])

    def test_tipo_sem_volumes_nao_apaga_total_legado(self):
        self.rows[0][9] = ""
        payload = serializar({"adulto": "Adulto"}, [])
        self.assertTrue(self.ns["sheets_editar_cliente"]("C001", "Alfa", "", "", "", "", piscinas_cadastro=payload))
        self.assertEqual(self.updates[0]["values"][0][3], "100.0")


if __name__ == "__main__":
    unittest.main()
