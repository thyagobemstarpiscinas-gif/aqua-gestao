"""F1/F2/F3: funções reais extraídas por AST, dados fictícios, sem app ou rede."""
import ast
import copy
import json
import os
from pathlib import Path
import socket
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from utils.rt_contexto import caminho_ultimos, registrar_ultimos, carregar_ultimos

SRC = (RAIZ / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SRC)


class Estado(dict):
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__


def carregar(nomes, **extras):
    ns = {"st": SimpleNamespace(session_state=Estado()), "Path": Path,
          "conectar_sheets": lambda: None, "ANALISES_PADRAO": 12, "_RASCUNHO_RT_NAO_RESTAURAR": set()}
    ns.update(extras)
    nodes = [n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name in nomes]
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
    return ns


class TestRegressoes(unittest.TestCase):
    def setUp(self):
        ambiente = patch.dict(os.environ, {"CODESPACES": "true"})
        ambiente.start()
        self.addCleanup(ambiente.stop)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.raiz = Path(self.tmp.name)
        self.chave = "aqua_gestao|id:C1|09/2026"
        self.snapshot = {"parametros_ultimos": [{"data": "01/09/2026", "ph": "7.3", "temperatura": "28", "orp": "700"}],
                         "dosagens_ultimas": [{"produto": "Cloro", "quantidade": "0"}]}
        bloqueio = patch.object(socket.socket, "connect", side_effect=AssertionError("Rede proibida"))
        bloqueio.start()
        self.addCleanup(bloqueio.stop)

    def recuperar(self):
        ns = carregar(["recuperar_ultimos_valores_rt", "aplicar_dosagens_ultimas_no_relatorio"],
                      GENERATED_DIR=self.raiz, validar_contexto_rt=lambda: [])
        ns["st"].session_state.update(rel_origem_dados="aqua_gestao|id:C1|10/2026",
                                      rel_analises_total=12, rel_analise_data_0="15/10/2026",
                                      rel_analise_ph_0="ATUAL", rel_dos_produto_0="ATUAL")
        return ns

    def test_ambos_recuperam_entre_periodos_sem_datas_e_mostram_origem(self):
        registrar_ultimos(self.raiz, self.chave, self.snapshot)
        ns = self.recuperar()
        ss = ns["st"].session_state
        ns["recuperar_ultimos_valores_rt"]("parametros")
        self.assertEqual(ss["rel_analise_ph_0"], "7.3")
        self.assertEqual(ss["rel_analise_temperatura_0"], "28")
        self.assertEqual(ss["rel_analise_orp_0"], "700")
        self.assertEqual(ss["rel_analise_data_0"], "15/10/2026")
        self.assertIn("09/2026", ss["_rt_ultimos_aviso"])
        ns["recuperar_ultimos_valores_rt"]("dosagens")
        self.assertEqual(ss["rel_dos_produto_0"], "Cloro")
        self.assertEqual(ss["rel_dos_qtd_0"], "0")
        self.assertIn("09/2026", ss["_rt_ultimos_aviso"])
        self.assertNotIn("data", carregar_ultimos(self.raiz, self.chave)["parametros_ultimos"][0])

    def test_ambos_preservam_em_ausencia_corrupcao_identidade_e_legado(self):
        for caso in ("ausente", "json", "identidade", "estrutura", "legado", "sem_id", "contexto", "outra_empresa", "outro_id"):
            for tipo in ("parametros", "dosagens"):
                with self.subTest(caso=caso, tipo=tipo):
                    registrar_ultimos(self.raiz, self.chave, self.snapshot)
                    p = caminho_ultimos(self.raiz, self.chave)
                    if caso in ("ausente", "legado"):
                        p.unlink()
                    elif caso == "json":
                        p.write_text("{", encoding="utf-8")
                    elif caso in ("identidade", "estrutura"):
                        d = json.loads(p.read_text(encoding="utf-8"))
                        d["id_cliente" if caso == "identidade" else "dosagens_ultimas"] = "OUTRO"
                        p.write_text(json.dumps(d), encoding="utf-8")
                    legado = self.raiz / "Residencial_Sol"
                    legado.mkdir(exist_ok=True)
                    (legado / "dados_condominio.json").write_text(json.dumps(self.snapshot), encoding="utf-8")
                    ns = self.recuperar()
                    if caso == "sem_id":
                        ns["st"].session_state["rel_origem_dados"] = "aqua_gestao|nome:sol|10/2026"
                    if caso == "contexto":
                        ns["validar_contexto_rt"] = lambda: ["Identidade não validada"]
                    if caso == "outra_empresa":
                        ns["st"].session_state["rel_origem_dados"] = "bem_star|id:C1|10/2026"
                    if caso == "outro_id":
                        ns["st"].session_state["rel_origem_dados"] = "aqua_gestao|id:C2|10/2026"
                    antes = dict(ns["st"].session_state)
                    ns["recuperar_ultimos_valores_rt"](tipo)
                    depois = dict(ns["st"].session_state)
                    self.assertIn("Não foi possível", depois.pop("_rt_ultimos_aviso"))
                    self.assertEqual(antes, depois)

    def test_chave_separa_empresas_ids_e_independe_nome_periodo(self):
        registrar_ultimos(self.raiz, self.chave, self.snapshot)
        self.assertEqual(caminho_ultimos(self.raiz, self.chave), caminho_ultimos(self.raiz, "aqua_gestao|id:C1|10/2026"))
        for chave in ("bem_star|id:C1|09/2026", "aqua_gestao|id:C2|09/2026"):
            self.assertNotEqual(caminho_ultimos(self.raiz, self.chave), caminho_ultimos(self.raiz, chave))
            with self.assertRaises(FileNotFoundError):
                carregar_ultimos(self.raiz, chave)

    def test_substituicao_atomica_falha_preserva_e_ordem_de_geracao(self):
        registrar_ultimos(self.raiz, self.chave, self.snapshot)
        p = caminho_ultimos(self.raiz, self.chave)
        antes = p.read_bytes()
        with patch("utils.rt_contexto.os.replace", side_effect=OSError("fictício")):
            with self.assertRaises(OSError):
                registrar_ultimos(self.raiz, "aqua_gestao|id:C1|08/2026", self.snapshot)
        self.assertEqual(p.read_bytes(), antes)
        with self.assertRaises(ValueError):
            registrar_ultimos(self.raiz, self.chave, {"parametros_ultimos": [{}], "dosagens_ultimas": []})
        self.assertEqual(p.read_bytes(), antes)
        registrar_ultimos(self.raiz, "aqua_gestao|id:C1|08/2026", self.snapshot)
        self.assertEqual(carregar_ultimos(self.raiz, self.chave)["periodo"], "08/2026")

    def test_rascunho_invalido_preserva_pendente_e_todo_estado(self):
        for dados in ({"rel_analises_total": "abc", "campos": {"rel_diagnostico": "novo"}},
                      {"rel_analises_total": 12, "campos": {"rel_analises_total": []}},
                      {"rel_analises_total": {}, "campos": {"rel_diagnostico": "novo"}},
                      {"campos": {"rel_fotos_preservadas": "{"}},
                      {"empresa_ativa": "bem_star", "campos": {"rel_diagnostico": "novo"}}):
            with self.subTest(dados=dados):
                guardar = Mock(side_effect=AssertionError("Não guardar"))
                ns = carregar(["_relatorio_rt_aplicar_rascunho", "_relatorio_rt_aplicar_pendente"], guardar_tabela_rt=guardar)
                ss = ns["st"].session_state
                ss.update(empresa_ativa="aqua_gestao", rel_origem_dados=self.chave,
                          rel_diagnostico="original", rel_fotos_preservadas="FOTO_ORIGINAL",
                          _rt_rasc_pendente={"dados": dados}, _rt_rasc_chave="sessao",
                          _rt_rasc_ref={"x": 1}, _rt_rasc_lista_cache=[1], _rt_fotos_nonce=3)
                antes = copy.deepcopy(ss)
                ns["_relatorio_rt_aplicar_pendente"]()
                self.assertIn("_relatorio_rt_restore_erro", ss)
                ss.pop("_relatorio_rt_restore_erro")
                self.assertEqual(ss, antes)
                guardar.assert_not_called()

    def test_contador_normalizado_nao_e_sobrescrito(self):
        ns = carregar(["_relatorio_rt_aplicar_rascunho"])
        self.assertTrue(ns["_relatorio_rt_aplicar_rascunho"]({"rel_analises_total": "14", "campos": {"rel_analises_total": "2", "rel_diagnostico": "ok"}}))
        self.assertEqual(ns["st"].session_state["rel_analises_total"], 14)

    def test_callback_rejeitado_preserva_pendente_anterior(self):
        for anterior in (None, {"dados": {"campos": {"rel_diagnostico": "pendente antigo"}}}):
            ns = carregar(["_relatorio_rt_aplicar_rascunho", "_relatorio_rt_aplicar_pendente", "_relatorio_rt_cb_restaurar"])
            ss = ns["st"].session_state
            ss["rel_diagnostico"] = "atual"
            if anterior is not None:
                ss["_rt_rasc_pendente"] = anterior
            antes = copy.deepcopy(ss)
            ns["_relatorio_rt_cb_restaurar"]({"rel_analises_total": "erro", "campos": {"rel_diagnostico": "novo"}})
            ss.pop("_relatorio_rt_restore_erro")
            self.assertEqual(ss, antes)

    def test_falha_ao_guardar_tambem_reverte_controles(self):
        ns = carregar(["_relatorio_rt_aplicar_rascunho"])
        ss = ns["st"].session_state
        ss.update(rel_origem_dados=self.chave, rel_diagnostico="atual", _rt_fotos_nonce=3)
        antes = dict(ss)
        def falhar(_):
            ss.update(_rt_contexto_erro="fictício", _rt_fotos_nonce=4, rel_diagnostico="alterado")
            return False
        ns["guardar_tabela_rt"] = falhar
        self.assertFalse(ns["_relatorio_rt_aplicar_rascunho"]({"campos": {"rel_diagnostico": "novo"}}))
        ss.pop("_relatorio_rt_restore_erro")
        self.assertEqual(ss, antes)

    def test_parametros_adicionais_tem_datas_vazias(self):
        self.snapshot["parametros_ultimos"] *= 3
        registrar_ultimos(self.raiz, self.chave, self.snapshot)
        ns = self.recuperar()
        ss = ns["st"].session_state
        ss["rel_analises_total"] = 1
        ns["recuperar_ultimos_valores_rt"]("parametros")
        self.assertEqual(ss["rel_analises_total"], 3)
        self.assertEqual(ss["rel_analise_data_0"], "15/10/2026")
        self.assertEqual(ss["rel_analise_data_2"], "")
        # Executa o bloco real de proteção das linhas durante o rerun.
        node = next(n for n in TREE.body if isinstance(n, ast.If) and isinstance(n.test, ast.BoolOp)
                    and isinstance(n.test.values[0], ast.Name) and n.test.values[0].id == "_linhas_recuperadas")
        estado = {"st": ns["st"], "_linhas_recuperadas": ss["_rt_ultimos_linhas"], "_linhas_freq": 1}
        exec(compile(ast.Module(body=[node], type_ignores=[]), "rerun", "exec"), estado)
        self.assertEqual(estado["_linhas_freq"], 3)
        ss["rel_origem_dados"] = "aqua_gestao|id:C2|10/2026"
        estado["_linhas_freq"] = 1
        exec(compile(ast.Module(body=[node], type_ignores=[]), "rerun", "exec"), estado)
        self.assertEqual(estado["_linhas_freq"], 1)

    def test_nomes_colidentes_e_renomeacao_em_ambos_botoes(self):
        registrar_ultimos(self.raiz, self.chave, self.snapshot)
        for nome in ("Residencial Sol", "Residencial Sol!", "A" * 121, "Nome atualizado"):
            for tipo in ("parametros", "dosagens"):
                ns = self.recuperar()
                ss = ns["st"].session_state
                ss["rel_nome_condominio"] = nome
                ns["recuperar_ultimos_valores_rt"](tipo)
                self.assertIn("09/2026", ss["_rt_ultimos_aviso"])
                ss["rel_origem_dados"] = "aqua_gestao|id:C2|10/2026"
                antes = dict(ss)
                ns["recuperar_ultimos_valores_rt"](tipo)
                antes.pop("_rt_ultimos_aviso")
                depois = dict(ss)
                depois.pop("_rt_ultimos_aviso")
                self.assertEqual(antes, depois)

    def test_botoes_chamam_recuperacao_validada(self):
        for label, tipo in (("Carregar parâmetros usados pela última vez", "parametros"), ("Carregar usados pela última vez", "dosagens")):
            node = next(n for n in ast.walk(TREE) if isinstance(n, ast.If) and isinstance(n.test, ast.Call) and n.test.args and isinstance(n.test.args[0], ast.Constant) and n.test.args[0].value == label)
            recuperar = Mock()
            exec(compile(ast.Module(body=node.body, type_ignores=[]), "botao", "exec"), {"recuperar_ultimos_valores_rt": recuperar, "st": SimpleNamespace(rerun=lambda: None)})
            recuperar.assert_called_once_with(tipo)

    def test_renderizador_publica_so_final_com_documentos_concluidos(self):
        for caso in ("sucesso", "previa", "pdf_falha", "docx_falha", "manifest_falha", "fallback"):
            with self.subTest(caso=caso):
                registrar_ultimos(self.raiz, self.chave, self.snapshot)
                p = caminho_ultimos(self.raiz, self.chave)
                antes = p.read_bytes()
                def docx(_template, destino, *_a, **_k):
                    if caso == "docx_falha":
                        raise OSError("fictício")
                    destino.write_bytes(b"docx ficticio")
                def pdf(_dados, _fotos, destino):
                    if caso not in ("pdf_falha", "fallback"):
                        destino.write_bytes(b"pdf ficticio")
                    return caso not in ("pdf_falha", "fallback"), "ficticio"
                def fallback(_origem, destino):
                    if caso == "fallback":
                        destino.write_bytes(b"pdf fallback")
                    return caso == "fallback", "ficticio"
                from datetime import datetime
                ns = carregar(["_renderizar_relatorio_rt"], GENERATED_DIR=self.raiz, datetime=datetime,
                    validar_contexto_rt=lambda: [], montar_dados_relatorio=lambda: {"nome_condominio": "Ficticio"},
                    validar_relatorio_mensal=lambda _: [], obter_snapshot_relatorio_independente=lambda: self.snapshot,
                    salvar_dados_condominio=lambda *_: None, _resolver_fotos_relatorio_rt=lambda *a, **k: ([], "teste"),
                    limpar_nome_arquivo=lambda n: n, preencher_relatorio_mensal_docx=docx, TEMPLATE_RELATORIO=None,
                    gerar_pdf_relatorio_rt_premium_reportlab=pdf, converter_docx_para_pdf=fallback,
                    registrar_documento_manifest=Mock(side_effect=OSError("fictício") if caso == "manifest_falha" else None),
                    obter_status_art_texto=lambda _: "ficticio")
                ns["st"].warning = Mock()
                ns["st"].session_state["rel_origem_dados"] = "aqua_gestao|id:C1|10/2026"
                if caso in ("docx_falha", "manifest_falha"):
                    with self.assertRaises(OSError):
                        ns["_renderizar_relatorio_rt"]()
                else:
                    ns["_renderizar_relatorio_rt"](preview=caso == "previa")
                if caso in ("sucesso", "fallback"):
                    self.assertEqual(carregar_ultimos(self.raiz, self.chave)["periodo"], "10/2026")
                else:
                    self.assertEqual(p.read_bytes(), antes)


if __name__ == "__main__":
    unittest.main(verbosity=2)
