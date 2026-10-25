"""Testes do rascunho do Relatório RT (autosave, proteção e restauração).

Carregam somente as funções de rascunho do app.py, sem executar o app, com
session_state, planilha e disco simulados (pasta temporária). Nenhum acesso a
Google Sheets, Drive ou rede.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import atexit
import copy
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from utils.formatacao import slugify_nome  # noqa: E402

APP = RAIZ / "app.py"
TREE = ast.parse(APP.read_text(encoding="utf-8"))

FUNCOES_RT = [
    "_rascunho_usa_sheets", "_relatorio_rt_tem_conteudo", "_relatorio_rt_tem_dados_em_tela",
    "_relatorio_rt_assinatura", "_relatorio_rt_rev", "_relatorio_rt_coletar_rascunho",
    "_relatorio_rt_chave_conflito", "_relatorio_rt_arquivo_local", "_relatorio_rt_aba",
    "_relatorio_rt_ler_json", "_relatorio_rt_ler_rascunhos", "_relatorio_rt_salvar_rascunho",
    "_relatorio_rt_carregar_rascunho", "_relatorio_rt_aplicar_rascunho", "_relatorio_rt_aplicar_pendente",
]
FUNCOES_OPERADOR = ["salvar_rascunho_operador", "carregar_rascunho_operador", "deletar_rascunho_operador"]
CONSTANTES = {
    "RASCUNHO_RT_USUARIO", "RASCUNHO_RT_LIMITE_SHEETS", "_RASCUNHO_RT_CAMPOS_AUTOMATICOS",
    "_RASCUNHO_RT_PREFIXOS_CONTEUDO", "_RASCUNHO_RT_CAMPOS_CONTEUDO", "_RASCUNHO_RT_NAO_RESTAURAR",
}


def _nos(nomes_funcoes):
    nos = [n for n in TREE.body if isinstance(n, ast.Assign)
           and any(isinstance(t, ast.Name) and t.id in CONSTANTES for t in n.targets)]
    for nome in nomes_funcoes:
        f = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == nome)
        f.decorator_list = []
        nos.append(f)
    return nos


class _Aba:
    """Planilha em memória com a mesma interface usada pelo app."""

    def __init__(self, titulo):
        self.titulo = titulo
        self.linhas = []
        self.gravacoes = 0

    def get_all_values(self):
        return copy.deepcopy(self.linhas)

    def update(self, *args, range_name=None, values=None, **_):
        if args:
            range_name, values = args[0], args[1]
        linha = int(re.match(r"[A-Z]+(\d+)", range_name).group(1))
        while len(self.linhas) < linha:
            self.linhas.append([])
        self.linhas[linha - 1] = list(values[0])
        self.gravacoes += 1


class _Planilha:
    def __init__(self):
        self.abas = {}
        self.conexoes = 0

    def add_worksheet(self, title, rows, cols):
        self.abas[title] = _Aba(title)
        return self.abas[title]


class _St:
    def __init__(self):
        self.session_state = {}


class Ambiente:
    """Armazenamento compartilhado (planilha + disco) entre sessões simuladas."""

    def __init__(self, codespaces=False):
        self.dir = Path(tempfile.mkdtemp(prefix="rascunho_rt_teste_"))
        atexit.register(shutil.rmtree, self.dir, ignore_errors=True)
        self.planilha = _Planilha()
        self.codespaces = codespaces
        self._relogio = 0

    def agora(self):
        self._relogio += 1
        return f"25/10/2026 10:{self._relogio // 60:02d}:{self._relogio % 60:02d}"

    def conectar(self):
        self.planilha.conexoes += 1
        return self.planilha

    def obter_aba(self, nome):
        if nome not in self.planilha.abas:
            raise KeyError(nome)
        return self.planilha.abas[nome]

    def sessao(self):
        st = _St()
        ns = {
            "st": st, "os": os, "re": re, "json": json, "hashlib": hashlib, "datetime": datetime, "Path": Path,
            "ANALISES_PADRAO": 12, "GENERATED_DIR": self.dir, "slugify_nome": slugify_nome,
            "_agora_brasilia": self.agora, "conectar_sheets": self.conectar, "obter_aba_sheets": self.obter_aba,
            "_relatorio_rt_rascunho_path": lambda: self.dir / "_rascunho_relatorio_rt_em_andamento.json",
        }
        exec(compile(ast.Module(body=_nos(FUNCOES_RT + FUNCOES_OPERADOR), type_ignores=[]), str(APP), "exec"), ns)
        ns["abrir_formulario"] = lambda: _abrir_formulario(st.session_state)
        return ns

    def aba(self):
        return self.planilha.abas.get("_Rascunhos_RT")

    def linha(self, chave):
        aba = self.aba()
        for row in (aba.linhas if aba else [])[1:]:
            if row and row[0] == chave:
                return json.loads(row[2])
        return None


def _abrir_formulario(ss):
    """Estado de um formulário recém-aberto: só valores que o app preenche sozinho."""
    ss.update({"rel_status_agua": "CONFORME", "rel_status_agua_select": "CONFORME", "rel_verificacoes_semanais": 1,
               "rel_analises_total": 4, "rel_mes_referencia": "10", "rel_ano_referencia": "2026",
               "rel_tipo_atendimento": "Contrato ativo"})
    for i in range(4):
        for suf in ("data", "ph", "cl", "ct", "operador"):
            ss.setdefault(f"rel_analise_{suf}_{i}", "")


def _relatorio_completo(ss, linhas=6):
    ss.update({"rel_nome_condominio": "Condomínio Fictício Alfa", "rel_endereco_condominio": "Rua Teste, 100",
               "rel_representante": "Síndico Fictício", "rel_mes_referencia": "09", "rel_ano_referencia": "2026",
               "rel_diagnostico": "Água dentro dos parâmetros.", "rel_observacoes_gerais": "Observação fictícia.",
               "rel_art_numero": "ART-0000", "rel_analises_total": linhas, "rel_obs_cmaq": "Casa de máquinas ok",
               "rel_dos_produto_0": "Cloro granulado", "rel_dos_qtd_0": "2", "rel_dos_un_0": "kg",
               "rel_fotos_upload": "[<UploadedFile fictício>]"})
    for i in range(linhas):
        ss.update({f"rel_analise_data_{i}": f"0{i + 1}/09/2026", f"rel_analise_ph_{i}": f"7.{i}",
                   f"rel_analise_cl_{i}": "1.5", f"rel_analise_ct_{i}": "1.8", f"rel_analise_operador_{i}": "Op. Teste"})


class TestAutosave(unittest.TestCase):
    def setUp(self):
        self.amb = Ambiente()
        self.s = self.amb.sessao()
        self.s["abrir_formulario"]()
        self.salvar = self.s["_relatorio_rt_salvar_rascunho"]

    def test_01_abrir_formulario_sem_alteracoes_nao_grava(self):
        for _ in range(3):  # 1ª renderização + reruns
            self.assertFalse(self.salvar("autosave_apos_renderizar_linhas"))
        self.assertIsNone(self.amb.linha("thyago"))
        self.assertFalse((self.amb.dir / "_rascunho_relatorio_rt_em_andamento.json").exists())

    def test_campos_automaticos_nao_contam_como_conteudo(self):
        self.assertFalse(self.s["_relatorio_rt_tem_dados_em_tela"]())

    def test_widget_criado_depois_com_valor_padrao_nao_grava(self):
        self.salvar()
        self.s["st"].session_state.update({"rel_nbr11238_chuveiro": "Sim", "rel_epi_luvas_status": "Conforme",
                                           "rel_dos_produto_0": ""})
        self.assertFalse(self.salvar())
        self.assertIsNone(self.amb.linha("thyago"))
        self.s["st"].session_state["rel_nbr11238_chuveiro"] = "Não"  # escolha do usuário depois
        self.assertTrue(self.salvar())

    def test_campo_de_conteudo_que_surge_preenchido_grava(self):
        self.salvar()
        self.s["st"].session_state["rel_analise_ph_7"] = "7.6"  # ex.: linha nova vinda de importação
        self.assertTrue(self.salvar())

    def test_02_alterar_campo_grava(self):
        self.salvar()
        self.s["st"].session_state["rel_analise_ph_0"] = "7.4"
        self.assertTrue(self.salvar())
        self.assertEqual(self.amb.linha("thyago")["campos"]["rel_analise_ph_0"], "7.4")
        self.assertEqual(self.amb.aba().gravacoes, 2)  # cabeçalho + rascunho

    def test_03_rerun_sem_alteracao_nao_grava_de_novo(self):
        self.salvar()
        self.s["st"].session_state["rel_diagnostico"] = "Texto"
        self.salvar()
        antes = self.amb.aba().gravacoes
        for _ in range(3):
            self.assertFalse(self.salvar())
        self.assertEqual(self.amb.aba().gravacoes, antes)


class TestRestauracao(unittest.TestCase):
    def _salvar_completo(self, amb, linhas=6):
        a = amb.sessao()
        a["abrir_formulario"]()
        a["_relatorio_rt_salvar_rascunho"]()
        _relatorio_completo(a["st"].session_state, linhas)
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        return a

    def _restaurar_em_nova_sessao(self, amb):
        b = amb.sessao()
        rasc = b["_relatorio_rt_carregar_rascunho"]()
        b["st"].session_state["_rt_rasc_pendente"] = {"dados": rasc, "base_rev": b["_relatorio_rt_rev"](rasc)}
        b["_relatorio_rt_aplicar_pendente"]()  # topo do script, antes dos widgets
        return b, rasc

    def test_04_e_05_restaurar_relatorio_completo_com_varias_linhas(self):
        amb = Ambiente()
        a = self._salvar_completo(amb, linhas=6)
        b, rasc = self._restaurar_em_nova_sessao(amb)
        ss_a, ss_b = a["st"].session_state, b["st"].session_state
        for k, v in rasc["campos"].items():
            if k in ("rel_fotos_upload",):
                continue
            self.assertEqual(ss_b.get(k), v, k)
        for i in range(6):
            self.assertEqual(ss_b[f"rel_analise_ph_{i}"], ss_a[f"rel_analise_ph_{i}"])
            self.assertEqual(ss_b[f"rel_analise_data_{i}"], ss_a[f"rel_analise_data_{i}"])
        self.assertEqual(ss_b["rel_nome_condominio"], "Condomínio Fictício Alfa")
        self.assertEqual(ss_b["rel_diagnostico"], "Água dentro dos parâmetros.")
        self.assertNotIn("rel_fotos_upload", ss_b)  # upload não aceita valor via session_state
        self.assertNotIn("_relatorio_rt_restore_erro", ss_b)

    def test_restaurar_nao_dispara_gravacao_imediata(self):
        amb = Ambiente()
        self._salvar_completo(amb)
        b, _ = self._restaurar_em_nova_sessao(amb)
        antes = amb.aba().gravacoes
        self.assertFalse(b["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.aba().gravacoes, antes)
        b["st"].session_state["rel_analise_ph_0"] = "7.8"  # edição depois de restaurar vai para o principal
        self.assertTrue(b["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.linha("thyago")["campos"]["rel_analise_ph_0"], "7.8")

    def test_06_formulario_vazio_nao_sobrescreve_completo(self):
        amb = Ambiente()
        self._salvar_completo(amb)
        original = amb.linha("thyago")
        c = amb.sessao()
        c["abrir_formulario"]()
        c["_relatorio_rt_salvar_rascunho"]()
        c["st"].session_state["rel_mes_referencia"] = "11"  # muda algo, mas sem conteúdo
        self.assertFalse(c["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.linha("thyago"), original)

    def test_07_rascunho_parcial_recuperavel(self):
        amb = Ambiente()
        a = amb.sessao()
        a["abrir_formulario"]()
        a["_relatorio_rt_salvar_rascunho"]()
        a["st"].session_state.update({"rel_nome_condominio": "Residencial Fictício", "rel_analise_ph_0": "7.2"})
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        b, _ = self._restaurar_em_nova_sessao(amb)
        self.assertEqual(b["st"].session_state["rel_nome_condominio"], "Residencial Fictício")
        self.assertEqual(b["st"].session_state["rel_analise_ph_0"], "7.2")
        self.assertEqual(b["st"].session_state["rel_analise_ph_1"], "")

    def test_11_logout_nao_sobrescreve_completo(self):
        amb = Ambiente()
        self._salvar_completo(amb)
        original = amb.linha("thyago")
        d = amb.sessao()
        d["abrir_formulario"]()
        d["_relatorio_rt_salvar_rascunho"]("autosave_apos_renderizar_linhas")
        self.assertFalse(d["_relatorio_rt_salvar_rascunho"]("logout_admin"))
        self.assertEqual(amb.linha("thyago"), original)


class TestDuasSessoes(unittest.TestCase):
    def test_08_sessoes_simultaneas_sem_perda_silenciosa(self):
        amb = Ambiente()
        a, b = amb.sessao(), amb.sessao()
        for s in (a, b):
            s["abrir_formulario"]()
            s["_relatorio_rt_salvar_rascunho"]()  # ambas abrem a tela ao mesmo tempo
        a["st"].session_state.update({"rel_nome_condominio": "Condomínio A", "rel_analise_ph_0": "7.1"})
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        b["st"].session_state.update({"rel_nome_condominio": "Condomínio B", "rel_analise_ph_0": "7.9"})
        self.assertTrue(b["_relatorio_rt_salvar_rascunho"]())

        self.assertEqual(amb.linha("thyago")["campos"]["rel_nome_condominio"], "Condomínio A")
        chave_b = b["st"].session_state["_rt_rasc_conflito"]
        self.assertTrue(chave_b.startswith("thyago#conflito#"))
        self.assertEqual(amb.linha(chave_b)["campos"]["rel_nome_condominio"], "Condomínio B")

        b["st"].session_state["rel_analise_ph_1"] = "7.5"  # B continua na própria cópia
        b["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(amb.linha("thyago")["campos"]["rel_nome_condominio"], "Condomínio A")
        self.assertEqual(amb.linha(chave_b)["campos"]["rel_analise_ph_1"], "7.5")

        a["st"].session_state["rel_analise_ph_1"] = "7.3"  # A segue no principal normalmente
        a["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(amb.linha("thyago")["campos"]["rel_analise_ph_1"], "7.3")

        rascunhos, _ = a["_relatorio_rt_ler_rascunhos"](amb.aba())
        self.assertEqual(set(rascunhos), {"thyago", chave_b})


class TestCodespaces(unittest.TestCase):
    def setUp(self):
        self._antes = os.environ.get("CODESPACES")
        os.environ["CODESPACES"] = "true"

    def tearDown(self):
        if self._antes is None:
            os.environ.pop("CODESPACES", None)
        else:
            os.environ["CODESPACES"] = self._antes

    def test_09_codespaces_nao_acessa_planilha(self):
        amb = Ambiente(codespaces=True)
        s = amb.sessao()
        s["abrir_formulario"]()
        s["_relatorio_rt_salvar_rascunho"]()
        _relatorio_completo(s["st"].session_state)
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(s["_relatorio_rt_carregar_rascunho"]())
        self.assertEqual(amb.planilha.conexoes, 0)
        self.assertIsNone(amb.aba())
        self.assertTrue((amb.dir / "_rascunho_relatorio_rt_em_andamento.json").exists())

    def test_rascunhos_de_operador_tambem_isolados(self):
        amb = Ambiente(codespaces=True)
        s = amb.sessao()
        s["salvar_rascunho_operador"]("Condomínio Fictício", {"obs": "x"}, salvar_sheets=True)
        s["carregar_rascunho_operador"]("Outro Condomínio")
        s["deletar_rascunho_operador"]("Condomínio Fictício")
        self.assertEqual(amb.planilha.conexoes, 0)

    def test_producao_continua_usando_planilha(self):
        os.environ.pop("CODESPACES", None)
        amb = Ambiente()
        s = amb.sessao()
        s["abrir_formulario"]()
        s["_relatorio_rt_salvar_rascunho"]()
        s["st"].session_state["rel_analise_ph_0"] = "7.4"
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertGreater(amb.planilha.conexoes, 0)
        self.assertIsNotNone(amb.linha("thyago"))


class TestRascunhoGrande(unittest.TestCase):
    def test_10_rascunho_grande_nao_gera_json_invalido(self):
        amb = Ambiente()
        s = amb.sessao()
        ss = s["st"].session_state
        s["abrir_formulario"]()
        s["_relatorio_rt_salvar_rascunho"]()
        ss["rel_diagnostico"] = "Versão válida"
        s["_relatorio_rt_salvar_rascunho"]()
        ss["rel_observacoes_gerais"] = "x" * 60000
        s["_relatorio_rt_salvar_rascunho"]()
        for row in amb.aba().linhas[1:]:
            json.loads(row[2])  # nenhuma célula com JSON inválido
        self.assertEqual(amb.linha("thyago")["campos"]["rel_diagnostico"], "Versão válida")
        self.assertNotIn("rel_observacoes_gerais", {k for k, v in amb.linha("thyago")["campos"].items() if v})
        self.assertIn("limite da planilha", ss["_rt_rasc_aviso"])
        local = json.loads((amb.dir / "_rascunho_relatorio_rt_em_andamento.json").read_text(encoding="utf-8"))
        self.assertEqual(len(local["campos"]["rel_observacoes_gerais"]), 60000)

    def test_rascunho_antigo_truncado_e_ignorado(self):
        amb = Ambiente()
        aba = amb.planilha.add_worksheet("_Rascunhos_RT", 100, 3)
        aba.linhas = [["Usuario", "Salvo em", "Dados JSON"], ["thyago", "01/01/2026", '{"campos": {"rel_ana...']]
        s = amb.sessao()
        self.assertEqual(s["_relatorio_rt_carregar_rascunho"](), {})


if __name__ == "__main__":
    unittest.main()
