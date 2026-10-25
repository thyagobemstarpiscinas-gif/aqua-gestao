"""Testes do rascunho do Relatório RT (autosave, proteção, restauração e descarte).

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
    "_relatorio_rt_coletar_rascunho", "_relatorio_rt_chave_sessao", "_relatorio_rt_arquivo_local",
    "_relatorio_rt_aba", "_relatorio_rt_ler_json", "_relatorio_rt_serializar", "_relatorio_rt_desserializar",
    "_relatorio_rt_eh_chave", "_relatorio_rt_ordem", "_relatorio_rt_listar_rascunhos", "_relatorio_rt_descartada",
    "_relatorio_rt_gravar_linha", "_relatorio_rt_enviar_planilha", "_relatorio_rt_sincronizar_pendentes",
    "_relatorio_rt_coluna", "_relatorio_rt_salvar_rascunho", "_relatorio_rt_carregar_rascunho",
    "_relatorio_rt_descartar", "_relatorio_rt_aplicar_rascunho", "_relatorio_rt_aplicar_pendente",
    "_relatorio_rt_resumo",
]
FUNCOES_OPERADOR = ["salvar_rascunho_operador", "carregar_rascunho_operador", "deletar_rascunho_operador"]
PREFIXO_CONSTANTES = ("RASCUNHO_RT_", "_RASCUNHO_RT_")


def _nos(nomes_funcoes):
    nos = [n for n in TREE.body if isinstance(n, ast.Assign)
           and any(isinstance(t, ast.Name) and t.id.startswith(PREFIXO_CONSTANTES) for t in n.targets)]
    for nome in nomes_funcoes:
        f = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == nome)
        f.decorator_list = []
        nos.append(f)
    return nos


class _Aba:
    """Planilha em memória com a interface do gspread usada pelo app."""

    def __init__(self, titulo, cols=4):
        self.titulo = titulo
        self.linhas = []
        self.col_count = cols
        self.updates = 0
        self.appends = 0
        self.falhar = False
        self.antes_de_gravar = None  # gancho para simular outra sessão no meio da gravação

    def get_all_values(self):
        return copy.deepcopy(self.linhas)

    def _gancho(self):
        if self.antes_de_gravar:
            gancho, self.antes_de_gravar = self.antes_de_gravar, None
            gancho()
        if self.falhar:
            raise ConnectionError("falha simulada da planilha")

    def update(self, *args, range_name=None, values=None, **_):
        if args:
            range_name, values = args[0], args[1]
        self._gancho()
        ini, fim = (range_name.split(":") + [range_name])[:2]
        linha = int(re.match(r"[A-Z]+(\d+)", ini).group(1))
        col_ini = _col_num(re.match(r"([A-Z]+)", ini).group(1))
        if _col_num(re.match(r"([A-Z]+)", fim).group(1)) > self.col_count:
            raise ValueError("range além da grade")
        while len(self.linhas) < linha:
            self.linhas.append([])
        row = self.linhas[linha - 1]
        while len(row) < col_ini - 1 + len(values[0]):
            row.append("")
        row[col_ini - 1:col_ini - 1 + len(values[0])] = list(values[0])
        self.updates += 1

    def batch_update(self, blocos, **_):
        self._gancho()
        antes, gancho = self.updates, self.antes_de_gravar
        for b in blocos:  # uma única chamada: todos os intervalos juntos
            self.update(range_name=b["range"], values=b["values"])
        self.updates = antes + 1

    def append_row(self, values, **_):
        self._gancho()
        self.linhas.append(list(values))  # values.append: sempre uma linha nova
        self.col_count = max(self.col_count, len(values))
        self.appends += 1

    def add_cols(self, n):
        self.col_count += n


def _col_num(letras):
    n = 0
    for c in letras:
        n = n * 26 + ord(c) - 64
    return n


class _Planilha:
    def __init__(self):
        self.abas = {}
        self.conexoes = 0

    def add_worksheet(self, title, rows, cols):
        if title in self.abas:
            raise ValueError("aba já existe")
        self.abas[title] = _Aba(title, cols)
        return self.abas[title]

    def worksheet(self, title):
        if title not in self.abas:
            raise KeyError(title)
        return self.abas[title]


class _St:
    def __init__(self):
        self.session_state = {}


class Ambiente:
    """Armazenamento compartilhado (planilha + disco do servidor) entre sessões simuladas."""

    def __init__(self):
        self.dir = Path(tempfile.mkdtemp(prefix="rascunho_rt_teste_"))
        atexit.register(shutil.rmtree, self.dir, ignore_errors=True)
        self.planilha = _Planilha()
        self._relogio = 0

    def agora(self):
        self._relogio += 1
        return f"25/10/2026 10:{self._relogio // 60:02d}:{self._relogio % 60:02d}"

    def conectar(self):
        self.planilha.conexoes += 1
        return self.planilha

    def obter_aba(self, nome):
        return self.planilha.abas.get(nome)  # como no app: None quando a aba não existe

    def sessao(self, abrir=True):
        st = _St()
        ns = {
            "st": st, "os": os, "re": re, "json": json, "hashlib": hashlib, "datetime": datetime, "Path": Path,
            "ANALISES_PADRAO": 12, "GENERATED_DIR": self.dir, "slugify_nome": slugify_nome,
            "_agora_brasilia": self.agora, "conectar_sheets": self.conectar, "obter_aba_sheets": self.obter_aba,
            "_relatorio_rt_rascunho_path": lambda: self.dir / "_rascunho_relatorio_rt_em_andamento.json",
        }
        exec(compile(ast.Module(body=_nos(FUNCOES_RT + FUNCOES_OPERADOR), type_ignores=[]), str(APP), "exec"), ns)
        ns["ss"] = st.session_state
        if abrir:
            _abrir_formulario(st.session_state)
            ns["_relatorio_rt_salvar_rascunho"]()  # 1ª renderização registra a referência
        return ns

    def aba(self):
        return self.planilha.abas.get("_Rascunhos_RT")

    def criar_aba(self):
        aba = self.planilha.add_worksheet("_Rascunhos_RT", 100, 3)
        aba.linhas = [["Usuario", "Salvo em", "Dados JSON"]]  # cabeçalho no formato antigo
        return aba

    def linhas(self, chave=None):
        aba = self.aba()
        return [r for r in (aba.linhas[1:] if aba else []) if chave is None or (r and r[0] == chave)]

    def campos(self, sessao):
        rows = self.linhas(sessao["ss"]["_rt_rasc_chave"])
        assert len(rows) == 1, rows
        return sessao["_relatorio_rt_desserializar"](rows[0])["campos"]

    def limpar_disco(self):
        for p in self.dir.glob("*.json"):
            p.unlink()


def _abrir_formulario(ss):
    """Estado de um formulário recém-aberto: só valores que o app preenche sozinho."""
    ss.update({"rel_status_agua": "CONFORME", "rel_status_agua_select": "CONFORME", "rel_verificacoes_semanais": 1,
               "rel_analises_total": 4, "rel_mes_referencia": "10", "rel_ano_referencia": "2026",
               "rel_tipo_atendimento": "Contrato ativo"})
    for i in range(4):
        for suf in ("data", "ph", "cl", "ct", "operador"):
            ss.setdefault(f"rel_analise_{suf}_{i}", "")


def _relatorio_completo(ss, linhas=6, nome="Condomínio Fictício Alfa"):
    ss.update({"rel_nome_condominio": nome, "rel_endereco_condominio": "Rua Teste, 100",
               "rel_representante": "Síndico Fictício", "rel_mes_referencia": "09", "rel_ano_referencia": "2026",
               "rel_diagnostico": "Água dentro dos parâmetros.", "rel_observacoes_gerais": "Observação fictícia.",
               "rel_art_numero": "ART-0000", "rel_analises_total": linhas, "rel_obs_cmaq": "Casa de máquinas ok",
               "rel_dos_produto_0": "Cloro granulado", "rel_dos_qtd_0": "2", "rel_dos_un_0": "kg",
               "rel_fotos_upload": "[<UploadedFile fictício>]"})
    for i in range(linhas):
        ss.update({f"rel_analise_data_{i}": f"0{i + 1}/09/2026", f"rel_analise_ph_{i}": f"7.{i}",
                   f"rel_analise_cl_{i}": "1.5", f"rel_analise_ct_{i}": "1.8", f"rel_analise_operador_{i}": "Op. Teste"})


def _restaurar(sessao, versao_dados):
    sessao["ss"]["_rt_rasc_pendente"] = {"dados": versao_dados}
    sessao["_relatorio_rt_aplicar_pendente"]()  # topo do script, antes dos widgets


class TestAutosave(unittest.TestCase):
    def setUp(self):
        self.amb = Ambiente()
        self.s = self.amb.sessao()
        self.salvar = self.s["_relatorio_rt_salvar_rascunho"]

    def test_01_abrir_formulario_sem_alteracoes_nao_grava(self):
        for _ in range(3):
            self.assertFalse(self.salvar("autosave_apos_renderizar_linhas"))
        self.assertEqual(self.amb.linhas(), [])
        self.assertEqual(list(self.amb.dir.glob("*.json")), [])

    def test_campos_automaticos_nao_contam_como_conteudo(self):
        self.assertFalse(self.s["_relatorio_rt_tem_dados_em_tela"]())

    def test_widget_criado_depois_com_valor_padrao_nao_grava(self):
        self.s["ss"].update({"rel_nbr11238_chuveiro": "Sim", "rel_epi_luvas_status": "Conforme", "rel_dos_produto_0": ""})
        self.assertFalse(self.salvar())
        self.s["ss"]["rel_nbr11238_chuveiro"] = "Não"
        self.s["ss"]["rel_analise_ph_0"] = "7.0"
        self.assertTrue(self.salvar())

    def test_campo_de_conteudo_que_surge_preenchido_grava(self):
        self.s["ss"]["rel_analise_ph_7"] = "7.6"
        self.assertTrue(self.salvar())

    def test_02_alterar_campo_grava(self):
        self.s["ss"]["rel_analise_ph_0"] = "7.4"
        self.assertTrue(self.salvar())
        self.assertEqual(self.amb.campos(self.s)["rel_analise_ph_0"], "7.4")

    def test_03_rerun_sem_alteracao_nao_grava_de_novo(self):
        self.s["ss"]["rel_diagnostico"] = "Texto"
        self.salvar()
        antes = (self.amb.aba().updates, self.amb.aba().appends)
        for _ in range(3):
            self.assertFalse(self.salvar())
        self.assertEqual((self.amb.aba().updates, self.amb.aba().appends), antes)

    def test_formulario_vazio_nunca_e_gravado(self):
        self.s["ss"]["rel_mes_referencia"] = "11"
        self.assertFalse(self.salvar())
        self.assertEqual(self.amb.linhas(), [])


class TestSessoes(unittest.TestCase):
    def test_s1_duas_sessoes_gravando_simultaneamente(self):
        amb = Ambiente()
        a, b = amb.sessao(), amb.sessao()
        a["ss"].update({"rel_nome_condominio": "Condomínio A", "rel_analise_ph_0": "7.1"})
        b["ss"].update({"rel_nome_condominio": "Condomínio B", "rel_analise_ph_0": "7.9"})
        c = amb.sessao()
        c["ss"]["rel_obs_x"] = "semente"
        c["_relatorio_rt_salvar_rascunho"]()  # aba já existe
        amb.aba().antes_de_gravar = lambda: b["_relatorio_rt_salvar_rascunho"]()  # B grava no meio da gravação de A
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.campos(a)["rel_nome_condominio"], "Condomínio A")
        self.assertEqual(amb.campos(b)["rel_nome_condominio"], "Condomínio B")
        self.assertEqual(amb.campos(c)["rel_obs_x"], "semente")

    def test_s2_duas_sessoes_novas_criando_linhas_ao_mesmo_tempo(self):
        amb = Ambiente()
        amb.criar_aba()
        a, b = amb.sessao(), amb.sessao()
        a["ss"]["rel_analise_ph_0"] = "7.1"
        b["ss"]["rel_analise_ph_0"] = "7.2"
        amb.aba().antes_de_gravar = lambda: b["_relatorio_rt_salvar_rascunho"]()
        a["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(len(amb.linhas()), 2)
        self.assertEqual(amb.aba().appends, 2)
        self.assertNotEqual(a["ss"]["_rt_rasc_chave"], b["ss"]["_rt_rasc_chave"])

    def test_s3_atualizacao_repetida_da_mesma_sessao(self):
        amb = Ambiente()
        s = amb.sessao()
        chave = None
        for i in range(5):
            s["ss"]["rel_analise_ph_0"] = f"7.{i}"
            self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
            chave = chave or s["ss"]["_rt_rasc_chave"]
            self.assertEqual(s["ss"]["_rt_rasc_chave"], chave)  # chave estável
        self.assertEqual(len(amb.linhas(chave)), 1)
        self.assertEqual(amb.campos(s)["rel_analise_ph_0"], "7.4")
        self.assertEqual(amb.aba().appends, 1)

    def test_s3b_gravacao_simultanea_da_mesma_sessao_nao_perde_edicao(self):
        amb = Ambiente()
        s = amb.sessao()
        s["ss"]["rel_analise_ph_0"] = "7.1"
        s["ss"]["_rt_rasc_gravando"] = True  # outra gravação desta sessão em andamento
        self.assertFalse(s["_relatorio_rt_salvar_rascunho"]())
        s["ss"]["_rt_rasc_gravando"] = False
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())  # próximo rerun grava a edição pendente
        self.assertEqual(amb.campos(s)["rel_analise_ph_0"], "7.1")


class TestLegado(unittest.TestCase):
    def _amb_legado(self):
        amb = Ambiente()
        aba = amb.criar_aba()
        legado = {"salvo_em": "24/10/2026 18:00:00", "motivo": "autosave", "rel_analises_total": 12,
                  "campos": {"rel_nome_condominio": "Condomínio Legado", "rel_analise_ph_0": "7.3", "rel_diagnostico": "Antigo"}}
        aba.linhas.append(["thyago", legado["salvo_em"], json.dumps(legado, ensure_ascii=False)])
        return amb, copy.deepcopy(aba.linhas)

    def test_l4_le_e_restaura_rascunho_legado(self):
        amb, _ = self._amb_legado()
        s = amb.sessao()
        versoes = s["_relatorio_rt_listar_rascunhos"]()["versoes"]
        self.assertEqual(versoes[0]["chave"], "thyago")
        _restaurar(s, versoes[0]["dados"])
        self.assertEqual(s["ss"]["rel_nome_condominio"], "Condomínio Legado")
        self.assertIn("Condomínio Legado", s["_relatorio_rt_resumo"](versoes[0]["dados"]))

    def test_l11_integridade_dos_dados_anteriores(self):
        amb, original = self._amb_legado()
        a, b = amb.sessao(), amb.sessao()
        _restaurar(a, a["_relatorio_rt_carregar_rascunho"]())
        a["_relatorio_rt_salvar_rascunho"]()
        a["ss"]["rel_analise_ph_1"] = "7.5"
        a["_relatorio_rt_salvar_rascunho"]()
        b["ss"]["rel_diagnostico"] = "Outro"
        b["_relatorio_rt_salvar_rascunho"]()
        b["_relatorio_rt_descartar"](b["ss"]["_rt_rasc_chave"])
        self.assertEqual(amb.aba().linhas[1][:3], original[1])  # linha antiga intacta
        self.assertEqual(amb.campos(a)["rel_nome_condominio"], "Condomínio Legado")


class TestRascunhoGrande(unittest.TestCase):
    def _grande(self, amb, n=120000):
        s = amb.sessao()
        s["ss"]["rel_observacoes_gerais"] = "x" * n
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
        return s

    def test_g5_dividido_em_partes_e_reconstruido(self):
        amb = Ambiente()
        s = self._grande(amb)
        row = amb.linhas(s["ss"]["_rt_rasc_chave"])[0]
        self.assertTrue(row[2].startswith("__PARTES__:"))
        n = int(row[2].split(":")[1])
        self.assertGreaterEqual(n, 4)
        self.assertTrue(all(len(c) <= 40000 for c in row))
        self.assertEqual(len(amb.campos(s)["rel_observacoes_gerais"]), 120000)
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))

    def test_g5b_reduzir_tamanho_limpa_partes_antigas(self):
        amb = Ambiente()
        s = self._grande(amb)
        s["ss"]["rel_observacoes_gerais"] = "curto"
        s["_relatorio_rt_salvar_rascunho"]()
        row = amb.linhas(s["ss"]["_rt_rasc_chave"])[0]
        self.assertFalse(row[2].startswith("__PARTES__"))
        self.assertTrue(all(c == "" for c in row[4:]))
        self.assertEqual(amb.campos(s)["rel_observacoes_gerais"], "curto")

    def test_g6_parte_ausente_ou_alterada_nao_e_restaurada(self):
        for corromper in ("alterar", "remover"):
            amb = Ambiente()
            s = self._grande(amb)
            row = amb.aba().linhas[-1]
            if corromper == "alterar":
                row[5] = row[5][:-1] + "y"
            else:
                del row[6:]
            nova = amb.sessao(abrir=False)
            amb.limpar_disco()  # sem cópia local, só a planilha
            lista = nova["_relatorio_rt_listar_rascunhos"]()
            self.assertEqual(lista["versoes"], [], corromper)
            self.assertEqual(lista["ignorados"], 1, corromper)

    def test_g7_falha_ao_gravar_rascunho_grande_preserva_versao_valida(self):
        amb = Ambiente()
        s = amb.sessao()
        s["ss"]["rel_diagnostico"] = "versão válida"
        s["_relatorio_rt_salvar_rascunho"]()
        antes = copy.deepcopy(amb.linhas(s["ss"]["_rt_rasc_chave"]))
        amb.aba().falhar = True
        s["ss"]["rel_observacoes_gerais"] = "x" * 90000
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(amb.linhas(s["ss"]["_rt_rasc_chave"]), antes)  # planilha intacta
        self.assertIn("falha simulada", s["ss"]["_rt_rasc_aviso"])  # aviso explícito no painel
        versoes = s["_relatorio_rt_listar_rascunhos"]()["versoes"]
        self.assertEqual(versoes[0]["origem"], "local")  # cópia temporária identificada como tal
        amb.aba().falhar = False
        s["ss"]["rel_observacoes_gerais"] += "!"
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))
        self.assertEqual(len(amb.campos(s)["rel_observacoes_gerais"]), 90001)

    def test_g8_reinicio_do_app_restaura_pela_planilha(self):
        amb = Ambiente()
        s = self._grande(amb)
        _relatorio_completo(s["ss"], linhas=6)
        s["_relatorio_rt_salvar_rascunho"]()
        amb.limpar_disco()  # reinício do Streamlit Cloud
        nova = amb.sessao()
        rasc = nova["_relatorio_rt_carregar_rascunho"]()
        _restaurar(nova, rasc)
        self.assertEqual(len(nova["ss"]["rel_observacoes_gerais"]), len("Observação fictícia."))
        for i in range(6):
            self.assertEqual(nova["ss"][f"rel_analise_ph_{i}"], f"7.{i}")
        self.assertEqual(nova["ss"]["rel_nome_condominio"], "Condomínio Fictício Alfa")

    def test_rascunho_antigo_truncado_e_ignorado(self):
        amb = Ambiente()
        aba = amb.criar_aba()
        aba.linhas.append(["thyago", "01/01/2026", '{"campos": {"rel_ana...'])
        s = amb.sessao(abrir=False)
        self.assertEqual(s["_relatorio_rt_carregar_rascunho"](), {})


class TestRestauracao(unittest.TestCase):
    def test_04_e_05_restaurar_relatorio_completo_com_varias_linhas(self):
        amb = Ambiente()
        a = amb.sessao()
        _relatorio_completo(a["ss"], linhas=6)
        a["_relatorio_rt_salvar_rascunho"]()
        b = amb.sessao(abrir=False)
        rasc = b["_relatorio_rt_carregar_rascunho"]()
        _restaurar(b, rasc)
        for k, v in rasc["campos"].items():
            if k != "rel_fotos_upload":
                self.assertEqual(b["ss"].get(k), v, k)
        self.assertNotIn("rel_fotos_upload", b["ss"])
        self.assertNotIn("_relatorio_rt_restore_erro", b["ss"])

    def test_restaurar_nao_grava_e_edicao_seguinte_vai_para_linha_propria(self):
        amb = Ambiente()
        a = amb.sessao()
        _relatorio_completo(a["ss"])
        a["_relatorio_rt_salvar_rascunho"]()
        b = amb.sessao(abrir=False)
        _restaurar(b, b["_relatorio_rt_carregar_rascunho"]())
        self.assertFalse(b["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(len(amb.linhas()), 1)
        b["ss"]["rel_analise_ph_0"] = "7.8"
        self.assertTrue(b["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.campos(a)["rel_analise_ph_0"], "7.0")  # versão de A preservada
        self.assertEqual(amb.campos(b)["rel_analise_ph_0"], "7.8")

    def test_07_rascunho_parcial_recuperavel(self):
        amb = Ambiente()
        a = amb.sessao()
        a["ss"].update({"rel_nome_condominio": "Residencial Fictício", "rel_analise_ph_0": "7.2"})
        a["_relatorio_rt_salvar_rascunho"]()
        b = amb.sessao(abrir=False)
        _restaurar(b, b["_relatorio_rt_carregar_rascunho"]())
        self.assertEqual(b["ss"]["rel_nome_condominio"], "Residencial Fictício")
        self.assertEqual(b["ss"]["rel_analise_ph_1"], "")

    def test_11_logout_nao_grava_formulario_vazio(self):
        amb = Ambiente()
        a = amb.sessao()
        _relatorio_completo(a["ss"])
        a["_relatorio_rt_salvar_rascunho"]()
        original = copy.deepcopy(amb.linhas())
        d = amb.sessao()
        self.assertFalse(d["_relatorio_rt_salvar_rascunho"]("logout_admin"))
        self.assertEqual(amb.linhas(), original)


class TestDescarte(unittest.TestCase):
    def _duas_versoes(self):
        amb = Ambiente()
        a, b = amb.sessao(), amb.sessao()
        _relatorio_completo(a["ss"], nome="Condomínio A")
        a["_relatorio_rt_salvar_rascunho"]()
        _relatorio_completo(b["ss"], nome="Condomínio B")
        b["_relatorio_rt_salvar_rascunho"]()
        return amb, a, b

    def test_d9_descarte_confirmado_marca_somente_aquela_versao(self):
        amb, a, b = self._duas_versoes()
        chave_b = b["ss"]["_rt_rasc_chave"]
        linha_a = copy.deepcopy(amb.linhas(a["ss"]["_rt_rasc_chave"]))
        conteudo_b = amb.linhas(chave_b)[0][2]
        a["_relatorio_rt_descartar"](chave_b)
        row_b = amb.linhas(chave_b)[0]
        self.assertTrue(row_b[3].startswith("descartado em"))
        self.assertEqual(row_b[2], conteudo_b)  # conteúdo preservado para recuperação
        self.assertEqual(amb.linhas(a["ss"]["_rt_rasc_chave"]), linha_a)
        nomes = [v["dados"]["campos"]["rel_nome_condominio"] for v in a["_relatorio_rt_listar_rascunhos"]()["versoes"]]
        self.assertEqual(nomes, ["Condomínio A"])
        todas = a["_relatorio_rt_listar_rascunhos"](incluir_descartados=True)["versoes"]
        self.assertEqual(len(todas), 2)

    def test_d9b_descarte_cancelado_nao_altera_nada(self):
        amb, a, b = self._duas_versoes()
        antes = copy.deepcopy(amb.aba().linhas)
        a["ss"]["_rt_rasc_confirmar_descarte"] = b["ss"]["_rt_rasc_chave"]
        a["ss"].pop("_rt_rasc_confirmar_descarte")  # "Cancelar"
        self.assertEqual(amb.aba().linhas, antes)

    def test_d10_descartado_nao_reaparece_e_nao_e_regravado(self):
        amb, a, b = self._duas_versoes()
        chave_b = b["ss"]["_rt_rasc_chave"]
        b["_relatorio_rt_descartar"](chave_b)  # descarta a própria versão
        self.assertNotEqual(b["ss"]["_rt_rasc_chave"], chave_b)
        for _ in range(3):
            self.assertFalse(b["_relatorio_rt_salvar_rascunho"]())  # reruns não regravam
        self.assertEqual(len(amb.linhas()), 2)
        amb.limpar_disco()
        nova = amb.sessao()  # novo login
        nomes = [v["dados"]["campos"]["rel_nome_condominio"] for v in nova["_relatorio_rt_listar_rascunhos"]()["versoes"]]
        self.assertNotIn("Condomínio B", nomes)
        b["ss"]["rel_analise_ph_0"] = "6.9"  # nova edição vira nova versão
        self.assertTrue(b["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(amb.linhas(chave_b)[0][3].startswith("descartado"))
        self.assertEqual(len(amb.linhas()), 3)

    def test_versao_descartada_por_outro_acesso_continua_descartada(self):
        amb, a, b = self._duas_versoes()
        chave_b = b["ss"]["_rt_rasc_chave"]
        a["_relatorio_rt_descartar"](chave_b)
        b["ss"]["rel_analise_ph_0"] = "6.5"  # B ainda editando
        self.assertTrue(b["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(amb.linhas(chave_b)[0][3].startswith("descartado"))
        self.assertNotEqual(b["ss"]["_rt_rasc_chave"], chave_b)
        self.assertEqual(amb.campos(b)["rel_analise_ph_0"], "6.5")


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
        amb = Ambiente()
        s = amb.sessao()
        _relatorio_completo(s["ss"])
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(s["_relatorio_rt_carregar_rascunho"]())
        s["_relatorio_rt_descartar"](s["ss"]["_rt_rasc_chave"])
        self.assertEqual(amb.planilha.conexoes, 0)
        self.assertIsNone(amb.aba())
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))

    def test_rascunhos_de_operador_tambem_isolados(self):
        amb = Ambiente()
        s = amb.sessao(abrir=False)
        s["salvar_rascunho_operador"]("Condomínio Fictício", {"obs": "x"}, salvar_sheets=True)
        s["carregar_rascunho_operador"]("Outro Condomínio")
        s["deletar_rascunho_operador"]("Condomínio Fictício")
        self.assertEqual(amb.planilha.conexoes, 0)

    def test_producao_continua_usando_planilha(self):
        os.environ.pop("CODESPACES", None)
        amb = Ambiente()
        s = amb.sessao()
        s["ss"]["rel_analise_ph_0"] = "7.4"
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertGreater(amb.planilha.conexoes, 0)
        self.assertEqual(len(amb.linhas()), 1)


class TestAjustesFinais(unittest.TestCase):
    def test_descarte_com_planilha_indisponivel_nao_conclui(self):
        amb = Ambiente()
        a = amb.sessao()
        a["ss"]["rel_diagnostico"] = "versão X"
        a["_relatorio_rt_salvar_rascunho"]()
        chave = a["ss"]["_rt_rasc_chave"]
        antes = copy.deepcopy(amb.aba().linhas)
        real = a["_relatorio_rt_aba"]
        a["_relatorio_rt_aba"] = lambda criar=False: None  # planilha fora do ar
        self.assertFalse(a["_relatorio_rt_descartar"](chave))
        self.assertIn("NÃO foi descartado", a["ss"]["_rt_rasc_erro_descarte"])
        self.assertEqual(amb.aba().linhas, antes)  # planilha intacta
        self.assertEqual(a["ss"]["_rt_rasc_chave"], chave)  # nada mudou na sessão
        local = json.loads((amb.dir / f"_rascunho_relatorio_rt_{slugify_nome(chave)}.json").read_text(encoding="utf-8"))
        self.assertNotIn("_descartado", local)  # cópia local também preservada
        a["_relatorio_rt_aba"] = real  # planilha voltou: nova tentativa
        self.assertTrue(a["_relatorio_rt_descartar"](chave))
        self.assertTrue(amb.linhas(chave)[0][3].startswith("descartado"))
        amb.limpar_disco()
        nova = amb.sessao(abrir=False)
        self.assertFalse(any(v["chave"] == chave for v in nova["_relatorio_rt_listar_rascunhos"]()["versoes"]))

    def test_descarte_com_falha_ao_gravar_marca_preserva_rascunho(self):
        amb = Ambiente()
        a = amb.sessao()
        a["ss"]["rel_diagnostico"] = "versão X"
        a["_relatorio_rt_salvar_rascunho"]()
        chave = a["ss"]["_rt_rasc_chave"]
        amb.aba().falhar = True
        self.assertFalse(a["_relatorio_rt_descartar"](chave))
        amb.aba().falhar = False
        self.assertEqual(amb.linhas(chave)[0][3], "")
        self.assertTrue(any(v["chave"] == chave for v in a["_relatorio_rt_listar_rascunhos"]()["versoes"]))

    def test_restaurar_outra_versao_na_mesma_sessao_preserva_x_e_y(self):
        amb = Ambiente()
        b = amb.sessao()
        b["ss"].update({"rel_nome_condominio": "Y", "rel_analise_ph_0": "7.7"})
        b["_relatorio_rt_salvar_rascunho"]()
        chave_y = b["ss"]["_rt_rasc_chave"]
        linha_y = copy.deepcopy(amb.linhas(chave_y))
        a = amb.sessao()
        a["ss"].update({"rel_nome_condominio": "X", "rel_analise_ph_0": "7.1"})
        a["_relatorio_rt_salvar_rascunho"]()
        chave_x = a["ss"]["_rt_rasc_chave"]
        linha_x = copy.deepcopy(amb.linhas(chave_x))
        y = next(v["dados"] for v in a["_relatorio_rt_listar_rascunhos"]()["versoes"] if v["chave"] == chave_y)
        _restaurar(a, y)
        a["_relatorio_rt_salvar_rascunho"]()
        a["ss"]["rel_analise_ph_1"] = "7.2"
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(amb.linhas(chave_x), linha_x)  # X intacto
        self.assertEqual(amb.linhas(chave_y), linha_y)  # Y original intacto
        nova = amb.campos(a)
        self.assertNotIn(a["ss"]["_rt_rasc_chave"], (chave_x, chave_y))
        self.assertEqual((nova["rel_nome_condominio"], nova["rel_analise_ph_1"]), ("Y", "7.2"))

    def test_descarte_no_meio_do_salvamento_mantem_marca_e_partes_integras(self):
        amb = Ambiente()
        a = amb.sessao()
        a["ss"]["rel_observacoes_gerais"] = "a" * 90000  # versão dividida em partes
        a["_relatorio_rt_salvar_rascunho"]()
        chave = a["ss"]["_rt_rasc_chave"]
        b = amb.sessao(abrir=False)
        a["ss"]["rel_observacoes_gerais"] = "b" * 95000
        amb.aba().antes_de_gravar = lambda: self.assertTrue(b["_relatorio_rt_descartar"](chave))
        self.assertTrue(a["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(amb.linhas(chave)[0][3].startswith("descartado"))  # marca não foi apagada
        nova_chave = a["ss"]["_rt_rasc_chave"]
        self.assertNotEqual(nova_chave, chave)
        self.assertEqual(amb.campos(a)["rel_observacoes_gerais"], "b" * 95000)  # edição em linha ativa, partes válidas
        lista = b["_relatorio_rt_listar_rascunhos"]()
        self.assertEqual([v["chave"] for v in lista["versoes"]], [nova_chave])
        self.assertEqual(lista["ignorados"], 0)
        for row in amb.linhas():
            self.assertIsNotNone(b["_relatorio_rt_desserializar"](row))  # nenhuma linha corrompida

    def test_codespaces_descarte_local_continua_permitido(self):
        os.environ["CODESPACES"] = "true"
        try:
            amb = Ambiente()
            s = amb.sessao()
            s["ss"]["rel_diagnostico"] = "dev"
            s["_relatorio_rt_salvar_rascunho"]()
            chave = s["ss"]["_rt_rasc_chave"]
            self.assertTrue(s["_relatorio_rt_descartar"](chave))
            self.assertEqual(s["_relatorio_rt_listar_rascunhos"]()["versoes"], [])
            self.assertEqual(amb.planilha.conexoes, 0)
        finally:
            os.environ.pop("CODESPACES", None)


class TestSincronizacaoPendente(unittest.TestCase):
    """Falha temporária da planilha: edição fica local, pendente, e é reenviada sem nova edição."""

    def _com_falha(self):
        amb = Ambiente()
        s = amb.sessao()  # 1. abre o formulário
        s["ss"]["rel_diagnostico"] = "versão válida"
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
        chave = s["ss"]["_rt_rasc_chave"]
        s["ss"]["rel_diagnostico"] = "edição durante a falha"  # 2. edita
        amb.aba().falhar = True  # 3. falha temporária
        self.assertFalse(s["_relatorio_rt_salvar_rascunho"]())  # não informa persistência
        return amb, s, chave

    def _local(self, amb, chave):
        return json.loads((amb.dir / f"_rascunho_relatorio_rt_{slugify_nome(chave)}.json").read_text(encoding="utf-8"))

    def _ops(self, amb):
        return amb.aba().updates + amb.aba().appends

    def test_cenario_completo_reenvia_sem_nova_edicao(self):
        amb, s, chave = self._com_falha()
        self.assertEqual(self._local(amb, chave)["campos"]["rel_diagnostico"], "edição durante a falha")  # 4. local salvo
        self.assertEqual(amb.campos(s)["rel_diagnostico"], "versão válida")  # planilha com a versão válida anterior
        self.assertIn("falha simulada", s["ss"]["_rt_rasc_aviso"])
        self.assertEqual(s["ss"]["_rt_rasc_sync_pendente"], [chave])
        amb.aba().falhar = False  # 5. planilha volta
        s["ss"]["_rt_rasc_proxima_tentativa"] = 0  # intervalo entre tentativas já passou
        antes = self._ops(amb)
        s["_relatorio_rt_salvar_rascunho"]()  # 6. rerun sem modificar nada
        self.assertEqual(amb.campos(s)["rel_diagnostico"], "edição durante a falha")  # 7. persistido
        self.assertEqual(self._ops(amb) - antes, 1)
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))
        self.assertIsNone(s["ss"].get("_rt_rasc_sync_pendente"))
        antes = self._ops(amb)
        for _ in range(3):
            s["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(self._ops(amb), antes)  # sem gravações duplicadas
        self.assertEqual(len(amb.linhas()), 1)

    def test_respeita_intervalo_entre_tentativas(self):
        amb, s, chave = self._com_falha()
        amb.aba().falhar = False
        antes = self._ops(amb)
        s["_relatorio_rt_salvar_rascunho"]()  # logo em seguida: ainda dentro do intervalo
        self.assertEqual(self._ops(amb), antes)
        self.assertEqual(s["ss"]["_rt_rasc_sync_pendente"], [chave])
        s["ss"]["_rt_rasc_proxima_tentativa"] = 0
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertIsNone(s["ss"].get("_rt_rasc_sync_pendente"))

    def test_planilha_ainda_fora_mantem_pendencia_e_aviso(self):
        amb, s, chave = self._com_falha()
        s["ss"]["_rt_rasc_proxima_tentativa"] = 0
        s["_relatorio_rt_salvar_rascunho"]()  # nova tentativa falha
        self.assertEqual(s["ss"]["_rt_rasc_sync_pendente"], [chave])
        self.assertTrue(s["ss"]["_rt_rasc_aviso"])
        self.assertGreater(s["ss"]["_rt_rasc_proxima_tentativa"], 0)
        amb.aba().falhar = False
        self.assertEqual(amb.campos(s)["rel_diagnostico"], "versão válida")

    def test_planilha_indisponivel_na_conexao_tambem_fica_pendente(self):
        amb = Ambiente()
        s = amb.sessao()
        real = s["_relatorio_rt_aba"]
        s["_relatorio_rt_aba"] = lambda criar=False: None
        s["ss"]["rel_diagnostico"] = "sem conexão"
        self.assertFalse(s["_relatorio_rt_salvar_rascunho"]())
        self.assertTrue(s["ss"]["_rt_rasc_aviso"])
        s["_relatorio_rt_aba"] = real
        s["ss"]["_rt_rasc_proxima_tentativa"] = 0
        s["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(amb.campos(s)["rel_diagnostico"], "sem conexão")
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))

    def test_nova_edicao_com_pendencia_grava_uma_vez(self):
        amb, s, chave = self._com_falha()
        amb.aba().falhar = False
        s["ss"]["rel_diagnostico"] = "edição depois que a planilha voltou"
        antes = self._ops(amb)
        self.assertTrue(s["_relatorio_rt_salvar_rascunho"]())
        self.assertEqual(self._ops(amb) - antes, 1)  # a versão nova já cobre a pendência da mesma chave
        self.assertEqual(amb.campos(s)["rel_diagnostico"], "edição depois que a planilha voltou")
        self.assertIsNone(s["ss"].get("_rt_rasc_sync_pendente"))
        self.assertIsNone(s["ss"].get("_rt_rasc_aviso"))

    def test_restaurar_outra_versao_com_pendencia_nao_perde_edicao(self):
        amb, s, chave = self._com_falha()
        outra = amb.sessao()
        amb.aba().falhar = False
        outra["ss"]["rel_nome_condominio"] = "Versão Y"
        outra["_relatorio_rt_salvar_rascunho"]()
        chave_y = outra["ss"]["_rt_rasc_chave"]
        linha_y = copy.deepcopy(amb.linhas(chave_y))
        y = next(v["dados"] for v in s["_relatorio_rt_listar_rascunhos"]()["versoes"] if v["chave"] == chave_y)
        _restaurar(s, y)
        s["ss"]["_rt_rasc_proxima_tentativa"] = 0
        for _ in range(2):
            s["_relatorio_rt_salvar_rascunho"]()  # reruns sem edição
        linha_x = amb.linhas(chave)[0]
        self.assertEqual(s["_relatorio_rt_desserializar"](linha_x)["campos"]["rel_diagnostico"], "edição durante a falha")
        self.assertEqual(amb.linhas(chave_y), linha_y)
        self.assertEqual(len(amb.linhas()), 2)  # a versão restaurada não é regravada sem edição
        self.assertIsNone(s["ss"].get("_rt_rasc_sync_pendente"))

    def test_abrir_escritorio_continua_sem_gravar(self):
        amb = Ambiente()
        amb.criar_aba()
        s = amb.sessao()
        for _ in range(3):
            s["_relatorio_rt_salvar_rascunho"]()
        self.assertEqual(self._ops(amb), 0)
        self.assertIsNone(s["ss"].get("_rt_rasc_sync_pendente"))


if __name__ == "__main__":
    unittest.main()
