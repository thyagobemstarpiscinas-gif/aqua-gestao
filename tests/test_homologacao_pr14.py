"""Regressões dos defeitos da homologação do PR #14 (condomínio fictício Épsilon).

Reproduz, com dados FICTÍCIOS e o Google falso do teste de OAuth (escopo
drive.file, sem rede), o caminho real da homologação:

1. PDF individual da visita Bem Star sem a fotografia (arquivo temporário
   apagado antes do doc.build) e sem fotos quando regerado a partir do Sheets.
2. PDF mensal Bem Star com a visita V00005 duplicada (cópia local sem ID +
   cópia do Sheets com ID).
3. PDF mensal com a piscina "Piscina adulta" além da cadastrada "Piscina Adulto".
4. Condomínio Gama (Bem Star) listado na Aqua Gestão por causa de uma pasta
   local sem empresa definida.
5/6. Erro ``csr_sel`` no botão do relatório Bem Star e status da água
   divergente do seletor (ver também test_homologacao_pr14_apptest.py).

Nenhum Secret, credencial, planilha real, Drive real ou lógica de dosagem é
usado ou alterado.
"""
import __future__
import ast
import base64
import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT))

from test_google_oauth_homologacao import _Base  # noqa: E402
from utils import bem_star_mensal as bsm  # noqa: E402
from utils import pdf_bem_star_mensal as pdf_bs  # noqa: E402
from utils.formatacao import slugify_nome  # noqa: E402

EPSILON = "Condomínio Épsilon Fictício"
GAMA = "Condomínio Gama Fictício"
CONSTANTES_APP = {"LOGO_CANDIDATOS", "LOGO_BEM_STAR_CANDIDATOS", "NBR_LIM", "NBR_TXT", "NBR_CC_MAX",
                  # cópia persistente das pastas de condomínio e fotos de campo
                  "_INTERVALO_TENTATIVA_PASTAS_S", "NOME_BASE_PASTAS_PRODUCAO", "NOME_PASTA_FOTOS_PRODUCAO"}


def _foto(cor=(200, 30, 30), tamanho=(640, 480)) -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", tamanho, cor).save(buf, format="JPEG", quality=90)
    return buf.getvalue()


def _imagens_pdf(pdf: bytes) -> list[tuple[int, int]]:
    with fitz.open(stream=pdf, filetype="pdf") as doc:
        return [(img["width"], img["height"]) for page in doc for img in
                (doc.extract_image(x[0]) for x in page.get_images())]


def _texto_pdf(caminho_ou_bytes) -> str:
    if isinstance(caminho_ou_bytes, (bytes, bytearray)):
        doc = fitz.open(stream=caminho_ou_bytes, filetype="pdf")
    else:
        doc = fitz.open(str(caminho_ou_bytes))
    with doc:
        return "\n".join(p.get_text() for p in doc)


class _Homologacao(_Base):
    """Planilha de homologação (Alfa, Beta, Gama) + Épsilon e a visita V00005."""

    def setUp(self):
        super().setUp()
        import importlib.util
        spec = importlib.util.spec_from_file_location("oauth_setup", ROOT / "scripts" / "oauth_homologacao_setup.py")
        setup_mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(setup_mod)
        self.planilha = self.google.planilhas[self.sheet_teste]
        setup_mod.preencher_planilha(self.planilha)
        tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
        for node in tree.body:
            if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id in CONSTANTES_APP
                                                    for t in node.targets):
                exec(compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / "app.py"), "exec",
                             flags=__future__.annotations.compiler_flag), self.app)
        self.app["DADOS_JSON_NAME"] = "dados_condominio.json"
        self.app["CLIENTES_SEM_RT_JSON"] = Path(self._tmp.name) / "_clientes_sem_rt.json"
        self.gerados = Path(self._tmp.name)
        # Épsilon: cliente Bem Star com uma piscina cadastrada ("Piscina Adulto").
        self.app["sheets_salvar_cliente"](EPSILON, "", "Rua Fictícia E, 5", "Síndico Épsilon Fictício", "",
                                         vol_adulto=45, empresa="Bem Star Piscinas")
        self.foto = _foto()
        self.foto_id = self.app["drive_upload_foto"](self.foto, "antes_epsilon.jpg", EPSILON, "2026-10")
        self.assertTrue(self.foto_id)

    def _lancamento_operador(self) -> dict:
        """Lançamento como o modo campo monta antes de gravar (sem ID)."""
        return {
            "data": "27/10/2026", "operador": "OPERADOR BEM STAR FICTÍCIO", "condominio": EPSILON,
            "salvo_em": "27/10/2026 10:15:00", "observacao": "DADO FICTÍCIO — HOMOLOGAÇÃO",
            "ph": "7,4", "cloro_livre": "2,0", "cloro_total": "2,2", "alcalinidade": "100",
            "dureza": "220", "cianurico": "40",
            "piscinas": [{"nome": "Piscina Adulto", "ph": "7,4", "cloro_livre": "2,0", "cloro_total": "2,2",
                          "alcalinidade": "100", "dureza": "220", "cianurico": "40"}],
            "fotos_antes": ["antes_20261027_101500_1_epsilon.jpg"],
            "fotos_antes_ids": [self.foto_id],
            "fotos_antes_b64": [base64.b64encode(self.foto).decode()],
            "dosagens": [],
        }

    def _gravar_v00005(self, com_id_local: bool) -> tuple[dict, dict]:
        """Grava a visita no Sheets (gera V00005) e a cópia local do modo campo."""
        lanc = self._lancamento_operador()
        local = json.loads(json.dumps(lanc))  # cópia gravada antes do Sheets (sem ID)
        self.assertTrue(self.app["sheets_salvar_lancamento_campo"](lanc, EPSILON))
        if com_id_local:
            local["id_visita"] = lanc["id_visita"]
        pasta = self.gerados / slugify_nome(EPSILON)
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "dados_condominio.json").write_text(json.dumps(
            {"nome_condominio": EPSILON, "lancamentos_campo": [local]}, ensure_ascii=False), encoding="utf-8")
        return lanc, local


# ── 1. PDF individual da visita mantém a fotografia ─────────────────────────
class FotoNoPdfDaVisita(_Homologacao):
    def test_foto_base64_permanece_ate_o_build(self):
        lanc = self._lancamento_operador()
        pdf = self.app["gerar_pdf_relatorio_visita_bem_star"](lanc, EPSILON)
        self.assertIn((640, 480), _imagens_pdf(pdf), "a fotografia da visita sumiu do PDF")
        self.assertIn("REGISTRO FOTOGRAFICO", _texto_pdf(pdf))

    def test_nenhum_arquivo_temporario_e_criado_ou_apagado_antes_do_build(self):
        apagados = []
        import os as _os
        original = _os.unlink

        def unlink(caminho, *a, **k):
            apagados.append(str(caminho))
            return original(caminho, *a, **k)

        from unittest import mock
        with mock.patch("os.unlink", unlink):
            pdf = self.app["gerar_pdf_relatorio_visita_bem_star"](self._lancamento_operador(), EPSILON)
        self.assertEqual([c for c in apagados if c.endswith(".jpg")], [])
        self.assertIn((640, 480), _imagens_pdf(pdf))

    def test_visita_regerada_do_sheets_baixa_a_foto_do_drive(self):
        lanc, _ = self._gravar_v00005(com_id_local=True)
        visitas = self.app["situacao_visitas_condominio"](EPSILON)["visitas"]
        v5 = next(v for v in visitas if v["id_visita"] == "V00005")
        self.assertFalse(any(k.endswith("_b64") for k in v5), "o Sheets não guarda base64")
        shutil.rmtree(self.gerados / slugify_nome(EPSILON))  # outro servidor: sem cópia local
        pdf = self.app["gerar_pdf_relatorio_visita_bem_star"](v5, EPSILON)
        self.assertIn((640, 480), _imagens_pdf(pdf))
        # A foto da homologação continua no Drive (nada foi apagado).
        self.assertEqual(self.app["drive_baixar_foto"](self.foto_id), self.foto)

    def test_visita_sem_foto_declarada_nao_recebe_foto_de_outra_visita(self):
        self._gravar_v00005(com_id_local=True)
        pasta_campo = self.gerados / slugify_nome(EPSILON) / "fotos_campo"
        pasta_campo.mkdir(parents=True, exist_ok=True)
        (pasta_campo / "antes_20261020_090000_1_outra.jpg").write_bytes(_foto((0, 0, 255), (320, 240)))
        sem_foto = {"data": "28/10/2026", "operador": "OP", "condominio": EPSILON, "piscinas": []}
        pdf = self.app["gerar_pdf_relatorio_visita_bem_star"](sem_foto, EPSILON)
        self.assertNotIn((320, 240), _imagens_pdf(pdf))


# ── 2 e 3. PDF mensal: uma leitura por visita e piscinas do cadastro ───────
class RelatorioMensalEpsilon(_Homologacao):
    def _visitas_mes(self, com_id_local: bool):
        _, local = self._gravar_v00005(com_id_local)
        sheets = self.app["situacao_visitas_condominio"](EPSILON)["visitas"]
        self.assertEqual([v["id_visita"] for v in sheets], ["V00005"])
        return [local], sheets

    def test_sheets_devolve_id_para_a_copia_local(self):
        lanc, _ = self._gravar_v00005(com_id_local=False)
        self.assertEqual(lanc["id_visita"], "V00005")

    def test_copia_local_legada_sem_id_nao_duplica(self):
        local, sheets = self._visitas_mes(com_id_local=False)  # dado da homologação original
        unidas = bsm.deduplicar_lancamentos(local + sheets)
        self.assertEqual(len(unidas), 1)
        self.assertEqual(unidas[0]["id_visita"], "V00005")
        self.assertTrue(unidas[0].get("fotos_antes_b64"), "a foto da cópia local deve ser preservada")
        self.assertEqual(len(bsm.expandir(local + sheets)), 1)

    def test_copia_local_com_id_nao_duplica(self):
        local, sheets = self._visitas_mes(com_id_local=True)
        self.assertEqual(len(bsm.deduplicar_lancamentos(local + sheets)), 1)

    def test_visitas_distintas_nao_sao_unidas(self):
        a = {"id_visita": "V00005", "data": "27/10/2026", "salvo_em": "27/10/2026 10:15:00", "operador": "OP"}
        b = dict(a, id_visita="V00006")
        c = {"data": "27/10/2026", "operador": "OP", "salvo_em": "27/10/2026 16:00:00"}
        d = {"data": "27/10/2026", "operador": "OP", "ph": "7,2"}  # sem ID e sem horário
        e = dict(d, ph="7,6")
        self.assertEqual(len(bsm.deduplicar_lancamentos([a, b, c, d, e])), 5)

    def test_pdf_mensal_uma_leitura_e_sem_piscina_adulta(self):
        local, sheets = self._visitas_mes(com_id_local=False)
        cliente = next(c for c in self.app["sheets_listar_clientes_completo"]() if c["nome"] == EPSILON)
        cadastradas = self.app["_piscinas_cadastradas_cliente"](cliente)
        self.assertEqual(cadastradas, ["Piscina Adulto"])
        lancamentos = bsm.deduplicar_lancamentos(local + sheets)
        # Agenda preenchida com a sugestão antiga "Piscina adulta" no mesmo dia.
        agenda = [{"data": "27/10/2026", "piscina": "Piscina adulta", "ph": "7,4", "cloro_livre": "2,0",
                   "cloro_total": "2,2", "alcalinidade": "100", "dureza": "220", "cianurico": "40",
                   "observacao": "", "dosagens": [], "fonte": "Agenda mensal Bem Star"}]
        unidas = bsm.unir_agenda(lancamentos, agenda, cadastradas)
        self.assertEqual([(l["data"], l["piscina"]) for l in unidas], [("27/10/2026", "Piscina Adulto")])
        ctx = {"cliente": EPSILON, "mes": "10", "ano": "2026", "dados_cliente": cliente,
               "lancamentos": unidas, "piscinas": ["Piscina adulta"], "piscinas_cadastradas": cadastradas,
               "tipos_piscinas": {"Piscina Adulto": "Adulto"}, "fotos": [], "operador": "", "obs_geral": "",
               "parecer_tecnico": ""}
        saida = self.gerados / "mensal_epsilon.pdf"
        ok, erro = pdf_bs.gerar_pdf(saida, ctx, cnpj="00.000.000/0000-00", responsavel="RESPONSÁVEL FICTÍCIO",
                                    logo_path=ROOT / "bem_star_logo.png")
        self.assertTrue(ok, erro)
        texto = _texto_pdf(saida)
        self.assertIn("Piscina Adulto", texto)
        self.assertNotIn("Piscina adulta", texto)
        self.assertNotIn("piscina adulta", texto.casefold().replace("piscina adulto", ""))

    def test_piscinas_diferentes_continuam_separadas(self):
        chave = bsm.chave_piscina
        self.assertEqual(chave("Piscina adulta"), chave("Piscina  ADULTO"))
        self.assertEqual(chave("Piscina Infantis"), chave("piscina infantil"))
        for a, b in [("Piscina Adulto", "Piscina Adulto 2"), ("Piscina Adulto", "Piscina Infantil"),
                     ("Piscina Adulto", "Piscina aquecida"), ("Spa", "Piscina principal")]:
            self.assertNotEqual(chave(a), chave(b), (a, b))
        linhas = bsm.expandir([{"id_visita": "V1", "data": "27/10/2026", "piscinas": [
            {"nome": "Piscina Adulto", "ph": "7,4"}, {"nome": "Piscina Adulto 2", "ph": "7,5"},
            {"nome": "Piscina Infantil", "ph": "7,6"}]}])
        self.assertEqual([l["piscina"] for l in linhas], ["Piscina Adulto", "Piscina Adulto 2", "Piscina Infantil"])
        self.assertEqual(bsm.canonizar_nomes_piscinas(["Piscina adulta", "Piscina Infantil"], ["Piscina Adulto"]),
                         ["Piscina Adulto", "Piscina Infantil"])


# ── 4. Isolamento entre empresas ────────────────────────────────────────────
class IsolamentoEmpresas(_Homologacao):
    def _pasta_local_sem_empresa(self, nome):
        pasta = self.gerados / slugify_nome(nome)
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "dados_condominio.json").write_text(json.dumps(
            {"nome_condominio": nome, "lancamentos_campo": [{"data": "20/10/2026", "operador": "OP"}]},
            ensure_ascii=False), encoding="utf-8")
        return pasta

    def test_gama_com_pasta_local_sem_empresa_continua_bem_star(self):
        self._pasta_local_sem_empresa(GAMA)
        clientes = self.app["sheets_listar_clientes_completo"]()
        gama = next(c for c in clientes if c["nome"] == GAMA)
        self.assertEqual(gama["servicos"], {"rt": False, "limpeza": True})
        filtrar = self.app["filtrar_clientes_por_empresa"]
        self.assertNotIn(GAMA, [c["nome"] for c in filtrar(clientes, "aqua_gestao")])
        self.assertIn(GAMA, [c["nome"] for c in filtrar(clientes, "bem_star")])
        self.assertNotIn(GAMA, self.app["_listar_condominios_rt_aqua"](clientes))
        # Registro local preservado.
        dados = json.loads((self.gerados / slugify_nome(GAMA) / "dados_condominio.json").read_text(encoding="utf-8"))
        self.assertEqual(len(dados["lancamentos_campo"]), 1)

    def test_planilha_prevalece_sobre_json_local_divergente(self):
        pasta = self._pasta_local_sem_empresa(GAMA)
        dados = json.loads((pasta / "dados_condominio.json").read_text(encoding="utf-8"))
        dados["empresa"] = "Aqua Gestão"  # cópia local antiga
        (pasta / "dados_condominio.json").write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
        gama = next(c for c in self.app["sheets_listar_clientes_completo"]() if c["nome"] == GAMA)
        self.assertEqual(gama["empresa"], "Bem Star Piscinas")

    def test_json_local_explicito_vale_quando_planilha_nao_informa(self):
        resolver = self.app["_servicos_cliente_resolvidos"]
        self.assertEqual(resolver({"empresa": ""}, {"_empresa_definida_localmente": True, "empresa": "Ambas",
                                                    "servicos": {"rt": True, "limpeza": True}}),
                         {"rt": True, "limpeza": True})
        self.assertEqual(resolver({"empresa": ""}, {}), {"rt": True, "limpeza": False})  # legado inalterado

    def test_pastas_locais_filtradas_por_empresa(self):
        clientes = self.app["sheets_listar_clientes_completo"]()
        gama = self._pasta_local_sem_empresa(GAMA)
        alfa = self._pasta_local_sem_empresa("Condomínio Alfa Fictício")
        legado = self._pasta_local_sem_empresa("Condomínio Legado Sem Cadastro")
        filtrar = self.app["filtrar_pastas_por_empresa"]
        self.assertEqual(filtrar([gama, alfa, legado], "aqua_gestao", clientes), [alfa, legado])
        self.assertEqual(filtrar([gama, alfa, legado], "bem_star", clientes), [gama, legado])


# ── 6. Status da água e botão do relatório Bem Star ────────────────────────
class StatusAguaSeletor(unittest.TestCase):
    def setUp(self):
        tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
        func = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "sincronizar_status_agua_widget")

        class _St:
            session_state = {}
        self.st = _St()
        self.ns = {"st": self.st}
        exec(compile(ast.Module(body=[func], type_ignores=[]), "app.py", "exec",
                     flags=__future__.annotations.compiler_flag), self.ns)
        self.opcoes = ["CONFORME", "NÃO CONFORME", "EM CORREÇÃO"]

    def _fim_execucao(self):
        """O que o app faz depois do seletor."""
        ss = self.st.session_state
        ss["rel_status_agua"] = ss["rel_status_agua_select"]
        ss["_rel_status_agua_sinc"] = ss["rel_status_agua"]

    def test_status_calculado_aparece_no_seletor(self):
        ss = self.st.session_state
        self.ns["sincronizar_status_agua_widget"](self.opcoes)
        self._fim_execucao()
        self.assertEqual(ss["rel_status_agua_select"], "CONFORME")
        ss["rel_status_agua"] = "NÃO CONFORME"  # parecer automático / rascunho restaurado
        self.assertEqual(self.ns["sincronizar_status_agua_widget"](self.opcoes), "NÃO CONFORME")
        self.assertEqual(ss["rel_status_agua_select"], "NÃO CONFORME")

    def test_escolha_manual_no_seletor_prevalece(self):
        ss = self.st.session_state
        self.ns["sincronizar_status_agua_widget"](self.opcoes)
        self._fim_execucao()
        ss["rel_status_agua_select"] = "EM CORREÇÃO"  # usuário trocou no seletor
        self.assertEqual(self.ns["sincronizar_status_agua_widget"](self.opcoes), "EM CORREÇÃO")
        self.assertEqual(ss["rel_status_agua"], "EM CORREÇÃO")


class BotaoRelatorioBemStar(unittest.TestCase):
    def test_botao_so_existe_com_cliente_e_nao_usa_csr_sel(self):
        """O handler ficava fora do bloco que define csr_sel → NameError 'csr_sel'."""
        fonte = (ROOT / "app.py").read_text(encoding="utf-8")
        tree = ast.parse(fonte)
        alvo = None
        for node in ast.walk(tree):
            if isinstance(node, ast.If) and "Gerar relatório Bem Star (PDF)" in ast.get_source_segment(fonte, node.test):
                alvo = node
                break
        self.assertIsNotNone(alvo)
        self.assertIn("opcoes_csr", ast.get_source_segment(fonte, alvo.test))
        nomes = {n.id for n in ast.walk(alvo) if isinstance(n, ast.Name)}
        self.assertFalse(nomes & {"csr_sel", "csr_mes", "csr_ano", "csr_dados_sel"}, nomes)


if __name__ == "__main__":
    unittest.main()
