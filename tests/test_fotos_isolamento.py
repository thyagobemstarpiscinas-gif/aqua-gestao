"""Fotos do operador e do relatório Bem Star, deduplicação e isolamento por condomínio.

Defeitos cobertos (todos reproduzidos na main a42a57c antes da correção):
1. coletar_rascunho_operador: KeyError 'extras' com uma foto em "Outras fotos" no
   rascunho (o dicionário não tinha a categoria). Em consequência, o salvamento
   automático falhava e a troca de condomínio ficava bloqueada.
2. Salvamento final da visita lia "Outras fotos" pela chave fixa "op_fotos_extras";
   depois de trocar de condomínio a chave do upload muda e as fotos eram ignoradas.
3. deduplicar_fotos: um Path era comparado pelo caminho e um upload pelo conteúdo,
   então a mesma foto nas duas formas (ou em duas pastas) não era deduplicada.
4. Relatório Bem Star: o PDF recebia todas as imagens da pasta do condomínio
   (outros meses, assinaturas, fotos do RT Aqua e cópias de pré-visualização).
5. Visitas de "Residencial Aurora II" entravam no relatório de "Residencial Aurora"
   (comparação por trecho do nome). Também: nome abreviado de vários condomínios,
   lista de clientes indisponível e nomes idênticos em cadastros diferentes.

Carregam só as funções envolvidas do app.py, com session_state simulado, pastas
temporárias e planilha simulada. Nenhum acesso a Google Sheets, Drive ou rede.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import atexit
import base64
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile
import types
import unittest
from datetime import datetime, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from utils.datas import hoje_br  # noqa: E402
from utils.datas_visitas import normalizar_data_visita  # noqa: E402
from utils.formatacao import slugify_nome  # noqa: E402
from utils.normalizacao import nomes_condominio_equivalentes, normalizar_texto_busca  # noqa: E402

APP = RAIZ / "app.py"
SRC = APP.read_text(encoding="utf-8")
TREE = ast.parse(SRC)

# PNG 1x1 fictícios com conteúdos diferentes (nenhuma foto real)
PNG_A = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==")
PNG_B = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")


def _carregar(nomes, extras=None):
    """Executa as definições de nível superior pedidas (funções sem decorador e constantes)."""
    nos = []
    for n in TREE.body:
        if isinstance(n, ast.FunctionDef) and n.name in nomes:
            n.decorator_list = []  # remove @st.cache_data
            nos.append(n)
        elif isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id in nomes for t in n.targets):
            nos.append(n)
        elif isinstance(n, ast.ClassDef) and n.name in nomes:
            nos.append(n)
    pasta = Path(tempfile.mkdtemp(prefix="fotos_teste_"))
    atexit.register(shutil.rmtree, pasta, ignore_errors=True)
    ns = {"st": types.SimpleNamespace(session_state={}, cache_data=lambda *a, **k: (lambda f: f)),
          "re": re, "json": json, "os": os, "shutil": shutil, "Path": Path, "datetime": datetime,
          "timedelta": timedelta, "GENERATED_DIR": pasta, "slugify_nome": slugify_nome,
          "normalizar_data_visita": normalizar_data_visita, "normalizar_texto_busca": normalizar_texto_busca,
          "nomes_condominio_equivalentes": nomes_condominio_equivalentes, "_log_sheets_erro": lambda *a, **k: None,
          "hoje_br": hoje_br, "__file__": str(APP)}
    ns.update(extras or {})
    exec(compile(ast.Module(body=nos, type_ignores=[]), str(APP), "exec"), ns)
    ns["pasta"] = pasta
    return ns


class _Upload(io.BytesIO):
    """Imita o UploadedFile do Streamlit."""

    def __init__(self, nome, dados):
        super().__init__(dados)
        self.name = nome


# ---------------------------------------------------------------- 1. "Outras fotos"
OPERADOR = ["_SLUG_PISCINAS_OPERADOR", "_slug_piscina_operador", "coletar_rascunho_operador",
            "_RE_CAMPO_PISCINA_OP", "_RE_CAMPO_DOSAGEM_OP", "limpar_formulario_operador",
            "_chave_fotos_operador", "trocar_condominio_operador", "salvar_rascunho_operador"]


def _operador():
    ns = _carregar(OPERADOR, {"_agora_brasilia": lambda: "26/10/2026 10:00:00", "_rascunho_usa_sheets": lambda: False})
    ns["ss"] = ns["st"].session_state
    return ns


def _fotos_no_rascunho(ns, cond, nomes):
    pasta = ns["pasta"] / slugify_nome(cond) / "fotos_rascunho"
    pasta.mkdir(parents=True, exist_ok=True)
    for nome in nomes:
        (pasta / nome).write_bytes(PNG_A)
    return pasta


class TestOutrasFotos(unittest.TestCase):
    def test_coletar_rascunho_com_outras_fotos_nao_quebra_e_lista_todas_as_categorias(self):
        ns = _operador()
        _fotos_no_rascunho(ns, "Condomínio Alfa", ["rasc_antes_a.jpg", "rasc_depois_b.jpg", "rasc_cmaq_c.jpg",
                                                   "rasc_extras_ralo.jpg", "rasc_extras_borda.jpg"])
        dados = ns["coletar_rascunho_operador"]("Condomínio Alfa", ["Piscina Adulto"])
        self.assertEqual(dados["fotos_rascunho"], {
            "antes": ["rasc_antes_a.jpg"], "depois": ["rasc_depois_b.jpg"], "cmaq": ["rasc_cmaq_c.jpg"],
            "extras": ["rasc_extras_borda.jpg", "rasc_extras_ralo.jpg"],
        })
        json.dumps(dados)  # o rascunho continua serializável

    def test_salvamento_automatico_com_outras_fotos_grava_o_rascunho(self):
        ns = _operador()
        _fotos_no_rascunho(ns, "Condomínio Alfa", ["rasc_extras_ralo.jpg"])
        ns["ss"].update({"op_adulto_ph": "7.4", "op_obs_campo": "Ralo trincado"})
        dados = ns["coletar_rascunho_operador"]("Condomínio Alfa", ["Piscina Adulto"])
        self.assertTrue(ns["salvar_rascunho_operador"]("Condomínio Alfa", dados))
        gravado = json.loads((ns["pasta"] / slugify_nome("Condomínio Alfa") / "_rascunho_operador.json").read_text())
        self.assertEqual(gravado["fotos_rascunho"]["extras"], ["rasc_extras_ralo.jpg"])
        self.assertEqual(gravado["obs"], "Ralo trincado")

    def test_troca_de_condominio_com_outras_fotos_nao_fica_bloqueada_e_preserva_as_fotos(self):
        ns = _operador()
        pasta = _fotos_no_rascunho(ns, "Condomínio Alfa", ["rasc_extras_ralo.jpg"])
        ns["ss"].update({"op_adulto_ph": "7.4", "_op_piscinas_form": ["Piscina Adulto"]})
        self.assertTrue(ns["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertNotIn("_op_troca_bloqueada", ns["ss"])
        self.assertTrue((pasta / "rasc_extras_ralo.jpg").exists())  # foto continua no rascunho de Alfa
        gravado = json.loads((pasta.parent / "_rascunho_operador.json").read_text())
        self.assertEqual((gravado["fotos_rascunho"]["extras"], gravado["piscinas"][0]["ph"]), (["rasc_extras_ralo.jpg"], "7.4"))

    def test_salvamento_final_le_outras_fotos_pela_chave_do_upload(self):
        self.assertNotIn('st.session_state.get("op_fotos_extras")', SRC)
        self.assertIn('op_fotos_extras_final = st.session_state.get(_chave_fotos_operador("extras")) or []', SRC)


# ---------------------------------------------------------------- 2. deduplicação
class TestDeduplicacao(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(["deduplicar_fotos"])
        self.dedup = self.ns["deduplicar_fotos"]
        self.pasta = self.ns["pasta"]

    def test_mesma_foto_como_path_e_como_upload_aparece_uma_vez(self):
        p = self.pasta / "rasc_extras_image.jpg"
        p.write_bytes(PNG_A)
        up = _Upload("image.jpg", PNG_A)
        self.assertEqual(self.dedup([up, p]), [up])
        self.assertEqual(self.dedup([p, up]), [p])
        self.assertEqual(up.tell(), 0)  # o upload continua legível

    def test_mesma_foto_em_duas_pastas_aparece_uma_vez(self):
        (self.pasta / "a").mkdir()
        (self.pasta / "b").mkdir()
        p1, p2 = self.pasta / "a" / "x.jpg", self.pasta / "b" / "y.jpg"
        p1.write_bytes(PNG_A)
        p2.write_bytes(PNG_A)
        self.assertEqual(self.dedup([p1, p2, str(p2)]), [p1])

    def test_fotos_diferentes_com_o_mesmo_nome_sao_mantidas(self):
        (self.pasta / "a").mkdir()
        (self.pasta / "b").mkdir()
        p1, p2 = self.pasta / "a" / "image.jpg", self.pasta / "b" / "image.jpg"
        p1.write_bytes(PNG_A)
        p2.write_bytes(PNG_B)
        up = _Upload("image.jpg", PNG_B)
        self.assertEqual(self.dedup([p1, p2, up]), [p1, p2])

    def test_caminho_inexistente_continua_comparado_pelo_texto(self):
        self.assertEqual(self.dedup(["nao/existe.jpg", "nao/existe.jpg", "outro.jpg"]), ["nao/existe.jpg", "outro.jpg"])


# ---------------------------------------------------------------- 3. relatório Bem Star
def _pasta_condominio_mista(raiz, cond):
    """Pasta local de um condomínio com fotos de vários períodos e origens."""
    base = raiz / slugify_nome(cond)
    arquivos = {
        "out_a": base / "fotos_campo" / "antes_20261006_101010_1_a.jpg",
        "out_b": base / "fotos_campo" / "extras_20261006_101010_1_ralo.jpg",
        "set": base / "fotos_campo" / "antes_20260915_090000_1_setembro.jpg",
        "assinatura": base / "assinaturas_visita" / "assinatura_responsavel_20261006_101010.png",
        "rt_aqua": base / "fotos_relatorio_rt_manual" / "rt_casa_maquinas.jpg",
        "copia_previa": base / "_previa_exata_relatorio" / "fotos_upload" / "antes_20261006_101010_1_a.jpg",
        "copia_rascunho": base / "fotos_rascunho" / "rasc_antes_a.jpg",
    }
    for i, (chave, p) in enumerate(arquivos.items()):
        p.parent.mkdir(parents=True, exist_ok=True)
        # out_a e as duas cópias têm o mesmo conteúdo; as demais são imagens distintas
        p.write_bytes(_png(0) if chave in ("out_a", "copia_previa", "copia_rascunho") else _png(i * 30))
    return arquivos


def _png(cor):
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), (cor % 256, 80, 160)).save(buf, format="PNG")
    return buf.getvalue()


class TestFotosRelatorioBemStar(unittest.TestCase):
    def setUp(self):
        self.ns = _carregar(["deduplicar_fotos", "_coletar_fotos_bem_star_preview"])
        self.arq = _pasta_condominio_mista(self.ns["pasta"], "Condomínio Alfa")
        self.lanc_out = [{"data": "06/10/2026", "fotos": ["antes_20261006_101010_1_a.jpg",
                                                          "extras_20261006_101010_1_ralo.jpg"]}]

    def test_somente_fotos_das_visitas_do_periodo_sem_duplicar(self):
        fotos = self.ns["_coletar_fotos_bem_star_preview"]("Condomínio Alfa", self.lanc_out)
        self.assertEqual(fotos, [self.arq["out_a"], self.arq["out_b"]])

    def test_sem_nomes_nos_lancamentos_usa_fotos_de_campo_do_mes_dos_lancamentos(self):
        fotos = self.ns["_coletar_fotos_bem_star_preview"]("Condomínio Alfa", [{"data": "06/10/2026"}])
        self.assertEqual(sorted(fotos), sorted([self.arq["out_a"], self.arq["out_b"]]))
        for chave in ("set", "assinatura", "rt_aqua", "copia_previa", "copia_rascunho"):
            self.assertNotIn(self.arq[chave], fotos, chave)

    def test_pdf_bem_star_tem_cada_foto_uma_unica_vez(self):
        from reportlab import rl_config
        nomes = ["CNPJ_BEM_STAR", "RESPONSAVEL_TÉCNICO", "_bs_eh_segunda",
                 "gerar_pdf_relatorio_mensal_bem_star_modelo", "deduplicar_fotos", "_coletar_fotos_bem_star_preview"]
        ns = _carregar(nomes)
        arq = _pasta_condominio_mista(ns["pasta"], "Condomínio Alfa")
        fotos = ns["_coletar_fotos_bem_star_preview"]("Condomínio Alfa", self.lanc_out)
        ctx = {"cliente": "Condomínio Alfa", "mes": "10", "ano": "2026", "operador": "Operador Fictício",
               "dados_cliente": {"nome": "Condomínio Alfa"}, "lancamentos": [
                   {"data": "06/10/2026", "ph": "7,4", "cloro_livre": "1,5", "operador": "Operador Fictício"}],
               "fotos": fotos}
        saida = ns["pasta"] / "bem_star.pdf"
        compressao = rl_config.pageCompression
        rl_config.pageCompression = 0
        try:
            ok, erro = ns["gerar_pdf_relatorio_mensal_bem_star_modelo"](saida, ctx)
        finally:
            rl_config.pageCompression = compressao
        self.assertTrue(ok, erro)
        pdf = saida.read_bytes().decode("latin-1")
        legendas = [arq["out_a"].stem.replace("_", " ").title(), arq["out_b"].stem.replace("_", " ").title()]
        for legenda in legendas:
            self.assertEqual(pdf.count(f"({legenda})"), 1, legenda)
        # (a cópia da pré-visualização tem o mesmo nome de out_a: coberta pela contagem acima)
        for chave in ("set", "assinatura", "rt_aqua", "copia_rascunho"):
            self.assertNotIn(f"({arq[chave].stem.replace('_', ' ').title()})", pdf, chave)


# ---------------------------------------------------------------- 4. identificação do condomínio
class _Aba:
    def __init__(self, linhas, falha=False):
        self.linhas, self.falha = linhas, falha

    def get_all_values(self):
        if self.falha:
            raise ConnectionError("planilha indisponível (simulado)")
        return self.linhas


def _cliente(id_, nome, empresa):
    """Linha da aba 👥 Clientes no layout real: B=ID, C=Nome, I=Status, M=Empresa."""
    linha = [""] * 15
    linha[1], linha[2], linha[8], linha[12] = id_, nome, "Ativo", empresa
    return linha


def _visita(id_, data, cond, id_cliente=""):
    return ["", id_, data, id_cliente, cond, "7,4"] + [""] * 17


CABECALHO_CLIENTES = ["", "ID", "Nome"] + [""] * 12
IDENTIFICACAO = ["sheets_listar_clientes", "sheets_listar_lancamentos", "_visita_pertence_ao_condominio",
                 "chave_validacao_origem", "origem_validada",
                 "_ClientesIndisponiveis", "_VisitasIndisponiveis", "_ler_cadastro_identificacao", "_resolver_condominio_da_visita",
                 "_situacao_visitas_cache", "situacao_visitas_condominio"]


class TestIdentificacaoDoCondominioNasVisitas(unittest.TestCase):
    """Funções reais (lista de clientes e de visitas) sobre abas no layout da planilha."""

    def _app(self, clientes, visitas, clientes_falham=False, visitas_falham=False, conexoes_ok=None):
        """conexoes_ok: quantas chamadas a conectar_sheets() funcionam (None = todas)."""
        abas = {"👥 Clientes": _Aba([CABECALHO_CLIENTES] + clientes, falha=clientes_falham),
                "🔬 Visitas": _Aba([["", "ID Visita", "Data", "ID Cliente", "Condomínio"]] + visitas, falha=visitas_falham)}
        chamadas = []

        def conectar_sheets():
            chamadas.append(1)
            return object() if conexoes_ok is None or len(chamadas) <= conexoes_ok else None
        ns = _carregar(IDENTIFICACAO, {"conectar_sheets": conectar_sheets, "obter_aba_sheets": lambda nome: abas[nome]})
        listar = lambda nome: [v["id_visita"] for v in ns["sheets_listar_lancamentos"](nome)]  # noqa: E731
        aviso = lambda nome: ns["situacao_visitas_condominio"](nome)["aviso"]  # noqa: E731
        return ns, listar, aviso

    def test_nome_exato_de_outro_condominio_nao_mistura(self):
        _, listar, _ = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("C2", "Residencial Aurora II", "Aqua Gestão")],
            [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V2", "06/10/2026", "Residencial Aurora II"),
             _visita("V3", "13/10/2026", "residencial  AURORA")])
        self.assertEqual(listar("Residencial Aurora"), ["V1", "V3"])
        self.assertEqual(listar("Residencial Aurora II"), ["V2"])

    def test_nome_abreviado_de_um_unico_condominio_continua_aceito(self):
        _, listar, aviso = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("C5", "Tríad Vertical", "Aqua Gestão")],
            [_visita("V1", "06/10/2026", "Aurora"), _visita("V2", "06/10/2026", "Condomínio Tríad")])
        self.assertEqual(listar("Residencial Aurora"), ["V1"])
        self.assertEqual(aviso("Residencial Aurora"), "")
        self.assertEqual(listar("Tríad Vertical"), [])  # nomes sem trecho em comum

    def test_defeito_1_nome_abreviado_de_varios_condominios_nao_e_importado(self):
        _, listar, aviso = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("C2", "Residencial Aurora II", "Aqua Gestão")],
            [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V9", "06/10/2026", "Aurora"),
             _visita("V5", "13/10/2026", "Residencial Aurora II")])
        self.assertEqual(listar("Residencial Aurora"), ["V1"])  # V9 ficaria nos dois
        self.assertEqual(listar("Residencial Aurora II"), ["V5"])
        self.assertIn("1 visita", aviso("Residencial Aurora"))
        self.assertIn("mais de um condomínio", aviso("Residencial Aurora"))

    def test_defeito_2_lista_de_clientes_indisponivel_interrompe_com_aviso(self):
        ns, listar, aviso = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("C2", "Residencial Aurora II", "Aqua Gestão")],
            [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V5", "13/10/2026", "Residencial Aurora II")],
            clientes_falham=True)
        self.assertEqual(listar("Residencial Aurora"), [])  # antes: V1 e V5 (de Aurora II)
        situacao = ns["situacao_visitas_condominio"]("Residencial Aurora")
        self.assertTrue(situacao["bloqueio"])
        self.assertIn("cadastro de clientes", aviso("Residencial Aurora"))

    def test_defeito_3_nomes_identicos_em_cadastros_diferentes_nao_misturam(self):
        ns, listar, aviso = self._app(
            [_cliente("C3", "Residencial Sol", "Aqua Gestão"), _cliente("C7", "residencial  SOL", "Bem Star Piscinas"),
             _cliente("C8", "Edifício Lua", "Bem Star Piscinas")],
            [_visita("V6", "08/10/2026", "Residencial Sol", "C3"), _visita("V7", "09/10/2026", "Residencial Sol", "C3"),
             _visita("V8", "09/10/2026", "Edifício Lua", "C8")])
        self.assertEqual(listar("Residencial Sol"), [])  # o ID gravado (C3 nas duas) não é confiável
        self.assertTrue(ns["situacao_visitas_condominio"]("Residencial Sol")["bloqueio"])
        texto = aviso("Residencial Sol")
        for parte in ("C3", "C7", "Aqua Gestão", "Bem Star Piscinas"):
            self.assertIn(parte, texto)
        self.assertEqual(listar("Edifício Lua"), ["V8"])  # outros condomínios seguem normais
        self.assertEqual(aviso("Edifício Lua"), "")

    # ---- três situações: consulta com visitas, consulta sem visitas, falha/bloqueio
    CLIENTES_AURORA = [_cliente("C1", "Residencial Aurora", "Aqua Gestão"),
                       _cliente("C2", "Residencial Aurora II", "Aqua Gestão"),
                       _cliente("C3", "Residencial Sol", "Aqua Gestão"), _cliente("C7", "Residencial Sol", "Bem Star Piscinas")]

    def test_consulta_com_visitas_e_consulta_sem_visitas_tem_status_ok(self):
        ns, _, _ = self._app(self.CLIENTES_AURORA, [_visita("V1", "06/10/2026", "Residencial Aurora")])
        com = ns["situacao_visitas_condominio"]("Residencial Aurora")
        sem = ns["situacao_visitas_condominio"]("Residencial Aurora II")
        self.assertEqual((com["status"], [v["id_visita"] for v in com["visitas"]], com["aviso"]), ("ok", ["V1"], ""))
        self.assertEqual((sem["status"], sem["visitas"], sem["aviso"], sem["bloqueio"]), ("ok", [], "", ""))

    def _falha_de_visitas(self, **kwargs):
        ns, listar, _ = self._app(self.CLIENTES_AURORA, [_visita("V1", "06/10/2026", "Residencial Aurora")], **kwargs)
        situacao = ns["situacao_visitas_condominio"]("Residencial Aurora")
        self.assertEqual((situacao["status"], situacao["bloqueio"], situacao["visitas"]), ("falha", "visitas_indisponiveis", []))
        self.assertIn("🔬 Visitas", situacao["aviso"])
        self.assertIn("não ausência de visitas", situacao["aviso"])
        # exceção (de visitas ou, se a conexão já caiu de vez, de cadastro): o st.cache_data não guarda a falha
        with self.assertRaises((ns["_VisitasIndisponiveis"], ns["_ClientesIndisponiveis"])):
            ns["_situacao_visitas_cache"]("Residencial Aurora")
        return ns

    def test_falha_de_leitura_da_aba_visitas_e_falha_e_nao_lista_vazia(self):
        self._falha_de_visitas(visitas_falham=True)

    def test_falha_de_conexao_na_leitura_das_visitas_e_falha(self):
        self._falha_de_visitas(conexoes_ok=1)  # o cadastro foi lido; a conexão cai em seguida

    def test_aba_visitas_inexistente_e_falha(self):
        abas = {"👥 Clientes": _Aba([CABECALHO_CLIENTES] + self.CLIENTES_AURORA)}
        ns = _carregar(IDENTIFICACAO, {"conectar_sheets": lambda: object(), "obter_aba_sheets": lambda nome: abas[nome]})
        self.assertEqual(ns["situacao_visitas_condominio"]("Residencial Aurora")["status"], "falha")

    def test_falha_do_cadastro_e_falha_e_duplicado_e_bloqueio(self):
        ns, _, _ = self._app(self.CLIENTES_AURORA, [], clientes_falham=True)
        self.assertEqual(ns["situacao_visitas_condominio"]("Residencial Aurora")["status"], "falha")
        ns, _, _ = self._app(self.CLIENTES_AURORA, [])
        self.assertEqual(ns["situacao_visitas_condominio"]("Residencial Sol")["status"], "bloqueio")

    # ---- Defeito B: nome digitado que não é cadastro, mas pode representar um
    def test_nome_digitado_semelhante_a_cadastro_e_bloqueado(self):
        visitas = [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V2", "06/10/2026", "Aurora"),
                   _visita("V8", "09/10/2026", "Sol"), _visita("V11", "06/10/2026", "Torre A"),
                   _visita("V12", "06/10/2026", "Residencial Auror")]
        ns, listar, aviso = self._app(self.CLIENTES_AURORA, visitas)
        for nome, cadastros in (("Sol", ("C3", "C7")), ("Aurora", ("C1", "C2")), ("Residencial Auror", ("C1", "C2")),
                                ("Residencial Aurora Torre A", ("C1",))):
            situacao = ns["situacao_visitas_condominio"](nome)
            self.assertEqual((situacao["status"], situacao["bloqueio"], situacao["visitas"]),
                             ("bloqueio", "nome_semelhante", []), nome)
            for c in cadastros:
                self.assertIn(c, situacao["aviso"], nome)

    def test_relatorio_avulso_sem_correspondencia_continua_e_nao_recebe_visitas_de_clientes(self):
        visitas = [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V20", "06/10/2026", "Chácara Avulsa"),
                   _visita("V21", "07/10/2026", "chacara  avulsa")]
        ns, listar, aviso = self._app(self.CLIENTES_AURORA, visitas)
        situacao = ns["situacao_visitas_condominio"]("Chácara Avulsa")
        self.assertEqual((situacao["status"], [v["id_visita"] for v in situacao["visitas"]], situacao["aviso"]),
                         ("ok", ["V20", "V21"], ""))
        self.assertEqual(listar("Condomínio Inexistente"), [])

    def test_lista_com_visitas_ambiguas_e_incompleta_e_nao_ok(self):
        ns, listar, _ = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("C2", "Residencial Aurora II", "Aqua Gestão")],
            [_visita("V1", "06/10/2026", "Residencial Aurora"), _visita("V9", "13/10/2026", "Aurora")])
        situacao = ns["situacao_visitas_condominio"]("Residencial Aurora")
        self.assertEqual((situacao["status"], situacao["bloqueio"], [v["id_visita"] for v in situacao["visitas"]]),
                         ("ambiguo", "", ["V1"]))  # parcial: V1 disponível para importação manual
        self.assertEqual(situacao["ambiguas"], ["V9 (Aurora)"])

    def test_validacao_explicita_por_condominio_e_periodo(self):
        ns, _, _ = self._app([], [])
        chave = ns["chave_validacao_origem"]("rt", "Condomínio Alfa", "9", "2026")
        self.assertEqual(chave, ns["chave_validacao_origem"]("rt", "Condomínio Alfa", "09", "2026"))
        self.assertNotEqual(chave, ns["chave_validacao_origem"]("rt", "Condomínio Alfa", "10", "2026"))
        self.assertNotEqual(chave, ns["chave_validacao_origem"]("bem_star", "Condomínio Alfa", "09", "2026"))
        self.assertTrue(chave.startswith("_"))  # fora do rascunho RT (só chaves rel_/csr_)
        validada = ns["origem_validada"]
        ss = ns["st"].session_state
        self.assertTrue(validada({"status": "ok"}, chave))
        for status in ("ambiguo", "bloqueio", "falha"):
            self.assertFalse(validada({"status": status}, chave), status)
        ss[chave] = True
        self.assertTrue(validada({"status": "ambiguo"}, chave))
        self.assertTrue(validada({"status": "bloqueio"}, chave))
        self.assertFalse(validada({"status": "falha"}, chave))  # falha nunca é validada manualmente

    def test_cadastro_sem_id_tambem_conta_para_ambiguidade(self):
        _, listar, _ = self._app(
            [_cliente("C1", "Residencial Aurora", "Aqua Gestão"), _cliente("", "Residencial Aurora II", "Aqua Gestão")],
            [_visita("V9", "06/10/2026", "Aurora")])
        self.assertEqual(listar("Residencial Aurora"), [])


class TestAvisosDeIdentificacaoNasTelas(unittest.TestCase):
    def test_todas_as_telas_mostram_o_aviso_e_respeitam_o_bloqueio(self):
        self.assertEqual(len(re.findall(r"\bsheets_listar_lancamentos\(", SRC)), 2)  # definição + docstring
        self.assertEqual(SRC.count("situacao_visitas_condominio("), 9)  # definição, docstring, wrapper, 5 telas e a confirmação do RT
        for trecho in ('st.warning(_situacao_csr_tela["aviso"])', 'st.warning(_situacao_rel["aviso"])',
                       "st.warning(f\"{nome_consulta}: {_situacao_rt['aviso']}\")",
                       'st.caption(_situacao_hist["aviso"])', "msg = f\"{_situacao_csr['aviso']} {msg}\""):
            self.assertIn(trecho, SRC)
        self.assertEqual(SRC.count("lancamentos_local = []  # a pasta local é identificada só pelo nome"), 1)
        self.assertIn('lancamentos_local = []  # pasta local identificada só pelo nome: não confirma o cadastro', SRC)
        self.assertIn("_lanc_local_csr = []  # pasta local identificada só pelo nome", SRC)


class TestIsolamentoDeCondominiosComNomesSemelhantes(unittest.TestCase):
    def test_rascunhos_e_fotos_ficam_em_pastas_separadas(self):
        ns = _carregar(OPERADOR + ["carregar_rascunho_operador", "deduplicar_fotos", "_coletar_fotos_bem_star_preview"],
                       {"_agora_brasilia": lambda: "26/10/2026 10:00:00", "_rascunho_usa_sheets": lambda: False})
        a, b = "Residencial Aurora", "Residencial Aurora II"
        self.assertNotEqual(slugify_nome(a), slugify_nome(b))
        self.assertTrue(ns["salvar_rascunho_operador"](a, {"obs": "rascunho de Aurora"}))
        self.assertEqual(ns["carregar_rascunho_operador"](b), {})  # Aurora II não vê o rascunho de Aurora
        self.assertEqual(ns["carregar_rascunho_operador"](a)["obs"], "rascunho de Aurora")
        for cond, cor in ((a, 10), (b, 200)):
            p = ns["pasta"] / slugify_nome(cond) / "fotos_campo" / "antes_20261006_101010_1_foto.jpg"
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_bytes(_png(cor))
        lanc = [{"data": "06/10/2026", "fotos": ["antes_20261006_101010_1_foto.jpg"]}]
        fotos_a = ns["_coletar_fotos_bem_star_preview"](a, lanc)
        self.assertEqual(len(fotos_a), 1)
        self.assertEqual(fotos_a[0].parent.parent.name, slugify_nome(a))


if __name__ == "__main__":
    unittest.main()
