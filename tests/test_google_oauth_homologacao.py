"""OAuth 2.0 de homologação (Gmail pessoal) com escopo mínimo drive.file.

Sem rede e sem credenciais reais: a renovação do token é simulada e o Google
(Drive + Sheets) é um falso em memória que aplica a regra do drive.file — o app
só acessa arquivos que ele próprio criou. Cobre segurança, gravação/leitura na
planilha de teste, fotografias no Drive e recuperação de rascunhos.
"""
import ast
import __future__
import builtins
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils import google_auth as ga  # noqa: E402
from utils.ambiente import destinos_google  # noqa: E402
from utils.formatacao import slugify_nome  # noqa: E402

DRIVE_FILE = ga.ESCOPO_OAUTH_HOMOLOGACAO
PROD_SHEET, PROD_PASTA = "PRODUCAO_PLANILHA_0001", "PRODUCAO_PASTA_0001"
SEGREDO = "SEGREDO-FICTICIO-NAO-PODE-VAZAR"
OAUTH = {"client_id": "cliente-teste.apps.googleusercontent.com", "client_secret": SEGREDO,
         "refresh_token": SEGREDO + "-refresh", "scopes": [DRIVE_FILE]}


# ── Google falso: drive.file só enxerga o que o app criou ────────────────────
class Negado(Exception):
    """Equivalente ao 404/403 do Google para arquivo fora do alcance do token."""


class GoogleFalso:
    def __init__(self):
        self.arquivos = {}   # id -> dict(nome, pais, mime, dados, dono)
        self.planilhas = {}  # id -> PlanilhaFalsa
        self.seq = 0
        # Recursos que existem na conta, mas NÃO foram criados pelo app (produção).
        self._novo("pasta", "ACQUA PRODUCAO FOTOS", [], dono="usuario", id_=PROD_PASTA)
        self.planilhas[PROD_SHEET] = PlanilhaFalsa("PRODUCAO", dono="usuario")

    def _novo(self, mime, nome, pais, dono="app", dados=b"", id_=None):
        self.seq += 1
        id_ = id_ or f"ID{self.seq:04d}_teste"
        self.arquivos[id_] = {"nome": nome, "pais": list(pais), "mime": mime, "dados": dados, "dono": dono}
        return id_

    def acessivel(self, id_):
        return id_ in self.arquivos and self.arquivos[id_]["dono"] == "app"

    def criar_recursos_teste(self):
        pasta = self._novo("application/vnd.google-apps.folder", "ACQUA HOMOLOGACAO", [])
        planilha = f"PLANILHA{self.seq:04d}_teste"
        self.planilhas[planilha] = PlanilhaFalsa("ACQUA HOMOLOGACAO", dono="app")
        return planilha, pasta


class _Exec:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class DriveFalso:
    def __init__(self, google, creds):
        self.g, self.creds = google, creds

    def _exigir_escopo(self):
        if list(getattr(self.creds, "scopes", []) or []) != [DRIVE_FILE]:
            raise AssertionError(f"Drive usado com escopo inesperado: {self.creds.scopes}")

    def files(self):
        return self

    def list(self, q, fields=None, **_):
        def run():
            self._exigir_escopo()
            nome = re.search(r"name='((?:\\.|[^'\\])*)'", q).group(1)
            nome = re.sub(r"\\(.)", r"\1", nome)
            pai = re.search(r"'((?:\\.|[^'\\])*)' in parents", q).group(1)
            achados = [{"id": i, "name": a["nome"]} for i, a in self.g.arquivos.items()
                       if self.g.acessivel(i) and a["nome"] == nome and pai in a["pais"]]
            return {"files": achados}
        return _Exec(run)

    def create(self, body, media_body=None, fields=None, **_):
        def run():
            self._exigir_escopo()
            for pai in body.get("parents", []):
                if not self.g.acessivel(pai):
                    raise Negado(f"pasta-pai fora do alcance do token: {pai}")
            dados = b""
            if media_body is not None:
                dados = media_body.getbytes(0, media_body.size())
            id_ = self.g._novo(body.get("mimeType", "arquivo"), body["name"], body.get("parents", []), dados=dados)
            return {"id": id_, "webViewLink": f"https://drive.invalid/{id_}"}
        return _Exec(run)

    def get_media(self, fileId):
        def run():
            self._exigir_escopo()
            if not self.g.acessivel(fileId):
                raise Negado(f"arquivo fora do alcance do token: {fileId}")
            return self.g.arquivos[fileId]["dados"]
        return _Exec(run)


class WorksheetNotFound(Exception):
    pass


class AbaFalsa:
    def __init__(self, titulo):
        self.title, self.linhas = titulo, []

    def get_all_values(self):
        return [list(l) for l in self.linhas]

    def update(self, range_name=None, values=None, value_input_option=None, **_):
        lin = int(re.search(r"\d+", range_name.split(":")[0]).group(0))
        while len(self.linhas) < lin:
            self.linhas.append([])
        self.linhas[lin - 1] = list(values[0])


class PlanilhaFalsa:
    def __init__(self, titulo, dono):
        self.title, self.dono, self.abas = titulo, dono, {}

    def worksheet(self, nome):
        if nome not in self.abas:
            raise WorksheetNotFound(nome)
        return self.abas[nome]

    def add_worksheet(self, title, rows, cols):
        self.abas[title] = AbaFalsa(title)
        return self.abas[title]


def gspread_falso(google):
    mod = types.ModuleType("gspread")
    mod.WorksheetNotFound = WorksheetNotFound

    class Cliente:
        def __init__(self, creds):
            self.creds = creds

        def open_by_key(self, key):
            if list(self.creds.scopes) != [DRIVE_FILE]:
                raise AssertionError("Sheets usado com escopo inesperado")
            planilha = google.planilhas.get(key)
            if planilha is None or planilha.dono != "app":
                raise Negado(f"planilha fora do alcance do token: {key}")
            return planilha

    mod.authorize = lambda creds: Cliente(creds)
    return mod


def _renovar_sem_rede(escopos_concedidos):
    def refresh(self, request):
        self.token = "token-acesso-ficticio"
        self._granted_scopes = list(escopos_concedidos)
    return refresh


# ── App real, somente as funções necessárias (sem executar o Streamlit) ─────
def _carregar_app(secrets, gerados: Path, sheets_id, pasta_id):
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    defs = {n.name: n for n in tree.body if isinstance(n, ast.FunctionDef)}

    class _Cache:
        def __call__(self, *a, **k):
            if a and callable(a[0]) and not k:
                return a[0]
            return lambda f: f

        def clear(self):
            pass

    class _St:
        cache_resource = _Cache()
        cache_data = _Cache()

        def __init__(self):
            self.secrets = secrets
            self.session_state = {}

        def __getattr__(self, _nome):
            return lambda *a, **k: None

    class _NS(dict):
        def __missing__(self, nome):
            if nome in defs:
                exec(compile(ast.Module(body=[defs[nome]], type_ignores=[]), str(ROOT / "app.py"), "exec",
                             flags=__future__.annotations.compiler_flag), self)
                return self[nome]
            if hasattr(builtins, nome):
                return getattr(builtins, nome)
            raise KeyError(nome)

    return _NS(st=_St(), AMBIENTE_APP="homologacao", SHEETS_ID=sheets_id, DRIVE_FOTOS_FOLDER_ID=pasta_id,
               GENERATED_DIR=gerados, BASE_DIR=ROOT, Path=Path, json=json, os=os, re=re, shutil=shutil,
               datetime=datetime, date=date, slugify_nome=slugify_nome, __file__=str(ROOT / "app.py"))


class _Base(unittest.TestCase):
    def setUp(self):
        self.google = GoogleFalso()
        self.sheet_teste, self.pasta_teste = self.google.criar_recursos_teste()
        self.secrets = {"aqua_app": {"ambiente": "homologacao", "sheets_id": self.sheet_teste,
                                     "drive_fotos_folder_id": self.pasta_teste},
                        "google_oauth": dict(OAUTH)}
        self._tmp = tempfile.TemporaryDirectory()
        self.patches = [
            mock.patch("google.oauth2.credentials.Credentials.refresh", _renovar_sem_rede([DRIVE_FILE])),
            mock.patch.dict(sys.modules, {"gspread": gspread_falso(self.google)}),
            mock.patch("googleapiclient.discovery.build",
                       lambda servico, versao, credentials=None, **_: DriveFalso(self.google, credentials)),
            mock.patch.dict(os.environ, {"CODESPACES": ""}),
        ]
        for p in self.patches:
            p.start()
        self.app = _carregar_app(self.secrets, Path(self._tmp.name), self.sheet_teste, self.pasta_teste)

    def tearDown(self):
        for p in reversed(self.patches):
            p.stop()
        self._tmp.cleanup()


# ── 1. Segurança ─────────────────────────────────────────────────────────────
class SegurancaOAuth(unittest.TestCase):
    def test_producao_continua_com_conta_de_servico_e_mesmos_escopos(self):
        capturado = {}

        def info(dados, scopes=None):
            capturado.update(dados=dados, scopes=scopes)
            return "credencial-conta-servico"

        with mock.patch("google.oauth2.service_account.Credentials.from_service_account_info", info):
            creds = ga.credenciais_google({"gcp_service_account": {"client_email": "sa@teste.invalid"}}, "producao")
        self.assertEqual(creds, "credencial-conta-servico")
        self.assertEqual(set(capturado["scopes"]), {"https://www.googleapis.com/auth/spreadsheets",
                                                    "https://www.googleapis.com/auth/drive"})
        self.assertEqual(ga.tipo_credencial({}, "producao"), "conta_servico")

    def test_oauth_e_recusado_em_producao(self):
        with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
            ga.tipo_credencial({"google_oauth": dict(OAUTH)}, "producao")

    def test_escopo_amplo_e_recusado_na_configuracao_e_no_token(self):
        for amplo in ("https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"):
            with self.subTest(escopo=amplo):
                with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
                    ga.tipo_credencial({"google_oauth": dict(OAUTH, scopes=[DRIVE_FILE, amplo])}, "homologacao")
                with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
                    ga.validar_escopos_concedidos([DRIVE_FILE, amplo])
                # Mesmo que o consentimento tenha sido dado com escopo amplo, o app recusa o token.
                with mock.patch("google.oauth2.credentials.Credentials.refresh", _renovar_sem_rede([DRIVE_FILE, amplo])):
                    with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
                        ga.credenciais_google({"google_oauth": dict(OAUTH)}, "homologacao")
        ga.validar_escopos_concedidos([DRIVE_FILE])

    def test_token_pede_somente_drive_file(self):
        with mock.patch("google.oauth2.credentials.Credentials.refresh", _renovar_sem_rede([DRIVE_FILE])):
            creds = ga.credenciais_google({"google_oauth": dict(OAUTH)}, "homologacao")
        self.assertEqual(list(creds.scopes), [DRIVE_FILE])

    def test_erros_nao_expoem_segredos(self):
        with self.assertRaises(ga.ConfiguracaoGoogleInvalida) as ctx:
            ga.tipo_credencial({"google_oauth": {"client_id": "x", "client_secret": SEGREDO}}, "homologacao")
        self.assertNotIn(SEGREDO, str(ctx.exception))
        self.assertIn("refresh_token", str(ctx.exception))

        def falha(self, request):
            raise RuntimeError(f"invalid_grant para {SEGREDO}")

        with mock.patch("google.oauth2.credentials.Credentials.refresh", falha):
            with self.assertRaises(ga.ConfiguracaoGoogleInvalida) as ctx:
                ga.credenciais_google({"google_oauth": dict(OAUTH)}, "homologacao")
        self.assertNotIn(SEGREDO, str(ctx.exception))
        self.assertIsNone(ctx.exception.__cause__)

    def test_homologacao_nunca_usa_arquivo_local_de_producao(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "aqua-gestao-rt-87316ebf5331.json").write_text("{}")
            with mock.patch("google.oauth2.service_account.Credentials.from_service_account_file",
                            side_effect=AssertionError("arquivo de produção usado")):
                self.assertIsNone(ga.credenciais_google({}, "homologacao", Path(d)))

    def test_ids_de_producao_continuam_bloqueados_em_homologacao(self):
        with self.assertRaises(ValueError):
            destinos_google({"aqua_app": {"ambiente": "homologacao", "sheets_id": PROD_SHEET,
                                          "drive_fotos_folder_id": "PASTA_TESTE_0001"}}, PROD_SHEET, PROD_PASTA)

    def test_app_usa_o_modulo_de_credenciais_e_valida_no_inicio(self):
        src = (ROOT / "app.py").read_text(encoding="utf-8")
        tree = ast.parse(src)
        for nome in ("conectar_sheets", "conectar_drive"):
            fonte = ast.get_source_segment(src, next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == nome))
            with self.subTest(funcao=nome):
                self.assertIn("credenciais_google(st.secrets, AMBIENTE_APP", fonte)
                self.assertNotIn("from_service_account_info", fonte)
        self.assertIn("TIPO_CREDENCIAL_GOOGLE = tipo_credencial(st.secrets, AMBIENTE_APP)", src)

    def test_repositorio_sem_tokens_chaves_ou_segredos(self):
        arquivos = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True).stdout.split()
        padroes = [r"1//0[A-Za-z0-9_-]{20,}", r"GOCSPX-[A-Za-z0-9_-]{10,}", r"-----BEGIN (RSA )?PRIVATE KEY-----",
                   r"ya29\.[A-Za-z0-9_-]{20,}", r'"refresh_token"\s*:\s*"[^"]{10,}"']
        for nome in arquivos:
            caminho = ROOT / nome
            if caminho.suffix.lower() not in {".py", ".md", ".toml", ".json", ".txt", ".sh", ".js", ".yml", ".yaml", ".cfg"}:
                continue
            texto = caminho.read_text(encoding="utf-8", errors="ignore")
            for padrao in padroes:
                with self.subTest(arquivo=nome, padrao=padrao):
                    self.assertIsNone(re.search(padrao, texto))
        ignorados = (ROOT / ".gitignore").read_text(encoding="utf-8")
        self.assertIn(".streamlit/secrets.toml", ignorados)
        self.assertIn("*.json", ignorados)


# ── 2. Gravação e leitura na planilha de teste ───────────────────────────────
class PlanilhaTeste(_Base):
    def test_abre_somente_a_planilha_de_teste(self):
        sh = self.app["conectar_sheets"]()
        self.assertEqual(sh.title, "ACQUA HOMOLOGACAO")
        gc = sys.modules["gspread"].authorize(ga.credenciais_google(self.secrets, "homologacao"))
        with self.assertRaises(Negado):
            gc.open_by_key(PROD_SHEET)  # drive.file: planilha não criada pelo app é inacessível

    def test_grava_e_le_aba_na_planilha_de_teste(self):
        sh = self.app["conectar_sheets"]()
        aba = sh.add_worksheet(title="👥 Clientes", rows=10, cols=9)
        aba.update(range_name="A1:C1", values=[["", "C1", "CONDOMÍNIO FICTÍCIO ALFA"]])
        lida = self.app["obter_aba_sheets"]("👥 Clientes").get_all_values()
        self.assertEqual(lida, [["", "C1", "CONDOMÍNIO FICTÍCIO ALFA"]])
        self.assertEqual(self.google.planilhas[PROD_SHEET].abas, {})


# ── 3. Fotografias no Drive de teste ─────────────────────────────────────────
class FotosTeste(_Base):
    def test_upload_estrutura_pastas_e_download(self):
        foto = b"\x89PNG-FOTO-FICTICIA"
        fid = self.app["drive_upload_foto"](foto, "antes.png", "Residencial D'Ávila Fictício", "2026-11")
        self.assertTrue(fid, self.app["st"].session_state.get("_sheets_ultimo_erro"))
        arq = self.google.arquivos[fid]
        pasta_mes = self.google.arquivos[arq["pais"][0]]
        pasta_cond = self.google.arquivos[pasta_mes["pais"][0]]
        self.assertEqual((pasta_mes["nome"], pasta_cond["nome"]), ("2026-11", "Residencial D'Ávila Fictício"))
        self.assertEqual(pasta_cond["pais"], [self.pasta_teste])
        self.assertEqual(self.app["drive_baixar_foto"](fid), foto)
        # Segundo envio reutiliza as pastas (consulta com apóstrofo escapado).
        fid2 = self.app["drive_upload_foto"](b"FOTO2", "depois.png", "Residencial D'Ávila Fictício", "2026-11")
        self.assertEqual(self.google.arquivos[fid2]["pais"], arq["pais"])

    def test_nao_grava_nem_le_fora_dos_recursos_de_teste(self):
        self.app["DRIVE_FOTOS_FOLDER_ID"] = PROD_PASTA
        self.assertIsNone(self.app["drive_upload_foto"](b"X", "a.png", "Cond", "2026-11"))
        self.assertEqual([a for a in self.google.arquivos.values() if PROD_PASTA in a["pais"]], [])
        self.assertIsNone(self.app["drive_baixar_foto"](PROD_PASTA))


# ── 4. Recuperação de rascunhos ──────────────────────────────────────────────
class RascunhosTeste(_Base):
    def test_salva_na_planilha_de_teste_e_recupera_sem_arquivo_local(self):
        dados = {"operador": "OPERADOR TESTE", "op_ph": "7,4", "foto_b64": "NAO-DEVE-IR"}
        self.assertTrue(self.app["salvar_rascunho_operador"]("Condomínio Fictício Beta", dict(dados), salvar_sheets=True))
        aba = self.google.planilhas[self.sheet_teste].abas.get("_Rascunhos")
        self.assertIsNotNone(aba, "a aba _Rascunhos deve ser criada na planilha de teste")
        self.assertEqual(len(aba.linhas), 2)
        self.assertNotIn("NAO-DEVE-IR", aba.linhas[1][3])
        # Simula outro dispositivo/reinício: apaga a cópia local e recupera da planilha.
        shutil.rmtree(Path(self._tmp.name))
        recuperado = self.app["carregar_rascunho_operador"]("Condomínio Fictício Beta")
        self.assertEqual((recuperado.get("op_ph"), recuperado.get("operador")), ("7,4", "OPERADOR TESTE"))
        self.assertEqual(self.app["carregar_rascunho_operador"]("Condomínio Fictício Gama"), {})
        self.assertEqual(self.google.planilhas[PROD_SHEET].abas, {})

    def test_regrava_mesma_linha_sem_duplicar(self):
        for ph in ("7,2", "7,6"):
            self.app["salvar_rascunho_operador"]("Cond Fictício", {"operador": "OP", "op_ph": ph}, salvar_sheets=True)
        aba = self.google.planilhas[self.sheet_teste].abas["_Rascunhos"]
        self.assertEqual(len(aba.linhas), 2)
        shutil.rmtree(Path(self._tmp.name))
        self.assertEqual(self.app["carregar_rascunho_operador"]("Cond Fictício")["op_ph"], "7,6")


if __name__ == "__main__":
    unittest.main()


# ── 5. Script de configuração (executado pelo usuário, aqui só com simulação) ─
class ScriptConfiguracao(unittest.TestCase):
    def _rodar(self, client_secret, escopos, refresh="token-ficticio"):
        import importlib.util
        spec = importlib.util.spec_from_file_location("oauth_setup", ROOT / "scripts" / "oauth_homologacao_setup.py")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        creds = types.SimpleNamespace(granted_scopes=escopos, refresh_token=refresh, client_id="cid", client_secret="cs")
        fluxo = mock.Mock()
        fluxo.run_local_server.return_value = creds
        drive = mock.Mock()
        drive.files.return_value.create.return_value.execute.return_value = {"id": "PASTA_TESTE_0001"}
        planilha = mock.Mock(id="PLANILHA_TESTE_0001")
        gs = types.ModuleType("gspread")
        gs.authorize = lambda c: mock.Mock(create=mock.Mock(return_value=planilha))
        oauthlib_flow = types.ModuleType("google_auth_oauthlib.flow")
        oauthlib_flow.InstalledAppFlow = mock.Mock(from_client_secrets_file=mock.Mock(return_value=fluxo))
        with mock.patch.dict(sys.modules, {"gspread": gs, "google_auth_oauthlib.flow": oauthlib_flow}), \
             mock.patch("googleapiclient.discovery.build", return_value=drive), \
             mock.patch.object(sys, "argv", ["setup", "--client-secret", str(client_secret)]):
            saida = __import__("io").StringIO()
            with mock.patch("sys.stdout", saida):
                codigo = mod.main()
        return codigo, saida.getvalue(), oauthlib_flow.InstalledAppFlow.from_client_secrets_file

    def test_pede_somente_drive_file_e_imprime_secrets_de_homologacao(self):
        with tempfile.TemporaryDirectory() as d:
            codigo, saida, flow = self._rodar(Path(d) / "client_secret.json", [DRIVE_FILE])
        self.assertEqual(codigo, 0)
        self.assertEqual(flow.call_args.kwargs["scopes"], [DRIVE_FILE])
        self.assertIn('ambiente = "homologacao"', saida)
        self.assertIn('sheets_id = "PLANILHA_TESTE_0001"', saida)

    def test_recusa_escopo_amplo_e_client_secret_dentro_do_projeto(self):
        with tempfile.TemporaryDirectory() as d:
            codigo, saida, _ = self._rodar(Path(d) / "c.json", [DRIVE_FILE, "https://www.googleapis.com/auth/drive"])
        self.assertEqual(codigo, 1)
        self.assertNotIn("refresh_token", saida)
        codigo, saida, flow = self._rodar(ROOT / "client_secret.json", [DRIVE_FILE])
        self.assertEqual(codigo, 1)
        flow.assert_not_called()
