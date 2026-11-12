"""Pastas de condomínio (dados contratuais, manifesto e documentos) persistem no Drive.

Simula o reinício do Streamlit Cloud apagando generated/ e confere, com as funções reais
do app.py e um Drive falso em memória (sem rede, dados fictícios):
- valores e datas dos contratos voltam iguais;
- documentos voltam com o mesmo conteúdo e a mesma data;
- painel de vencimentos, contadores e histórico iguais aos de antes do reinício;
- sem duplicação de arquivos e sem mistura entre empresas ou condomínios;
- exclusões não reaparecem; arquivo local mais novo nunca é sobrescrito.
"""
import json
import os
import re
import shutil
import tempfile
import time
import unittest
from datetime import date, timedelta
from pathlib import Path

from test_mensal_rt_app_pdf import _carregar
from utils import persistencia_pastas as pp

PASTA = "application/vnd.google-apps.folder"


class _Exec:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class DriveFalso:
    """Subconjunto do Drive v3 usado pelo módulo (list/create/update/get_media)."""

    def __init__(self):
        self.itens, self.seq, self.chamadas = {}, 0, []

    def files(self):
        return self

    @staticmethod
    def _valor(q, campo):
        m = re.search(campo + r"='((?:\\.|[^'\\])*)'", q)
        return re.sub(r"\\(.)", r"\1", m.group(1)) if m else None

    def list(self, q, fields=None, pageSize=None, pageToken=None, **_):
        def run():
            pai = re.sub(r"\\(.)", r"\1", re.search(r"'((?:\\.|[^'\\])*)' in parents", q).group(1))
            nome = self._valor(q, "name")
            res = []
            for i, it in self.itens.items():
                if it["trashed"] or pai not in it["parents"]:
                    continue
                if nome is not None and it["name"] != nome:
                    continue
                if f"mimeType='{PASTA}'" in q and it["mimeType"] != PASTA:
                    continue
                if f"mimeType!='{PASTA}'" in q and it["mimeType"] == PASTA:
                    continue
                res.append({"id": i, "name": it["name"], "mimeType": it["mimeType"],
                            "appProperties": dict(it["appProperties"])})
            return {"files": res}
        return _Exec(run)

    def create(self, body, media_body=None, fields=None, **_):
        def run():
            self.seq += 1
            i = f"id{self.seq}"
            self.itens[i] = {"name": body["name"], "parents": list(body.get("parents", [])),
                             "mimeType": body.get("mimeType", "arquivo"), "trashed": False,
                             "appProperties": dict(body.get("appProperties", {})),
                             "dados": media_body.getbytes(0, media_body.size()) if media_body else b""}
            self.chamadas.append(("create", body["name"]))
            return {"id": i}
        return _Exec(run)

    def update(self, fileId, body=None, media_body=None, fields=None, **_):
        def run():
            it = self.itens[fileId]
            body_ = body or {}
            if "trashed" in body_:
                it["trashed"] = body_["trashed"]
                for filho in self.itens.values():  # pasta na lixeira leva o conteúdo junto
                    if fileId in filho["parents"]:
                        filho["trashed"] = body_["trashed"]
            it["appProperties"].update(body_.get("appProperties", {}))
            if media_body is not None:
                it["dados"] = media_body.getbytes(0, media_body.size())
            self.chamadas.append(("update", it["name"]))
            return {"id": fileId}
        return _Exec(run)

    def get_media(self, fileId):
        return _Exec(lambda: self.itens[fileId]["dados"])

    # auxiliares de teste
    def ativos(self, nome=None):
        return [it for it in self.itens.values() if not it["trashed"] and it["mimeType"] != PASTA
                and (nome is None or it["name"] == nome)]

    def novo_root(self):
        self.seq += 1
        self.itens["root"] = {"name": "fotos", "parents": [], "mimeType": PASTA, "trashed": False,
                              "appProperties": {}, "dados": b""}
        return "root"


def _escrever(caminho: Path, conteudo, mtime=None):
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_bytes(conteudo if isinstance(conteudo, bytes) else conteudo.encode("utf-8"))
    if mtime:
        os.utime(caminho, (mtime, mtime))


class ModuloPersistencia(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.raiz = self.tmp / "generated"
        self.fake = DriveFalso()
        self.drive = pp.DrivePastas(self.fake, self.fake.novo_root())

    def test_sincroniza_so_o_necessario_sem_duplicar(self):
        pasta = self.raiz / "condominio_alfa"
        _escrever(pasta / "dados_condominio.json", '{"valor_mensal": "1.234,56"}', 1_700_000_000)
        _escrever(pasta / "Contrato_Alfa.docx", b"DOCX-ALFA", 1_700_000_100)
        _escrever(pasta / "_rascunho_operador.json", "{}")          # rascunho: fora
        _escrever(pasta / "fotos_campo" / "foto.jpg", b"JPG")       # subpasta: fora
        self.assertEqual(pp.sincronizar_pasta(self.drive, pasta), 2)
        self.assertEqual(pp.sincronizar_pasta(self.drive, pasta), 0)   # nada mudou
        _escrever(pasta / "dados_condominio.json", '{"valor_mensal": "2.000,00"}', 1_700_000_500)
        self.assertEqual(pp.sincronizar_pasta(self.drive, pasta), 1)
        self.assertEqual(sorted(it["name"] for it in self.fake.ativos()), ["Contrato_Alfa.docx", "dados_condominio.json"])
        self.assertEqual(self.fake.ativos("dados_condominio.json")[0]["dados"], b'{"valor_mensal": "2.000,00"}')

    def test_restaura_conteudo_e_data_preserva_local_mais_novo(self):
        alfa, beta = self.raiz / "condominio_alfa", self.raiz / "residencial_d_avila"
        _escrever(alfa / "dados_condominio.json", '{"empresa": "Aqua Gestão"}', 1_700_000_000)
        _escrever(alfa / "Contrato_Alfa.pdf", b"%PDF-ALFA", 1_700_000_100)
        _escrever(beta / "dados_condominio.json", '{"empresa": "Bem Star Piscinas"}', 1_700_000_200)
        for p in (alfa, beta):
            pp.sincronizar_pasta(self.drive, p)
        shutil.rmtree(self.raiz)                                       # reinício: disco vazio
        drive2 = pp.DrivePastas(self.fake, "root")
        resumo = pp.restaurar(drive2, self.raiz)
        self.assertEqual((resumo["pastas"], resumo["baixados"], resumo["enviados"]), (2, 3, 0))
        self.assertEqual((alfa / "Contrato_Alfa.pdf").read_bytes(), b"%PDF-ALFA")
        self.assertEqual(int((alfa / "Contrato_Alfa.pdf").stat().st_mtime), 1_700_000_100)
        self.assertEqual(json.loads((beta / "dados_condominio.json").read_text(encoding="utf-8"))["empresa"],
                         "Bem Star Piscinas")
        self.assertFalse((beta / "Contrato_Alfa.pdf").exists())          # sem mistura
        # Local mais novo não é sobrescrito; vai para o Drive.
        _escrever(alfa / "dados_condominio.json", '{"empresa": "Aqua Gestão", "novo": 1}', time.time())
        resumo = pp.restaurar(pp.DrivePastas(self.fake, "root"), self.raiz)
        self.assertEqual((resumo["baixados"], resumo["enviados"]), (0, 1))
        self.assertIn('"novo": 1', (alfa / "dados_condominio.json").read_text(encoding="utf-8"))
        self.assertEqual(len(self.fake.ativos("dados_condominio.json")), 2)  # um por condomínio

    def test_pastas_internas_nao_sao_copiadas_e_exclusao_vai_para_lixeira(self):
        _escrever(self.raiz / "_rt_contextos" / "abc" / "dados_condominio.json", "{}")
        self.assertFalse(pp.pasta_persistivel(self.raiz / "_rt_contextos", self.raiz))
        pasta = self.raiz / "condominio_alfa"
        _escrever(pasta / "dados_condominio.json", "{}", 1_700_000_000)
        _escrever(pasta / "Aditivo.docx", b"ADT", 1_700_000_000)
        pp.sincronizar_pasta(self.drive, pasta)
        pp.excluir_arquivo(self.drive, pasta, "Aditivo.docx")
        self.assertEqual([it["name"] for it in self.fake.ativos()], ["dados_condominio.json"])
        pp.excluir_pasta(self.drive, pasta)
        self.assertEqual(self.fake.ativos(), [])
        shutil.rmtree(self.raiz)
        self.assertEqual(pp.restaurar(pp.DrivePastas(self.fake, "root"), self.raiz)["baixados"], 0)
        self.assertFalse((self.raiz / "condominio_alfa").exists())


class _StTeste:
    """st mínimo: cache_resource real (por instância) e session_state."""

    class _Sessao(dict):
        __getattr__ = dict.get

        def __setattr__(self, k, v):
            self[k] = v

    def __init__(self):
        self.session_state, self.secrets, self.avisos = self._Sessao(), {}, []

    def cache_resource(self, **_):
        def deco(fn):
            memo = {}
            def wrapper():
                if "v" not in memo:
                    memo["v"] = fn()
                return memo["v"]
            return wrapper
        return deco

    def warning(self, texto):
        self.avisos.append(texto)

    def __getattr__(self, _nome):
        return lambda *a, **k: None


class AppAposReinicio(unittest.TestCase):
    """Funções reais do app.py: grava, 'reinicia' (disco vazio + processo novo) e compara."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.fake = DriveFalso()
        self.root = self.fake.novo_root()

    def _app(self):
        ns = _carregar()
        st = _StTeste()
        ns.update(json=json, shutil=shutil, os=os, time=time)
        ns.update(st=st, GENERATED_DIR=self.tmp / "generated", DRIVE_FOTOS_FOLDER_ID=self.root,
                  AMBIENTE_APP="homologacao", _servico_drive_persistencia=lambda: self.fake,
                  sheets_listar_clientes_completo=lambda: [],
                  _log_sheets_erro=lambda ctx, e: (_ for _ in ()).throw(e))
        ns["GENERATED_DIR"].mkdir(parents=True, exist_ok=True)
        ns["_st"] = st
        return ns

    def _estado(self, ns):
        painel = ns["listar_painel_vencimentos"](30)
        hist = ns["listar_historico"]()
        return (
            sorted((i["slug"], i["data_fim"], i["status"]["codigo"], i["tem_json"],
                    (i["ultimo_contrato"] or {}).get("name"), (i["ultimo_aditivo"] or {}).get("name"),
                    len(i["arquivos"])) for i in painel),
            sorted((h["nome"], h["total_arquivos"], h["data_fim"], h["tem_dados_salvos"],
                    tuple((a["name"], a["modificado_em"]) for a in h["arquivos"])) for h in hist),
        )

    def test_reinicio_preserva_contratos_documentos_historico_e_contadores(self):
        fim_alfa = (date.today() + timedelta(days=200)).strftime("%d/%m/%Y")
        fim_beta = (date.today() + timedelta(days=10)).strftime("%d/%m/%Y")
        ns = self._app()
        gen = ns["GENERATED_DIR"]
        alfa, beta = gen / "condominio_alfa_ficticio", gen / "residencial_beta_ficticio"
        dados_alfa = {"nome_condominio": "Condomínio Alfa Fictício", "empresa": "Aqua Gestão",
                      "servicos": {"rt": True, "limpeza": False}, "valor_mensal": "1.234,56",
                      "data_inicio": "01/01/2026", "data_fim": fim_alfa, "data_assinatura": "02/01/2026"}
        dados_beta = {"nome_condominio": "Residencial Beta Fictício", "empresa": "Aqua Gestão",
                      "servicos": {"rt": True, "limpeza": False}, "valor_mensal": "980,00",
                      "data_inicio": "01/02/2026", "data_fim": fim_beta, "data_assinatura": "03/02/2026"}
        for pasta, dados in ((alfa, dados_alfa), (beta, dados_beta)):
            pasta.mkdir()
            ns["salvar_dados_condominio"](pasta, dados)
        _escrever(alfa / "Contrato_Condominio_Alfa_Ficticio.docx", b"DOCX-CONTRATO-ALFA", 1_700_000_000)
        _escrever(alfa / "Contrato_Condominio_Alfa_Ficticio.pdf", b"%PDF-CONTRATO-ALFA", 1_700_000_050)
        ns["registrar_documento_manifest"](alfa, "Condomínio Alfa Fictício", "Contrato",
                                           alfa / "Contrato_Condominio_Alfa_Ficticio.docx",
                                           alfa / "Contrato_Condominio_Alfa_Ficticio.pdf", True, None)
        _escrever(beta / "Aditivo_Residencial_Beta_Ficticio.docx", b"DOCX-ADITIVO-BETA", 1_700_000_100)
        ns["registrar_documento_manifest"](beta, "Residencial Beta Fictício", "Aditivo",
                                           beta / "Aditivo_Residencial_Beta_Ficticio.docx", None, False, "sem PDF")
        # Pastas internas do app (criadas conforme o uso da sessão) não contam como condomínio.
        (gen / "_rt_contextos" / "abc").mkdir(parents=True)
        (gen / "_rascunhos_rt_sync").mkdir()
        antes = self._estado(ns)
        bytes_antes = {p.relative_to(gen): p.read_bytes() for p in gen.rglob("*")
                       if p.is_file() and not p.relative_to(gen).parts[0].startswith("_")}
        self.assertEqual([x[0] for x in antes[0]], ["condominio_alfa_ficticio", "residencial_beta_ficticio"])
        self.assertEqual([x[0] for x in antes[1]], ["condominio_alfa_ficticio", "residencial_beta_ficticio"])
        self.assertEqual(ns["_st"].avisos, [])

        shutil.rmtree(gen)                       # reinício do Streamlit Cloud
        ns2 = self._app()                        # processo novo (cache vazio)
        depois = self._estado(ns2)
        self.assertEqual(depois, antes)          # painel, contadores, histórico e datas iguais
        bytes_depois = {p.relative_to(gen): p.read_bytes() for p in gen.rglob("*") if p.is_file()}
        self.assertEqual(bytes_depois, bytes_antes)
        self.assertEqual(json.loads((alfa / "dados_condominio.json").read_text(encoding="utf-8"))["valor_mensal"],
                         "1.234,56")
        manifest = ns2["carregar_manifest_condominio"](beta)
        self.assertEqual([d["arquivo_docx"] for d in manifest["documentos"]], ["Aditivo_Residencial_Beta_Ficticio.docx"])
        # Um único arquivo por nome em cada condomínio, sem mistura entre pastas.
        for nome in ("dados_condominio.json", "manifest.json"):
            self.assertEqual(len(self.fake.ativos(nome)), 2)
        self.assertEqual(len(self.fake.ativos("Contrato_Condominio_Alfa_Ficticio.pdf")), 1)
        self.assertFalse((beta / "Contrato_Condominio_Alfa_Ficticio.pdf").exists())

        # Exclusões feitas pelo app não voltam depois de outro reinício.
        ns2["deletar_arquivo_individual"](alfa / "Contrato_Condominio_Alfa_Ficticio.docx")
        ns2["deletar_pasta_condominio"](beta)
        shutil.rmtree(gen)
        ns3 = self._app()
        slugs = [i["slug"] for i in ns3["listar_painel_vencimentos"](30)]
        self.assertEqual(slugs, ["condominio_alfa_ficticio"])
        self.assertFalse((alfa / "Contrato_Condominio_Alfa_Ficticio.docx").exists())
        self.assertTrue((alfa / "Contrato_Condominio_Alfa_Ficticio.pdf").exists())

    def test_empresas_continuam_separadas_apos_reinicio(self):
        ns = self._app()
        gen = ns["GENERATED_DIR"]
        aqua, bem = gen / "condominio_alfa_ficticio", gen / "condominio_gama_ficticio"
        for pasta, emp, srv in ((aqua, "Aqua Gestão", {"rt": True, "limpeza": False}),
                                (bem, "Bem Star Piscinas", {"rt": False, "limpeza": True})):
            pasta.mkdir()
            ns["salvar_dados_condominio"](pasta, {"nome_condominio": pasta.name, "empresa": emp, "servicos": srv,
                                                  "data_fim": "31/12/2027"})
        shutil.rmtree(gen)
        ns2 = self._app()
        ns2["garantir_pastas_restauradas"]()   # como painel e histórico fazem antes de listar
        pastas = sorted(p for p in gen.iterdir() if p.is_dir())
        self.assertEqual([p.name for p in pastas], ["condominio_alfa_ficticio", "condominio_gama_ficticio"])
        self.assertEqual([p.name for p in ns2["filtrar_pastas_por_empresa"](pastas, "aqua_gestao", [])],
                         ["condominio_alfa_ficticio"])
        self.assertEqual([p.name for p in ns2["filtrar_pastas_por_empresa"](pastas, "bem_star", [])],
                         ["condominio_gama_ficticio"])

    def test_drive_falha_logo_apos_reinicio_sem_perder_historico(self):
        ns = self._app()
        gen = ns["GENERATED_DIR"]
        alfa = gen / "condominio_alfa_ficticio"
        alfa.mkdir()
        ns["salvar_dados_condominio"](alfa, {"data_fim": "31/12/2027"})
        _escrever(alfa / "Contrato_1.docx", b"C1", 1_700_000_000)
        ns["registrar_documento_manifest"](alfa, "Alfa", "Contrato", alfa / "Contrato_1.docx", None, False, None)
        shutil.rmtree(gen)
        # Reinício com o Drive fora do ar (falha passageira): o novo documento só fica no disco.
        ns2 = self._app()
        ns2["_servico_drive_persistencia"] = lambda: (_ for _ in ()).throw(TimeoutError("rede fora do ar"))
        ns2["_log_sheets_erro"] = lambda ctx, e: None
        alfa.mkdir()
        _escrever(alfa / "Aditivo_2.docx", b"A2")
        ns2["registrar_documento_manifest"](alfa, "Alfa", "Aditivo", alfa / "Aditivo_2.docx", None, False, None)
        self.assertEqual(len(self.fake.ativos("manifest.json")), 1)
        self.assertNotIn(b"Aditivo_2", self.fake.ativos("manifest.json")[0]["dados"])  # nada sobrescrito
        # Drive volta: a restauração mescla os dois históricos e envia o resultado.
        ns2["_servico_drive_persistencia"] = lambda: self.fake
        ns2["_estado_persistencia_pastas"]()["ultima_tentativa"] = 0.0
        ns2["garantir_pastas_restauradas"]()
        docs = json.loads(self.fake.ativos("manifest.json")[0]["dados"])["documentos"]
        self.assertEqual([d["arquivo_docx"] for d in docs], ["Contrato_1.docx", "Aditivo_2.docx"])
        self.assertEqual(self.fake.ativos("Aditivo_2.docx")[0]["dados"], b"A2")
        self.assertTrue((alfa / "Contrato_1.docx").exists())  # restaurado junto

    def test_valores_e_datas_alterados_depois_do_nome_sao_os_gravados(self):
        ns = self._app()
        ss = ns["_st"].session_state
        ss.update(nome_condominio="Condomínio Alfa Fictício", valor_mensal="900,00",
                  data_inicio="01/01/2026", data_fim="31/12/2026")
        ns["on_change_nome_condominio"]()                    # sai do campo do nome
        ss["valor_mensal"] = "1.450,00"                       # depois altera valor e datas
        ns["on_change_valor_mensal"]()
        ss["data_inicio"] = "05022026"
        ns["on_change_data_inicio"]()
        ss["data_fim"] = "04/02/2028"
        ns["on_change_data_fim"]()
        pasta = ns["GENERATED_DIR"] / ns["slugify_nome"]("Condomínio Alfa Fictício")
        local = json.loads((pasta / "dados_condominio.json").read_text(encoding="utf-8"))
        remoto = json.loads(self.fake.ativos("dados_condominio.json")[0]["dados"])
        for dados in (local, remoto):
            self.assertEqual((dados["valor_mensal"], dados["data_inicio"], dados["data_fim"]),
                             (ss["valor_mensal"], "05/02/2026", "04/02/2028"))
            self.assertNotEqual(dados["valor_mensal"], "900,00")
        self.assertEqual(len(self.fake.ativos("dados_condominio.json")), 1)
        shutil.rmtree(ns["GENERATED_DIR"])
        ns2 = self._app()
        item = ns2["listar_painel_vencimentos"](30)[0]
        self.assertEqual((item["data_fim"], item["dados"]["valor_mensal"], item["dados"]["data_inicio"]),
                         ("04/02/2028", ss["valor_mensal"], "05/02/2026"))

    def test_documento_restaurado_pode_ser_baixado(self):
        ns = self._app()
        botoes = []
        ns["_st"].download_button = lambda rotulo, **kw: botoes.append((rotulo, kw))
        pasta = ns["GENERATED_DIR"] / "condominio_alfa_ficticio"
        _escrever(pasta / "Caderno_POPs_Alfa.pdf", b"%PDF-POPS", 1_700_000_000)
        ns["persistir_pasta_condominio"](pasta)
        ns["_estado_persistencia_pastas"]()["restaurado"] = True
        ns["persistir_pasta_condominio"](pasta)
        shutil.rmtree(ns["GENERATED_DIR"])
        ns2 = self._app()
        ns2["_st"].download_button = lambda rotulo, **kw: botoes.append((rotulo, kw))
        ns2["garantir_pastas_restauradas"]()
        ns2["botao_baixar_documento"](pasta / "Caderno_POPs_Alfa.pdf", "baixar_x")
        ns2["botao_baixar_documento"](pasta / "nao_existe.pdf", "baixar_y")
        self.assertEqual(len(botoes), 1)
        rotulo, kw = botoes[0]
        self.assertEqual((rotulo, kw["file_name"], kw["mime"], kw["key"]),
                         ("Baixar", "Caderno_POPs_Alfa.pdf", "application/pdf", "baixar_x"))
        self.assertTrue(callable(kw["data"]))                 # lido só no clique
        self.assertEqual(kw["data"](), b"%PDF-POPS")
        fonte = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
        self.assertIn('botao_baixar_documento(caminho, f"baixar_arq_{file_key}")', fonte)       # histórico
        for rot in ("contrato", "aditivo", "relatório"):
            self.assertIn(f'"Baixar último {rot}")', fonte)                                    # painel

    def test_copia_nao_configurada_nao_bloqueia_gravacao_e_avisa_corretamente(self):
        ns = self._app()
        ns["_servico_drive_persistencia"] = lambda: None             # produção sem [google_oauth_drive]
        ns["_log_sheets_erro"] = lambda ctx, e: None
        pasta = ns["GENERATED_DIR"] / "condominio_alfa_ficticio"
        pasta.mkdir()
        ns["salvar_dados_condominio"](pasta, {"data_fim": "31/12/2027"})
        self.assertTrue((pasta / "dados_condominio.json").exists())   # gravado no disco
        self.assertEqual(len(ns["_st"].avisos), 1)                      # aviso único na sessão
        aviso = ns["_st"].avisos[0]
        self.assertIn("Cópia persistente pendente: 1 pasta(s)", aviso)
        self.assertIn("não configurada", aviso)
        self.assertIn("Não haverá nova tentativa", aviso)
        self.assertNotIn("reenviada", aviso)                            # nada de promessa falsa
        ns["salvar_dados_condominio"](pasta, {"data_fim": "30/12/2027"})
        self.assertEqual(len(ns["_st"].avisos), 1)

    def test_erro_permanente_do_drive_para_as_tentativas(self):
        class Http403(Exception):
            resp = type("R", (), {"status": 403})()

            def __str__(self):
                return '<HttpError 403 returned "Service Accounts do not have storage quota.">'

        chamadas = []

        class DriveSemCota(DriveFalso):
            def create(self, body, media_body=None, fields=None, **_):
                chamadas.append(body["name"])
                return _Exec(lambda: (_ for _ in ()).throw(Http403()))

        ns = self._app()
        self.fake = DriveSemCota()
        self.fake.novo_root()
        ns["_servico_drive_persistencia"] = lambda: self.fake
        ns["_log_sheets_erro"] = lambda ctx, e: None
        pasta = ns["GENERATED_DIR"] / "condominio_alfa_ficticio"
        _escrever(pasta / "dados_condominio.json", "{}")
        ns["garantir_pastas_restauradas"]()
        tentativas = len(chamadas)
        self.assertGreaterEqual(tentativas, 1)
        situacao = ns["situacao_persistencia_pastas"]()
        self.assertEqual((situacao["ok"], situacao["permanente"], situacao["pendentes"]), (False, True, 1))
        self.assertIn("não tem cota", situacao["motivo"])
        for _ in range(3):                                   # sem novas chamadas ao Drive
            ns["_estado_persistencia_pastas"]()["ultima_tentativa"] = 0.0
            ns["garantir_pastas_restauradas"]()
            ns["salvar_dados_condominio"](pasta, {"data_fim": "31/12/2027"})
        self.assertEqual(len(chamadas), tentativas)
        self.assertTrue((pasta / "dados_condominio.json").exists())
        texto = ns["texto_aviso_persistencia_pastas"](ns["situacao_persistencia_pastas"]())
        self.assertIn("guardadas só neste servidor", texto)
        self.assertIn("Não haverá nova tentativa", texto)

    def test_falha_passageira_tenta_de_novo(self):
        ns = self._app()
        ns["_log_sheets_erro"] = lambda ctx, e: None
        ns["_servico_drive_persistencia"] = lambda: (_ for _ in ()).throw(TimeoutError("rede"))
        pasta = ns["GENERATED_DIR"] / "condominio_alfa_ficticio"
        _escrever(pasta / "dados_condominio.json", "{}")
        ns["garantir_pastas_restauradas"]()
        situacao = ns["situacao_persistencia_pastas"]()
        self.assertEqual((situacao["permanente"], situacao["pendentes"]), (False, 1))
        self.assertIn("Nova tentativa automática", ns["texto_aviso_persistencia_pastas"](situacao))
        ns["_servico_drive_persistencia"] = lambda: self.fake
        ns["_estado_persistencia_pastas"]()["ultima_tentativa"] = 0.0
        ns["garantir_pastas_restauradas"]()
        self.assertTrue(ns["situacao_persistencia_pastas"]()["ok"])
        self.assertEqual(len(self.fake.ativos("dados_condominio.json")), 1)

    def test_producao_grava_e_restaura_na_pasta_propria_da_raiz(self):
        ns = self._app()
        ns["AMBIENTE_APP"] = "producao"
        ns["DRIVE_FOTOS_FOLDER_ID"] = "pasta_de_fotos_inacessivel"   # drive.file não alcança
        pasta = ns["GENERATED_DIR"] / "ASSOCIAÇÃO_TERRAS_ALTAS"
        _escrever(pasta / "dados_condominio.json", '{"nome_condominio": "ASSOCIAÇÃO TERRAS ALTAS"}', 1_700_000_000)
        ns["garantir_pastas_restauradas"]()
        bases = [it for it in self.fake.itens.values() if it["mimeType"] == PASTA and "root" in it["parents"]]
        self.assertEqual([b["name"] for b in bases], [ns["NOME_BASE_PASTAS_PRODUCAO"]])
        self.assertEqual(len(self.fake.ativos("dados_condominio.json")), 1)
        shutil.rmtree(ns["GENERATED_DIR"])
        ns2 = self._app()
        ns2["AMBIENTE_APP"] = "producao"
        ns2["DRIVE_FOTOS_FOLDER_ID"] = "pasta_de_fotos_inacessivel"
        ns2["garantir_pastas_restauradas"]()
        self.assertEqual(json.loads((pasta / "dados_condominio.json").read_text(encoding="utf-8"))["nome_condominio"],
                         "ASSOCIAÇÃO TERRAS ALTAS")
        self.assertTrue(ns2["situacao_persistencia_pastas"]()["ok"])


class CredencialDrivePersistencia(unittest.TestCase):
    def test_producao_usa_secao_propria_so_com_drive_file(self):
        from unittest.mock import patch
        from utils import google_auth as ga
        self.assertIsNone(ga.credenciais_drive_persistencia({}, "producao"))      # não configurada
        with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
            ga.credenciais_drive_persistencia({"google_oauth_drive": {"client_id": "x"}}, "producao")
        amplo = {"client_id": "a", "client_secret": "b", "refresh_token": "c",
                 "scopes": ["https://www.googleapis.com/auth/drive"]}
        with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
            ga.credenciais_drive_persistencia({"google_oauth_drive": amplo}, "producao")
        ok = dict(amplo, scopes=[ga.ESCOPO_OAUTH_HOMOLOGACAO])
        with patch.object(ga, "credenciais_oauth", return_value="CREDS") as m:
            self.assertEqual(ga.credenciais_drive_persistencia({"google_oauth_drive": ok}, "producao"), "CREDS")
            self.assertEqual(m.call_args.kwargs["rotulo"], "produção (cópia persistente)")
        # A conta de serviço continua sendo a credencial principal; [google_oauth] segue recusado.
        self.assertEqual(ga.tipo_credencial({"google_oauth_drive": ok}, "producao"), "conta_servico")
        with self.assertRaises(ga.ConfiguracaoGoogleInvalida):
            ga.tipo_credencial({"google_oauth": ok}, "producao")


if __name__ == "__main__":
    unittest.main()
