"""Testes do formulário do operador e do fluxo RT: rascunho, limpeza e troca de condomínio.

Carregam somente as funções envolvidas do app.py, com session_state simulado e
pasta temporária. Nenhum acesso a Google Sheets, Drive ou rede; dados fictícios.

Executar:  python -m unittest discover -s tests -v
"""

import ast
import atexit
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
FUNCOES = [
    "_slug_piscina_operador", "coletar_rascunho_operador", "restaurar_rascunho_operador",
    "limpar_formulario_operador", "_chave_fotos_operador", "trocar_condominio_operador",
    "chave_fotos_rt", "_campos_rt_do_condominio", "limpar_formulario_rt_operador", "_rascunho_rt_operador_path",
    "trocar_condominio_rt_operador",
    "salvar_rascunho_operador",
]
CONSTANTES = {"_SLUG_PISCINAS_OPERADOR", "_RE_CAMPO_PISCINA_OP", "_RE_CAMPO_DOSAGEM_OP", "_RT_OP_PREFIXOS_TEXTO", "_RT_OP_PREFIXOS", "_RT_OP_CAMPOS",
              "_RT_OP_FORA_DO_RASCUNHO"}


def _nos():
    nos = [n for n in TREE.body if isinstance(n, ast.Assign)
           and any(isinstance(t, ast.Name) and t.id in CONSTANTES for t in n.targets)]
    for nome in FUNCOES:  # função ausente (versão antiga do app): só os testes que a usam falham
        nos.extend(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == nome)
    return nos


CODIGO = compile(ast.Module(body=_nos(), type_ignores=[]), str(APP), "exec")


class _St:
    def __init__(self):
        self.session_state = {}


def _sessao(resultado_salvar=True, pasta=None):
    """resultado_salvar: retorno do salvar_rascunho_operador simulado (ou uma exceção a lançar).
    pasta: GENERATED_DIR compartilhado (para simular duas sessões no mesmo servidor)."""
    if pasta is None:
        pasta = Path(tempfile.mkdtemp(prefix="operador_teste_"))
        atexit.register(shutil.rmtree, pasta, ignore_errors=True)
    salvos = []

    def salvar(nome, dados, salvar_sheets=False):
        salvos.append((nome, dados, salvar_sheets))
        if isinstance(resultado_salvar, BaseException):
            raise resultado_salvar
        return resultado_salvar

    ns = {"st": _St(), "re": re, "json": json, "os": os, "datetime": datetime, "Path": Path, "GENERATED_DIR": pasta,
          "slugify_nome": slugify_nome, "_agora_brasilia": lambda: "26/10/2026 10:00:00",
          "_rascunho_usa_sheets": lambda: False}
    exec(CODIGO, ns)
    ns["salvar_rascunho_operador_real"] = ns["salvar_rascunho_operador"]
    ns["salvar_rascunho_operador"] = salvar  # trocar_condominio_operador usa o simulado
    ns["ss"] = ns["st"].session_state
    ns["salvos"] = salvos
    ns["pasta"] = pasta
    return ns


def _preencher_visita(ss):
    ss.update({
        "op_data_visita": "26/10/2026", "op_operador": "Operador Teste",
        "op_adulto_ph": "7.4", "op_adulto_crl": "1.5", "op_adulto_ct": "1.8", "op_adulto_cya": "0",
        "op_infantil_ph": "7.2", "op_infantil_crl": "0",
        "op_dos_adulto_prod_0": "Cloro granulado", "op_dos_adulto_qtd_0": "0", "op_dos_adulto_un_0": "kg",
        "op_dos_adulto_fin_0": "Manutenção", "op_dos_infantil_prod_0": "Barrilha", "op_dos_infantil_qtd_0": "0,5",
        "op_obs_campo": "Água cristalina", "op_problemas": "Bomba com ruído", "op_resp_local": "Zelador Teste",
        "op_parecer_visita": "⚠️ Aceitável com ajustes pontuais",
    })


PISCINAS = ["Piscina Adulto", "Piscina Infantil"]


class TestRascunhoOperador(unittest.TestCase):
    def test_coletar_e_restaurar_preserva_campos_zeros_e_dosagens_por_piscina(self):
        a = _sessao()
        _preencher_visita(a["ss"])
        rasc = json.loads(json.dumps(a["coletar_rascunho_operador"]("Condomínio Fictício", PISCINAS)))  # como no arquivo
        b = _sessao()
        b["restaurar_rascunho_operador"](rasc)
        for k, v in a["ss"].items():
            self.assertEqual(b["ss"].get(k), v, k)

    def test_piscina_com_nome_personalizado_usa_a_mesma_chave_do_formulario(self):
        a = _sessao()
        slug = a["_slug_piscina_operador"]("Piscina Olímpica")
        self.assertEqual(slug, slugify_nome("Piscina Olímpica")[:12])
        a["ss"].update({f"op_{slug}_ph": "7.6", f"op_{slug}_cya": "0", f"op_dos_{slug}_prod_0": "Algicida"})
        rasc = a["coletar_rascunho_operador"]("Condomínio Fictício", ["Piscina Olímpica"])
        b = _sessao()
        b["restaurar_rascunho_operador"](rasc)
        self.assertEqual(b["ss"][f"op_{slug}_ph"], "7.6")
        self.assertEqual(b["ss"][f"op_{slug}_cya"], "0")
        self.assertEqual(b["ss"][f"op_dos_{slug}_prod_0"], "Algicida")

    def test_rascunho_antigo_com_lista_geral_de_dosagens_vai_para_a_primeira_piscina(self):
        antigo = {"piscinas": [{"nome": "Piscina Adulto", "ph": "7.3"}],
                  "dosagens": [{"produto": "Cloro", "quantidade": "1", "unidade": "kg", "finalidade": "Rotina"}]}
        b = _sessao()
        b["restaurar_rascunho_operador"](antigo)
        self.assertEqual(b["ss"]["op_adulto_ph"], "7.3")
        self.assertEqual(b["ss"]["op_dos_adulto_prod_0"], "Cloro")
        self.assertEqual(b["ss"]["op_dos_adulto_qtd_0"], "1")


class TestLimpezaETroca(unittest.TestCase):
    def test_limpar_cobre_piscinas_personalizadas_e_dosagens_por_piscina(self):
        s = _sessao()
        _preencher_visita(s["ss"])
        slug = s["_slug_piscina_operador"]("Piscina Olímpica")
        s["ss"].update({f"op_{slug}_ph": "7.6", f"op_dos_{slug}_qtd_2": "3", "_sug_pisc_adulto": [{"produto": "x"}],
                        "op_ph_apos": "7.5"})
        s["limpar_formulario_operador"]()
        ss = s["ss"]
        for k in ("op_adulto_ph", "op_adulto_cya", "op_infantil_crl", f"op_{slug}_ph", "op_dos_adulto_prod_0",
                  "op_dos_adulto_qtd_0", "op_dos_infantil_prod_0", f"op_dos_{slug}_qtd_2", "op_obs_campo",
                  "op_problemas", "op_resp_local", "op_ph_apos"):
            self.assertEqual(ss[k], "", k)
        self.assertNotIn("_sug_pisc_adulto", ss)
        self.assertEqual(ss["op_parecer_visita"], "✅ Satisfatório")
        self.assertEqual((ss["op_data_visita"], ss["op_operador"]), ("26/10/2026", "Operador Teste"))  # identidade/data ficam

    def test_trocar_condominio_guarda_rascunho_do_anterior_e_isola_o_novo(self):
        s = _sessao()
        _preencher_visita(s["ss"])
        s["ss"]["_op_piscinas_form"] = PISCINAS
        chave_antes = s["_chave_fotos_operador"]("antes")
        s["ss"]["_rascunho_visto_Residencial_Beta"] = True
        s["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta")
        self.assertEqual(len(s["salvos"]), 1)
        nome, dados, sheets = s["salvos"][0]
        self.assertEqual((nome, sheets), ("Condomínio Alfa", False))
        self.assertEqual(dados["piscinas"][0]["ph"], "7.4")
        self.assertEqual(dados["dosagens_por_piscina"]["Piscina Adulto"][0]["quantidade"], "0")
        self.assertEqual(s["ss"]["op_adulto_ph"], "")  # formulário do novo vazio
        self.assertNotEqual(s["_chave_fotos_operador"]("antes"), chave_antes)  # fotos do anterior não seguem
        self.assertNotIn("_rascunho_visto_Residencial_Beta", s["ss"])  # aviso de rascunho do novo volta a aparecer
        self.assertEqual(s["ss"]["_op_cond_form"], "Residencial Beta")

    def test_trocar_sem_conteudo_nao_grava_rascunho_vazio(self):
        s = _sessao()
        s["ss"]["op_parecer_visita"] = "✅ Satisfatório"
        s["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta")
        self.assertEqual(s["salvos"], [])

    def _troca_que_falha(self, resultado):
        s = _sessao(resultado_salvar=resultado)
        _preencher_visita(s["ss"])
        s["ss"]["_op_piscinas_form"] = PISCINAS
        antes = dict(s["ss"])
        chave_antes = s["_chave_fotos_operador"]("antes")
        self.assertFalse(s["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta"))
        bloq = s["ss"].pop("_op_troca_bloqueada")
        self.assertEqual((bloq["anterior"], bloq["novo"]), ("Condomínio Alfa", "Residencial Beta"))
        self.assertTrue(bloq["motivo"])
        self.assertEqual(s["ss"], antes)  # nenhum campo limpo, nada trocado
        self.assertEqual(s["_chave_fotos_operador"]("antes"), chave_antes)  # fotos continuam anexadas
        return bloq

    def test_troca_cancelada_se_salvar_rascunho_retorna_falha(self):
        self._troca_que_falha(False)

    def test_troca_cancelada_se_salvar_rascunho_lanca_excecao(self):
        self.assertIn("OSError", self._troca_que_falha(OSError("disco cheio (simulado)"))["motivo"])

    def test_troca_cancelada_se_salvar_retorna_nada(self):
        self._troca_que_falha(None)

    def test_troca_cancelada_com_assinatura_coletada(self):
        s = _sessao()
        _preencher_visita(s["ss"])
        s["ss"]["op_assinatura_responsavel_b64"] = "iVBORw0KGgo="  # assinatura fictícia
        self.assertFalse(s["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertIn("assinatura", s["ss"]["_op_troca_bloqueada"]["motivo"])
        self.assertEqual(s["ss"]["op_assinatura_responsavel_b64"], "iVBORw0KGgo=")
        self.assertEqual(s["ss"]["op_adulto_ph"], "7.4")

    def test_troca_bem_sucedida_remove_aviso_antigo(self):
        s = _sessao()
        _preencher_visita(s["ss"])
        s["ss"]["_op_troca_bloqueada"] = {"motivo": "antigo"}
        self.assertTrue(s["trocar_condominio_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertNotIn("_op_troca_bloqueada", s["ss"])

    def test_salvar_rascunho_real_informa_falha_da_gravacao_local(self):
        s = _sessao()
        salvar = s["salvar_rascunho_operador_real"]
        self.assertTrue(salvar("Condomínio Alfa", {"obs": "x"}))
        self.assertTrue((s["pasta"] / slugify_nome("Condomínio Alfa") / "_rascunho_operador.json").exists())
        (s["pasta"] / slugify_nome("Residencial Beta")).write_text("arquivo no lugar da pasta")  # mkdir falha
        self.assertFalse(salvar("Residencial Beta", {"obs": "x"}))

    def test_chave_das_fotos_comeca_com_o_nome_original(self):
        s = _sessao()
        self.assertEqual(s["_chave_fotos_operador"]("extras"), "op_fotos_extras")
        self.assertEqual(s["chave_fotos_rt"](), "rel_fotos_upload")


class TestFluxoRT(unittest.TestCase):
    def test_limpar_rt_remove_dados_do_condominio_anterior(self):
        s = _sessao()
        ss = s["ss"]
        ss.update({"rel_endereco_condominio": "Rua A, 1", "rel_cnpj_condominio": "00", "rel_representante": "Síndico A",
                   "rel_analise_ph_0": "7.4", "rel_analise_cya_0": "0", "rel_parecer_complemento_rt": "Parecer A",
                   "rel_foto_legenda_cmaq": "Casa A", "rel_art_numero": "ART-A", "rel_art_modo_rt2025": "Informar",
                   "rel_art_status": "Emitida", "rel_mes_referencia": "10", "rel_ano_referencia": "2026"})
        chave = s["chave_fotos_rt"]()
        s["limpar_formulario_rt_operador"]()
        for k in ("rel_endereco_condominio", "rel_cnpj_condominio", "rel_representante", "rel_analise_ph_0",
                  "rel_analise_cya_0", "rel_parecer_complemento_rt", "rel_foto_legenda_cmaq", "rel_art_numero"):
            self.assertEqual(ss[k], "", k)
        self.assertNotIn("rel_art_modo_rt2025", ss)
        self.assertNotIn("rel_art_status", ss)
        self.assertEqual((ss["rel_mes_referencia"], ss["rel_ano_referencia"]), ("10", "2026"))
        self.assertNotEqual(s["chave_fotos_rt"](), chave)
        self.assertFalse(s["chave_fotos_rt"]().startswith("rel_"))  # fora do rascunho RT (só rel_/csr_)


class TestTrocaFluxoRT(unittest.TestCase):
    def _rt(self, ss, endereco="Rua A, 1", ph="7.4"):
        ss.update({"rel_mes_referencia": "10", "rel_ano_referencia": "2026", "rel_endereco_condominio": endereco,
                   "rel_analise_ph_0": ph, "rel_analise_cya_0": "0", "rel_parecer_complemento_rt": "Parecer A",
                   "rel_art_numero": "ART-A", "rel_art_status": "Emitida"})

    def test_troca_guarda_rascunho_do_anterior_e_recupera_ao_voltar(self):
        s = _sessao()
        ss = s["ss"]
        self._rt(ss)
        self.assertTrue(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertEqual((ss["rel_endereco_condominio"], ss["rel_analise_ph_0"], ss["rel_parecer_complemento_rt"]), ("", "", ""))
        self.assertNotIn("_rt_rascunho_restaurado", ss)
        ss.update({"rel_endereco_condominio": "Rua B, 2", "rel_analise_ph_0": "7.0"})
        self.assertTrue(s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa"))
        self.assertEqual(ss["_rt_rascunho_restaurado"], "Condomínio Alfa")
        self.assertEqual((ss["rel_endereco_condominio"], ss["rel_analise_ph_0"], ss["rel_analise_cya_0"],
                          ss["rel_parecer_complemento_rt"], ss["rel_art_numero"], ss["rel_art_status"]),
                         ("Rua A, 1", "7.4", "0", "Parecer A", "ART-A", "Emitida"))
        self.assertFalse(s["_rascunho_rt_operador_path"]("Condomínio Alfa").exists())  # não volta uma 2ª vez

    def test_dados_apagados_depois_de_recuperar_nao_reaparecem(self):
        s = _sessao()
        ss = s["ss"]
        self._rt(ss)
        s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")  # recupera
        s["limpar_formulario_rt_operador"]()  # usuário apaga os dados de Alfa
        s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")
        self.assertEqual((ss["rel_endereco_condominio"], ss["rel_parecer_complemento_rt"]), ("", ""))

    def test_rascunho_de_outro_mes_nao_volta(self):
        s = _sessao()
        ss = s["ss"]
        self._rt(ss)
        ss["rel_mes_referencia"] = "09"
        s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        ss["rel_mes_referencia"] = "10"
        s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")
        self.assertEqual((ss["rel_endereco_condominio"], ss["rel_analise_ph_0"]), ("", ""))  # setembro não entra em outubro
        self.assertNotIn("_rt_rascunho_restaurado", ss)

    def test_troca_cancelada_com_fotos_anexadas(self):
        s = _sessao()
        ss = s["ss"]
        self._rt(ss)
        ss[s["chave_fotos_rt"]()] = ["foto_ficticia.png"]
        antes = dict(ss)
        self.assertFalse(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertIn("fotos", ss.pop("_rt_troca_bloqueada")["motivo"])
        self.assertEqual(ss, antes)

    def test_troca_cancelada_se_o_rascunho_nao_pode_ser_gravado(self):
        s = _sessao()
        ss = s["ss"]
        self._rt(ss)
        (s["pasta"] / slugify_nome("Condomínio Alfa")).write_text("arquivo no lugar da pasta")
        antes = dict(ss)
        self.assertFalse(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertTrue(ss.pop("_rt_troca_bloqueada")["motivo"])
        ss.pop("_rt_op_sessao", None)  # identificador da sessão para o nome do arquivo
        self.assertEqual(ss, antes)

    def test_troca_sem_dados_nao_grava_rascunho(self):
        s = _sessao()
        s["ss"].update({"rel_mes_referencia": "10", "rel_ano_referencia": "2026"})
        self.assertTrue(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertFalse(s["_rascunho_rt_operador_path"]("Condomínio Alfa").exists())


# Formulário RT completo de um condomínio fictício: análises, dosagens, observações,
# parecer, identificação, ART e seleções do relatório (EPI, NBR, tipo de atendimento).
RT_COMPLETO_A = {
    "rel_analise_data_0": "06/10/2026", "rel_analise_ph_0": "7,4", "rel_analise_cl_0": "1,5", "rel_analise_cya_0": "0",
    "rel_analise_data_1": "13/10/2026", "rel_analise_ph_1": "7,2", "rel_analise_turbidez_1": "0,30",
    "rel_dos_produto_0": "Cloro granulado", "rel_dos_qtd_0": "2", "rel_dos_un_0": "kg",
    "rel_dos_finalidade_0": "Manutenção", "rel_dos_lote_0": "L-01",
    "rel_observacoes_gerais": "Observação geral A", "rel_obs_0": "Observação item A",
    "rel_parecer_complemento_rt": "Parecer A", "rel_diagnostico": "Diagnóstico A",
    "rel_endereco_condominio": "Rua A, 1", "rel_cnpj_condominio": "00.000.000/0001-00",
    "rel_representante": "Síndico A", "rel_cpf_cnpj_representante": "000.000.000-00", "rel_cep": "30000-000",
    "rel_art_numero": "ART-A", "rel_art_status": "Emitida",
    "rel_foto_legenda_cmaq": "Casa de máquinas A", "rel_tipo_atendimento": "Visita técnica avulsa",
    "rel_epi_luvas_status": "Não", "rel_nbr11238_chuveiro": "Não conforme", "rel_salvar_alteracoes_cadastro": True,
}
PERIODO = {"rel_mes_referencia": "10", "rel_ano_referencia": "2026", "rel_data_emissao": "26/10/2026",
           "rel_nome_condominio": "Condomínio Alfa"}


class TestRascunhoRTCompleto(unittest.TestCase):
    def test_todos_os_campos_voltam_intactos_e_nada_passa_para_o_outro(self):
        s = _sessao()
        ss = s["ss"]
        ss.update(PERIODO, **RT_COMPLETO_A)
        self.assertTrue(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        for k in RT_COMPLETO_A:  # Beta não herda nada de Alfa
            self.assertIn(ss.get(k, ""), ("", None), k)
        self.assertEqual({k: ss[k] for k in PERIODO}, PERIODO)  # período e data de emissão ficam
        self.assertTrue(s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa"))
        self.assertEqual({k: ss.get(k) for k in RT_COMPLETO_A}, RT_COMPLETO_A)

    def test_cobre_todo_campo_rel_editavel_do_fluxo_rt(self):
        """Todo widget rel_ do fluxo RT do operador é guardado ou é refeito pelo próprio fluxo."""
        alvo = next(n for n in ast.walk(TREE) if isinstance(n, ast.If) and ast.unparse(n.test) == "_op_is_rt"
                    and any(isinstance(c, ast.Call) and getattr(c.func, "attr", "") == "file_uploader" for c in ast.walk(n)))
        chaves = set()
        for n in ast.walk(alvo):
            if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute):
                for kw in n.keywords:
                    if kw.arg == "key":
                        txt = ast.unparse(kw.value)
                        m = re.match(r"""^f?['"](rel_[a-z0-9_]+?)(_\{i\})?['"]$""", txt)
                        if m:
                            chaves.add(m.group(1) + ("_0" if m.group(2) else ""))
        self.assertGreaterEqual(len(chaves), 15)
        s = _sessao()
        ss = s["ss"]
        ss.update({k: "valor" for k in chaves})
        guardadas = set(s["_campos_rt_do_condominio"]())
        refeitas = {"rel_nome_condominio", "rel_data_emissao", "rel_art_modo_rt2025", "rel_art_numero_rt2025"}
        self.assertEqual(chaves - guardadas, refeitas & chaves)

    def test_campo_que_nao_pode_ser_guardado_bloqueia_a_troca(self):
        s = _sessao()
        ss = s["ss"]
        ss.update(PERIODO, **RT_COMPLETO_A)
        ss["rel_vigencia_art"] = datetime(2026, 10, 26)  # valor não serializável (simulado)
        antes = dict(ss)
        self.assertFalse(s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta"))
        self.assertIn("rel_vigencia_art", ss.pop("_rt_troca_bloqueada")["motivo"])
        self.assertEqual(ss, antes)

    def test_restauracao_nao_sobrescreve_edicoes_posteriores(self):
        s = _sessao()
        ss = s["ss"]
        ss.update(PERIODO, **RT_COMPLETO_A)
        s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")  # restaura uma vez
        ss.update(rel_analise_ph_0="7,0", rel_dos_qtd_0="3", rel_observacoes_gerais="Editada depois")
        self.assertFalse(s["_rascunho_rt_operador_path"]("Condomínio Alfa").exists())  # não há o que reaplicar
        s["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        s["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")
        self.assertEqual((ss["rel_analise_ph_0"], ss["rel_dos_qtd_0"], ss["rel_observacoes_gerais"]),
                         ("7,0", "3", "Editada depois"))

    def test_rascunho_de_uma_sessao_nao_aparece_em_outra(self):
        s1 = _sessao()
        s1["ss"].update(PERIODO, **RT_COMPLETO_A)
        s1["trocar_condominio_rt_operador"]("Condomínio Alfa", "Residencial Beta")
        s2 = _sessao(pasta=s1["pasta"])
        s2["ss"].update(PERIODO)
        s2["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")
        self.assertEqual(s2["ss"].get("rel_dos_produto_0", ""), "")
        self.assertNotIn("_rt_rascunho_restaurado", s2["ss"])
        s1["trocar_condominio_rt_operador"]("Residencial Beta", "Condomínio Alfa")  # a sessão 1 continua recuperando
        self.assertEqual(s1["ss"]["rel_dos_produto_0"], "Cloro granulado")


class TestLigacoesNoApp(unittest.TestCase):
    def test_uploads_usam_as_chaves_dinamicas(self):
        src = APP.read_text(encoding="utf-8")
        for cat in ("antes", "depois", "cmaq", "extras"):
            self.assertIn(f'key=_chave_fotos_operador("{cat}")', src)
            self.assertNotIn(f'key="op_fotos_{cat}"', src)
        self.assertNotIn('key="rel_fotos_upload"', src)
        self.assertEqual(src.count("key=chave_fotos_rt()"), 2)
        self.assertNotIn('st.session_state.get("rel_fotos_upload")', src)

    def test_trocas_canceladas_mantem_o_anterior_na_tela(self):
        src = APP.read_text(encoding="utf-8")
        self.assertIn("elif not trocar_condominio_operador(_cond_form_anterior, op_nome_cond):", src)
        self.assertIn("op_nome_cond = _cond_form_anterior", src)
        self.assertIn("if not trocar_condominio_rt_operador(_cond_rt_anterior, _sel_cond_rt):", src)
        self.assertIn("_sel_cond_rt = _cond_rt_anterior", src)
        # a seleção volta antes de os widgets serem criados
        self.assertLess(src.index('_op_rev = st.session_state.pop("_op_cond_reverter", None)'),
                        src.index('key="op_novo_cond"'))
        self.assertLess(src.index('_rt_rev = st.session_state.pop("_rt_cond_reverter", None)'),
                        src.index('key="rt_op_sel_cond"'))


if __name__ == "__main__":
    unittest.main()
