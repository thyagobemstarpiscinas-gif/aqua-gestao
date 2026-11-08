"""Regressões dos dois bloqueios finais da homologação do PR #14 (dados fictícios).

1. Restauração/troca do rascunho RT bloqueada por falha transitória da planilha:
   "Troca bloqueada: não foi possível preservar o rascunho…" mesmo com o
   conteúdo já gravado no disco local. A proteção contra perda de dados
   continua: falha da gravação LOCAL ainda bloqueia.
2. PDF mensal Bem Star sem a fotografia da V00005: no Streamlit Cloud a pasta
   fotos_campo não existe após reinício; o payload do Sheets traz só nomes e
   IDs do Drive, que não eram consultados.

Google falso (escopo drive.file), planilha em memória, rede bloqueada.
"""
import base64
import json
import shutil
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT))

from test_homologacao_pr14 import EPSILON, _Homologacao, _foto, _imagens_pdf  # noqa: E402
from test_rascunho_rt import Ambiente  # noqa: E402
from utils import bem_star_mensal as bsm  # noqa: E402
from utils import pdf_bem_star_mensal as pdf_bs  # noqa: E402
from utils.formatacao import slugify_nome  # noqa: E402


# ── 1. Troca/restauração RT com a planilha indisponível ────────────────────
class TrocaRTComPlanilhaIndisponivel(unittest.TestCase):
    def setUp(self):
        self.amb = Ambiente()
        self.s = self.amb.sessao()
        self.s["_rascunho_usa_sheets"] = lambda: True
        self.ss = self.s["ss"]

    def _registros(self):
        gestor = self.s["_relatorio_rt_gestor_sync"]()
        return [json.loads(p.read_text(encoding="utf-8")) for p in gestor.pasta.glob("*.json")]

    def _alfa_novembro(self):
        self.ss.update(rel_nome_condominio="Condomínio Alfa Fictício", rel_mes_referencia="11",
                       rel_ano_referencia="2026", rel_origem_dados="aqua_gestao|id:C1|11/2026",
                       rel_diagnostico="ALFA_NOVEMBRO_RESTAURADO")

    def test_falha_da_planilha_nao_bloqueia_troca_e_conteudo_fica_no_disco(self):
        self._alfa_novembro()
        self.amb.criar_aba().falhar = True  # cota/rede do Sheets indisponível
        self.assertTrue(self.s["_guardar_rascunho_rt_antes_da_troca"]())
        self.assertNotIn("_rt_contexto_erro", self.ss)
        guardados = [r for r in self._registros()
                     if r["dados"]["campos"].get("rel_diagnostico") == "ALFA_NOVEMBRO_RESTAURADO"]
        self.assertTrue(guardados and all(r["estado"] == "pendente" for r in guardados))
        # Planilha de volta: a pendência é enviada sem nova edição.
        self.amb.aba().falhar = False
        self.ss["_rt_rasc_proxima_tentativa"] = 0
        self.assertTrue(self.s["_relatorio_rt_sincronizar_pendentes"]())
        self.assertTrue(any("ALFA_NOVEMBRO_RESTAURADO" in "".join(map(str, row)) for row in self.amb.linhas()))

    def test_pendencia_divergente_vira_copia_nova_sem_apagar_nenhuma(self):
        self._alfa_novembro()
        self.amb.criar_aba().falhar = True
        self.assertFalse(self.s["_relatorio_rt_salvar_rascunho"]())  # pendente: ALFA_NOVEMBRO_RESTAURADO
        self.ss["rel_diagnostico"] = "EDICAO_POSTERIOR"
        self.assertTrue(self.s["_guardar_rascunho_rt_antes_da_troca"]())
        textos = sorted(r["dados"]["campos"]["rel_diagnostico"] for r in self._registros())
        self.assertEqual(textos, ["ALFA_NOVEMBRO_RESTAURADO", "EDICAO_POSTERIOR"])

    def test_restauracao_explicita_tambem_prossegue(self):
        self._alfa_novembro()
        self.amb.criar_aba().falhar = True
        self.assertTrue(self.s["_guardar_rascunho_rt_antes_da_troca"](restauracao_explicita=True))

    def test_falha_da_gravacao_local_continua_bloqueando(self):
        self._alfa_novembro()
        self.amb.criar_aba().falhar = True
        with patch("utils.rascunho_sync.os.replace", side_effect=OSError("DISCO_FICTICIO")):
            self.assertFalse(self.s["_guardar_rascunho_rt_antes_da_troca"]())
        self.assertIn("Troca bloqueada", self.ss["_rt_contexto_erro"])

    def test_tela_sem_conteudo_nao_cria_rascunho(self):
        self.amb.criar_aba().falhar = True
        self.assertTrue(self.s["_guardar_rascunho_rt_antes_da_troca"]())
        self.assertEqual(self._registros(), [])


# ── 2. Foto da V00005 no PDF mensal Bem Star ───────────────────────────────
class FotoNoPdfMensalBemStar(_Homologacao):
    def _v00005_do_sheets(self):
        self._gravar_v00005(com_id_local=False)
        # Streamlit Cloud reiniciado: nenhuma cópia local da visita ou das fotos.
        shutil.rmtree(self.gerados / slugify_nome(EPSILON))
        visitas = self.app["situacao_visitas_condominio"](EPSILON)["visitas"]
        self.assertEqual([v["id_visita"] for v in visitas], ["V00005"])
        self.assertEqual(visitas[0]["fotos_antes_ids"], [self.foto_id])
        return visitas

    def _gerar_mensal(self, lancamentos, fotos):
        cliente = next(c for c in self.app["sheets_listar_clientes_completo"]() if c["nome"] == EPSILON)
        ctx = {"cliente": EPSILON, "mes": "10", "ano": "2026", "dados_cliente": cliente,
               "lancamentos": bsm.unir_agenda(bsm.deduplicar_lancamentos(lancamentos), [], ["Piscina Adulto"]),
               "piscinas": ["Piscina Adulto"], "piscinas_cadastradas": ["Piscina Adulto"],
               "tipos_piscinas": {"Piscina Adulto": "Adulto"}, "fotos": fotos, "fotos_metadados": {},
               "operador": "", "obs_geral": "", "parecer_tecnico": ""}
        saida = self.gerados / "mensal_epsilon.pdf"
        ok, erro = pdf_bs.gerar_pdf(saida, ctx, cnpj="00.000.000/0000-00", responsavel="RESPONSÁVEL FICTÍCIO",
                                    logo_path=ROOT / "bem_star_logo.png")
        self.assertTrue(ok, erro)
        return saida.read_bytes()

    def test_foto_do_drive_entra_no_pdf_mensal(self):
        visitas = self._v00005_do_sheets()
        fotos = self.app["_coletar_fotos_bem_star_preview"](EPSILON, visitas)
        self.assertEqual(len(fotos), 1)
        meta = self.app["st"].session_state["_bs_fotos_metadados_visitas"][str(fotos[0])]
        self.assertEqual((meta["data"], meta["legenda"]), ("27/10/2026", "Visita V00005 — Antes do tratamento"))
        pdf = self._gerar_mensal(visitas, fotos)
        self.assertIn((640, 480), _imagens_pdf(pdf))
        with fitz.open(stream=pdf, filetype="pdf") as doc:
            self.assertIn("Registro fotográfico", "\n".join(p.get_text() for p in doc))
        self.assertEqual(self.app["st"].session_state.get("_bs_fotos_avisos"), [])
        # V00005 e a foto continuam únicas e intactas.
        linhas = self.planilha.abas["🔬 Visitas"].linhas
        self.assertEqual(sum(1 for r in linhas if len(r) > 1 and r[1] == "V00005"), 1)
        self.assertEqual(self.app["drive_baixar_foto"](self.foto_id), self.foto)

    def test_foto_persiste_ate_o_fim_da_montagem(self):
        visitas = self._v00005_do_sheets()
        fotos = self.app["_coletar_fotos_bem_star_preview"](EPSILON, visitas)
        self.assertTrue(all(Path(f).is_file() for f in fotos))
        self.assertIn(self.gerados, Path(fotos[0]).parents)  # pasta do app, não temporária
        self._gerar_mensal(visitas, fotos)
        self.assertTrue(all(Path(f).is_file() for f in fotos))

    def test_copia_local_com_arquivo_nao_duplica_foto(self):
        lanc, local = self._gravar_v00005(com_id_local=False)
        pasta_campo = self.gerados / slugify_nome(EPSILON) / "fotos_campo"
        pasta_campo.mkdir(parents=True, exist_ok=True)
        (pasta_campo / local["fotos_antes"][0]).write_bytes(self.foto)
        visitas = bsm.deduplicar_lancamentos([local] + self.app["situacao_visitas_condominio"](EPSILON)["visitas"])
        fotos = self.app["deduplicar_fotos"](self.app["_coletar_fotos_bem_star_preview"](EPSILON, visitas))
        self.assertEqual(len(fotos), 1)

    def test_sem_drive_usa_base64_da_copia_local(self):
        _, local = self._gravar_v00005(com_id_local=False)
        local = dict(local, fotos_antes_ids=["ID_INACESSIVEL"])
        fotos = self.app["_coletar_fotos_bem_star_preview"](EPSILON, [local])
        self.assertEqual(len(fotos), 1)

    def test_foto_irrecuperavel_gera_aviso_em_vez_de_silencio(self):
        visita = {"id_visita": "V00099", "data": "28/10/2026", "fotos_antes": ["x.jpg"],
                  "fotos_antes_ids": ["ID_INACESSIVEL"]}
        self.assertEqual(self.app["_coletar_fotos_bem_star_preview"](EPSILON, [visita]), [])
        avisos = self.app["st"].session_state.get("_bs_fotos_avisos")
        self.assertTrue(avisos and "V00099" in avisos[0], avisos)


if __name__ == "__main__":
    unittest.main()
