"""Pendências: disco temporário, planilha em memória e relógio controlado."""
import copy
import json
import unittest
from unittest.mock import patch

from test_rascunho_rt import Ambiente
from utils.rascunho_sync import ConflitoRascunho, hash_conteudo


class TestPendenciaDuravel(unittest.TestCase):
    def setUp(self):
        self.amb = Ambiente()
        self.s = self.sessao()

    def sessao(self):
        s = self.amb.sessao()
        s["_rascunho_usa_sheets"] = lambda: True
        return s

    def salvar(self, texto):
        self.s["ss"]["rel_diagnostico"] = texto
        return self.s["_relatorio_rt_salvar_rascunho"]()

    def pendencia(self):
        self.assertTrue(self.salvar("VERSAO_REMOTA"))
        chave = self.s["ss"]["_rt_rasc_chave"]
        self.amb.aba().falhar = True
        self.assertFalse(self.salvar("EDICAO_LOCAL_PENDENTE"))
        self.amb.aba().falhar = False
        return chave, self.s["_relatorio_rt_gestor_sync"]()

    def test_reabertura_reconstroi_controles_sem_alterar_campos(self):
        chave, gestor = self.pendencia()
        nova = self.sessao()
        nova["ss"]["rel_diagnostico"] = "CAMPO_ATIVO"
        antes = copy.deepcopy(nova["ss"])
        pendentes = nova["_relatorio_rt_gestor_sync"]().reconstruir()
        self.assertEqual(pendentes, [chave])
        for k, v in antes.items():
            self.assertEqual(nova["ss"][k], v)
        self.assertEqual(gestor.ler(chave)["dados"]["campos"]["rel_diagnostico"], "EDICAO_LOCAL_PENDENTE")

    def test_empate_duas_variantes_e_retentativa_nao_sobrescreve(self):
        chave, gestor = self.pendencia()
        remota = copy.deepcopy(self.amb.linhas(chave))
        local = gestor.ler(chave)["dados"]
        local["_salvo_iso"] = self.s["_relatorio_rt_desserializar"](remota[0])["_salvo_iso"]
        r = gestor.ler(chave)
        r["dados"] = local
        from utils.rascunho_sync import escrever_atomico
        escrever_atomico(gestor.caminho(chave), r)
        lista = self.s["_relatorio_rt_listar_rascunhos"]()["versoes"]
        variantes = [v for v in lista if v["chave"] == chave]
        self.assertEqual({v["origem"] for v in variantes}, {"local", "planilha"})
        self.assertTrue(all(v["conflito"] for v in variantes))
        self.assertEqual(lista[0]["origem"], "local")
        self.assertEqual(self.s["_relatorio_rt_carregar_rascunho"](), {})
        for s in (self.s, self.sessao()):
            s["ss"]["_rt_rasc_proxima_tentativa"] = 0
            self.assertFalse(s["_relatorio_rt_sincronizar_pendentes"]())
            self.assertEqual(self.amb.linhas(chave), remota)
            self.assertIn(chave, s["ss"]["_rt_rasc_sync_pendente"])

    def test_todos_caminhos_de_envio_bloqueiam_reenvio_divergente(self):
        chave, gestor = self.pendencia()
        remota = copy.deepcopy(self.amb.linhas(chave))
        with self.assertRaises(ConflitoRascunho):
            self.s["_relatorio_rt_enviar_planilha"](self.amb.aba(), chave, gestor.ler(chave)["dados"])
        self.assertFalse(self.salvar("OUTRA_EDICAO_NAO_DEVE_APAGAR_PENDENTE"))
        self.assertEqual(self.amb.linhas(chave), remota)
        self.assertEqual(gestor.ler(chave)["dados"]["campos"]["rel_diagnostico"], "EDICAO_LOCAL_PENDENTE")

    def test_restaura_ambas_variantes_em_novas_chaves(self):
        chave, gestor = self.pendencia()
        remota = copy.deepcopy(self.amb.linhas(chave))
        variantes = [v for v in self.s["_relatorio_rt_listar_rascunhos"]()["versoes"] if v["chave"] == chave]
        for v in variantes:
            nova = self.sessao()
            nova["ss"]["_rt_rasc_pendente"] = {"dados": v["dados"]}
            nova["_relatorio_rt_aplicar_pendente"]()
            self.assertEqual(nova["ss"]["rel_diagnostico"], v["dados"]["campos"]["rel_diagnostico"])
            self.assertNotEqual(nova["ss"]["_rt_rasc_chave"], chave)
            self.assertEqual(self.amb.linhas(chave), remota)
        self.assertEqual(gestor.ler(chave)["estado"], "conflito")

    def test_envio_confirmado_com_resposta_perdida_nao_duplica(self):
        chave, gestor = self.pendencia()
        # Simula a cópia já entregue, cuja resposta ao cliente não chegou.
        self.s["_relatorio_rt_gravar_linha"](self.amb.aba(), chave, gestor.ler(chave)["dados"])
        ops = self.amb.aba().updates + self.amb.aba().appends
        self.s["ss"]["_rt_rasc_proxima_tentativa"] = 0
        self.assertTrue(self.s["_relatorio_rt_sincronizar_pendentes"]())
        self.assertEqual(self.amb.aba().updates + self.amb.aba().appends, ops)
        self.assertEqual(gestor.ler(chave)["estado"], "confirmado")
        self.assertNotIn("_rt_rasc_sync_pendente", self.s["ss"])

    def test_ausencia_confirmada_permite_incluir_uma_vez(self):
        self.amb.criar_aba().falhar = True
        self.assertFalse(self.salvar("PRIMEIRO_ENVIO"))
        chave = self.s["ss"]["_rt_rasc_chave"]
        self.amb.aba().falhar = False
        nova = self.sessao()
        self.assertTrue(nova["_relatorio_rt_sincronizar_pendentes"]())
        self.assertEqual(len(self.amb.linhas(chave)), 1)
        self.assertTrue(nova["_relatorio_rt_sincronizar_pendentes"]())
        self.assertEqual(len(self.amb.linhas(chave)), 1)

    def test_consulta_falha_e_json_ilegivel_nao_significam_ausencia(self):
        chave, gestor = self.pendencia()
        for caso in ("consulta", "ilegivel", "indisponivel"):
            with self.subTest(caso=caso):
                aba = self.amb.aba()
                before = copy.deepcopy(aba.linhas)
                if caso == "consulta":
                    with patch.object(aba, "get_all_values", side_effect=ConnectionError("FALHA_FICTICIA")):
                        with self.assertRaises(ConnectionError):
                            gestor.enviar(aba, chave, gestor.ler(chave)["dados"])
                elif caso == "ilegivel":
                    aba.linhas[1][2] = "{"
                    with self.assertRaisesRegex(ValueError, "ilegível"):
                        gestor.enviar(aba, chave, gestor.ler(chave)["dados"])
                    self.assertEqual(aba.linhas[1][2], "{")
                    aba.linhas = before
                else:
                    with self.assertRaises(ConnectionError):
                        gestor.enviar(None, chave, gestor.ler(chave)["dados"])
                self.assertEqual(aba.linhas, before)
                self.assertNotEqual(gestor.ler(chave)["estado"], "confirmado")

    def test_falha_local_preserva_copia_anterior_e_nao_envia(self):
        self.assertTrue(self.salvar("ANTERIOR_VALIDA"))
        chave = self.s["ss"]["_rt_rasc_chave"]
        gestor = self.s["_relatorio_rt_gestor_sync"]()
        anterior = gestor.caminho(chave).read_bytes()
        remota = copy.deepcopy(self.amb.linhas(chave))
        with patch("utils.rascunho_sync.os.replace", side_effect=OSError("DISCO_FICTICIO")):
            self.assertFalse(self.salvar("NAO_PERSISTIDA"))
        self.assertEqual(gestor.caminho(chave).read_bytes(), anterior)
        self.assertEqual(self.amb.linhas(chave), remota)
        self.assertIn("DISCO_FICTICIO", self.s["ss"]["_rt_rasc_aviso"])

    def test_descarte_remoto_inclusive_json_ilegivel_sobrevive_offline(self):
        chave, gestor = self.pendencia()
        aba = self.amb.aba()
        aba.linhas[1][2] = "{"
        aba.linhas[1][3] = "descartado em teste"
        self.assertFalse(any(v["chave"] == chave for v in self.s["_relatorio_rt_listar_rascunhos"]()["versoes"]))
        self.assertEqual(gestor.ler(chave)["estado"], "descartado")
        nova = self.sessao()
        nova["_relatorio_rt_aba"] = lambda **_: None
        self.assertTrue(nova["_relatorio_rt_sincronizar_pendentes"]())
        self.assertFalse(any(v["chave"] == chave for v in nova["_relatorio_rt_listar_rascunhos"]()["versoes"]))

    def test_descarte_local_e_respeitado_antes_de_qualquer_consulta(self):
        chave, gestor = self.pendencia()
        local = self.s["_relatorio_rt_arquivo_local"](chave)
        dados = json.loads(local.read_text(encoding="utf-8"))
        dados["_descartado"] = "descartado localmente"
        local.write_text(json.dumps(dados), encoding="utf-8")
        with patch.object(self.amb.aba(), "get_all_values", side_effect=AssertionError("Não consultar remoto")):
            self.assertFalse(gestor.enviar(self.amb.aba(), chave, gestor.ler(chave)["dados"]))
        self.assertEqual(gestor.ler(chave)["estado"], "descartado")

    def test_autosaves_normais_mantem_uma_linha(self):
        for i in range(4):
            self.assertTrue(self.salvar(f"EDICAO_NORMAL_{i}"))
        self.assertEqual(len(self.amb.linhas()), 1)

    def test_restauracao_explicita_preserva_edicao_ativa_diferente_da_pendencia(self):
        chave, gestor = self.pendencia()
        remota = copy.deepcopy(self.amb.linhas(chave))
        self.s["ss"]["rel_diagnostico"] = "ATIVA_AINDA_NAO_SALVA"
        self.assertTrue(self.s["_guardar_rascunho_rt_antes_da_troca"](restauracao_explicita=True))
        self.assertEqual(gestor.ler(chave)["dados"]["campos"]["rel_diagnostico"], "EDICAO_LOCAL_PENDENTE")
        outras = [json.loads(p.read_text(encoding="utf-8")) for p in gestor.pasta.glob("*.json") if p != gestor.caminho(chave)]
        self.assertTrue(any(r["dados"]["campos"]["rel_diagnostico"] == "ATIVA_AINDA_NAO_SALVA" and r["estado"] == "pendente" for r in outras))
        self.assertEqual(self.s["ss"]["rel_diagnostico"], "ATIVA_AINDA_NAO_SALVA")
        self.assertEqual(self.amb.linhas(chave), remota)

    def test_ausencia_com_insercao_concorrente_nao_atualiza_outra_copia(self):
        self.amb.criar_aba().falhar = True
        self.assertFalse(self.salvar("LOCAL_NOVA"))
        chave = self.s["ss"]["_rt_rasc_chave"]
        gestor = self.s["_relatorio_rt_gestor_sync"]()
        outra = copy.deepcopy(gestor.ler(chave)["dados"])
        outra["campos"]["rel_diagnostico"] = "REMOTA_CONCORRENTE"
        self.amb.aba().falhar = False
        self.amb.aba().antes_de_gravar = lambda: self.s["_relatorio_rt_gravar_linha"](self.amb.aba(), chave, outra)
        with self.assertRaises(ConflitoRascunho):
            gestor.enviar(self.amb.aba(), chave, gestor.ler(chave)["dados"])
        rows = self.amb.linhas(chave)
        self.assertEqual(len(rows), 2)
        self.assertEqual({self.s["_relatorio_rt_desserializar"](r)["campos"]["rel_diagnostico"] for r in rows}, {"LOCAL_NOVA", "REMOTA_CONCORRENTE"})
        self.assertEqual(self.amb.aba().updates, 0)

    def test_falha_local_no_descarte_informa_e_preserva_conteudo(self):
        chave, gestor = self.pendencia()
        anterior = gestor.caminho(chave).read_bytes()
        with patch("utils.rascunho_sync.os.replace", side_effect=OSError("DISCO_FICTICIO")):
            self.assertFalse(self.s["_relatorio_rt_descartar"](chave))
        self.assertEqual(gestor.caminho(chave).read_bytes(), anterior)
        self.assertIn("DISCO_FICTICIO", self.s["ss"]["_rt_rasc_erro_descarte"])
        self.assertTrue(self.amb.linhas(chave)[0][3].startswith("descartado"))

    def test_insercao_remota_ilegivel_ou_estrangeira_nao_confirma_envio(self):
        chave, gestor = self.pendencia()
        dados = copy.deepcopy(gestor.ler(chave)["dados"])
        dados["empresa_ativa"] = "bem_star"
        dados["campos"].pop("rel_origem_dados", None)
        self.s["_relatorio_rt_gravar_linha"](self.amb.aba(), chave, dados)
        antes = copy.deepcopy(self.amb.linhas(chave))
        with self.assertRaises(ConflitoRascunho):
            gestor.enviar(self.amb.aba(), chave, gestor.ler(chave)["dados"])
        self.assertEqual(self.amb.linhas(chave), antes)

    def test_isolamento_por_empresa_e_legado_sem_metadados(self):
        chave, gestor = self.pendencia()
        outra = self.sessao()
        outra["ss"]["empresa_ativa"] = "bem_star"
        self.assertEqual(outra["_relatorio_rt_gestor_sync"]().reconstruir(), [])
        self.assertEqual(outra["_relatorio_rt_listar_rascunhos"]()["versoes"], [])
        remota = copy.deepcopy(self.amb.linhas(chave))
        self.assertFalse(outra["_relatorio_rt_descartar"](chave))
        self.assertEqual(self.amb.linhas(chave), remota)
        legado = {"campos": {"rel_diagnostico": "LEGADO"}, "salvo_em": "25/10/2026 12:00:00"}
        p = self.amb.dir / "_rascunho_relatorio_rt_em_andamento.json"
        p.write_text(json.dumps(legado), encoding="utf-8")
        lista = self.s["_relatorio_rt_listar_rascunhos"]()["versoes"]
        self.assertTrue(any(v["chave"] == "thyago" and v["dados"] == legado for v in lista))
        self.assertEqual(hash_conteudo(gestor.ler(chave)["dados"]), gestor.ler(chave)["hash"])
