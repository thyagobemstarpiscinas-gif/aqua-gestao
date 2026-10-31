"""Regressões do PDF Bem Star, com dados fictícios e sem Streamlit ou rede."""
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

import fitz
from test_fotos_isolamento import _carregar


def contexto():
    return {
        "cliente": "Tríad Fictício - TESTE", "mes": "09", "ano": "2026",
        "operador": "Operador Fictício - TESTE", "fotos": [],
        "obs_geral": "SIMULAÇÃO - DADOS FICTÍCIOS. SEM VALIDADE OPERACIONAL.",
        "parecer_tecnico": "\n\n".join(
            f"BLOCO_{i:02d}: " + "Análise fictícia para verificar a preservação integral do texto. " * 4
            for i in range(1, 13)
        ),
        "lancamentos": [
            {"data": f"{d:02d}/09/2026", "ph": "7.5", "cloro_livre": "2.1",
             "cloro_total": "2.3", "alcalinidade": "110", "dureza": "250", "cianurico": "40"}
            for d in [2, 4, 7, 9, 11, 14, 16, 18, 21, 23, 25, 28, 30]
        ],
    }


def renderizar(ctx, arquivo):
    ns = _carregar(["CNPJ_BEM_STAR", "RESPONSAVEL_TÉCNICO", "_bs_eh_segunda",
                    "gerar_pdf_relatorio_mensal_bem_star_modelo"])
    ns["st"].session_state["bs_rel_dias_semana"] = ["Segunda-feira", "Quarta-feira", "Sexta-feira"]
    # Simula o relógio ingênuo do servidor no dia seguinte.
    ns["hoje_br"] = lambda: "31/10/2026"
    return ns["gerar_pdf_relatorio_mensal_bem_star_modelo"](arquivo, ctx)


class TestBemStarPaginacao(unittest.TestCase):
    def gerar(self, ctx):
        pasta = tempfile.TemporaryDirectory()
        self.addCleanup(pasta.cleanup)
        arquivo = Path(pasta.name) / "teste.pdf"
        with patch("utils.relogio.hoje_brasilia", return_value="30/10/2026", create=True):
            ok, erro = renderizar(ctx, arquivo)
        self.assertTrue(ok, erro)
        doc = fitz.open(arquivo)
        paginas = [p.get_text() for p in doc]
        doc.close()
        return paginas

    def test_parecer_preserva_todos_os_blocos(self):
        texto = "\n".join(self.gerar(contexto()))
        for i in range(1, 13):
            self.assertIn(f"BLOCO_{i:02d}", texto)

    def test_assinatura_permanece_com_encerramento(self):
        paginas = self.gerar(contexto())
        assinatura = next(p for p in paginas if "Sócio-Administrador" in p)
        self.assertIn("Encerramento do Período", assinatura)
        self.assertIn("O presente dashboard", assinatura)

    def test_emissao_usa_brasilia_na_capa_e_assinatura(self):
        paginas = self.gerar(contexto())
        self.assertIn("Data de emissão: 30/10/2026", paginas[0])
        self.assertIn("Uberlândia/MG, 30/10/2026", "\n".join(paginas))

    def test_paragrafo_maior_que_pagina_nao_corta_nem_falha(self):
        ctx = contexto()
        ctx["parecer_tecnico"] = "INICIO_LONGO " + "Texto fictício longo. " * 900 + " FIM_LONGO"
        ctx["obs_geral"] = "INICIO_OBS " + "Observação fictícia. " * 750 + " FIM_OBS"
        texto = "\n".join(self.gerar(ctx))
        for marcador in ["INICIO_LONGO", "FIM_LONGO", "INICIO_OBS", "FIM_OBS"]:
            self.assertIn(marcador, texto)

    def test_cloro_total_sem_faixa_nao_recebe_conformidade(self):
        texto = "\n".join(self.gerar(contexto()))
        linha_ct = texto.split("Cloro Total\n", 1)[1].split("Cloro Combinado", 1)[0]
        self.assertNotIn("100%", linha_ct)


class TestRelogioBrasilia(unittest.TestCase):
    def test_virada_do_dia_utc(self):
        from utils.relogio import hoje_brasilia
        self.assertEqual(hoje_brasilia(datetime(2026, 10, 31, 1, 30, tzinfo=timezone.utc)), "30/10/2026")
        self.assertEqual(hoje_brasilia(datetime(2026, 10, 31, 3, 0, tzinfo=timezone.utc)), "31/10/2026")

    def test_virada_do_ano(self):
        from utils.relogio import hoje_brasilia
        self.assertEqual(hoje_brasilia(datetime(2027, 1, 1, 1, 0, tzinfo=timezone.utc)), "31/12/2026")


if __name__ == "__main__":
    unittest.main()
