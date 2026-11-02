"""Gera o PDF mensal RT real do app.py (build_monthly_rt_pdf_bytes) com dados fictícios.

Carrega somente as definições necessárias do app.py, sob demanda, sem executar o
Streamlit, sem Sheets e sem rede.
"""
import ast
import re
import __future__
import builtins
import unittest
from datetime import date, datetime
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]


def _carregar():
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    defs = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            defs[node.name] = node
        elif isinstance(node, ast.ImportFrom) and (node.module or "").startswith("utils"):
            for nome in node.names:
                defs.setdefault(nome.asname or nome.name, node)
        elif isinstance(node, ast.Assign):
            for alvo in node.targets:
                if isinstance(alvo, ast.Name):
                    defs.setdefault(alvo.id, node)

    class _St:
        session_state, secrets = {}, {}

        def __getattr__(self, _nome):
            return lambda *a, **k: None

    class _NS(dict):
        def __missing__(self, nome):
            if nome in defs:
                exec(compile(ast.Module(body=[defs[nome]], type_ignores=[]), str(ROOT / "app.py"), "exec",
                             flags=__future__.annotations.compiler_flag), self)
                if nome in self:
                    return self[nome]
            if hasattr(builtins, nome):
                return getattr(builtins, nome)
            raise KeyError(nome)

    return _NS(st=_St(), BASE_DIR=ROOT, Path=Path, datetime=datetime, date=date, re=re)


class MensalRTApp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _carregar()

    def _texto(self, registros):
        pdf = self.ns["build_monthly_rt_pdf_bytes"](registros, "CONDOMINIO FICTICIO QA", "11", "2026", "30/11/2026")
        with fitz.open(stream=pdf, filetype="pdf") as doc:
            return " ".join(" ".join(p.get_text() for p in doc).split())

    def test_decimal_com_ponto_nao_vira_desvio(self):
        texto = self._texto([{"data": "04/11/2026", "ph": "7.4", "cloro_livre": "2.5", "cloro_total": "2.8",
                              "alcalinidade": "100", "dureza": "200", "cianurico": "40"}])
        self.assertIn("7,40", texto)
        self.assertNotIn(" 74 ", texto)
        self.assertIn("AVALIAÇÃO PARCIAL", texto)
        self.assertIn("0,30", texto)  # CC = 2,8 - 2,5
        self.assertNotIn("<br/>", texto)
        self.assertIn("NBR 10818:2026", texto)

    def test_limites_2026_e_ct_inferior_ao_crl(self):
        texto = self._texto([
            {"data": "11/11/2026", "ph": "7,1", "cloro_livre": "0,8", "cloro_total": "1,6"},
            {"data": "18/11/2026", "ph": "7,6", "cloro_livre": "3,0", "cloro_total": "2,5"},
        ])
        self.assertIn("11/11/2026 — verificar: CRL, cloro combinado (limites da NBR 10818:2026)", texto)
        self.assertIn("18/11/2026 — verificar: CT < CRL (medição inconsistente)", texto)
        self.assertIn("Rever", texto)
        self.assertNotIn("pH, CRL", texto)  # pH 7,1 está dentro do limite 7,0–7,8


if __name__ == "__main__":
    unittest.main()


class AvaliacaoAutomatica(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _carregar()

    def test_limites_2026_cc_e_recomendacao_de_ph(self):
        analises = [
            {"ph": "6,8", "cloro_livre": "4,0", "cloro_total": "4,3", "alcalinidade": "70", "dureza": "120", "cianurico": "80"},
            {"ph": "7,4", "cloro_livre": "0,8", "cloro_total": "1,6"},
            {"ph": "7,4", "cloro_livre": "2,5", "cloro_total": "2,0"},
        ]
        aval = self.ns["avaliar_conformidade_analises"](analises)
        detalhes = " | ".join(aval["detalhes"])
        self.assertEqual(aval["status"], "NÃO CONFORME")
        self.assertIn("Linha 1: pH=6,8 fora do limite 7,0–7,8", detalhes)
        self.assertNotIn("Linha 1: cloro residual livre", detalhes)  # 4,0 está no limite 1,0–5,0
        self.assertNotIn("alcalinidade", detalhes)
        self.assertNotIn("dureza", detalhes)
        self.assertIn("Linha 2: cloro residual livre=0,8", detalhes)
        self.assertIn("Linha 2: cloro combinado=0,8", detalhes)
        self.assertIn("Linha 3: cloro total inferior ao livre", detalhes)
        self.assertIsNone(analises[2]["cloro_combinado"])
        textos = self.ns["gerar_textos_automaticos_relatorio"](analises, aval)
        recs = " ".join(textos["recomendacoes"])
        self.assertIn("Ajustar pH", recs)
        self.assertIn("Repetir as medições", recs)

    def test_sem_desvios_nao_afirma_conformidade_integral(self):
        analises = [{"ph": "7,5", "cloro_livre": "2,5", "cloro_total": "2,7"}]
        aval = self.ns["avaliar_conformidade_analises"](analises)
        texto = self.ns["gerar_textos_automaticos_relatorio"](analises, aval)["diagnostico"]
        self.assertEqual(aval["status"], "CONFORME")
        self.assertIn("avaliação é parcial", texto)
