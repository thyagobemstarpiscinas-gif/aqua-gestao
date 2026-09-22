import copy
import json
import subprocess
import sys
import unittest
from pathlib import Path

from motor_lsi import calcular_lsi
from simulador_lsi import analisar_sensibilidade, simular_cenario


PARAMETROS = {
    "ph": 7.8,
    "alcalinidade_total": 110,
    "dureza_calcica": 100,
    "temperatura_c": 29,
    "tds": 2000,
    "cya": 0,
}


class TestSimuladorLsi(unittest.TestCase):
    def test_cenario_sem_alteracoes_tem_delta_zero(self):
        resultado = simular_cenario(PARAMETROS, {})

        self.assertEqual(resultado["delta_lsi"], 0.0)
        self.assertFalse(resultado["mudou_classificacao"])
        self.assertEqual(resultado["parametros_simulados"], PARAMETROS)

    def test_alteracao_de_ph_afeta_somente_ph(self):
        resultado = simular_cenario(PARAMETROS, {"ph": 7.9})

        self.assertEqual(resultado["parametros_simulados"]["ph"], 7.9)
        for parametro in PARAMETROS:
            if parametro != "ph":
                self.assertEqual(
                    resultado["parametros_simulados"][parametro], PARAMETROS[parametro]
                )

    def test_alteracoes_simultaneas_preservam_demais_valores(self):
        resultado = simular_cenario(
            PARAMETROS, {"ph": 7.9, "temperatura_c": 30, "cya": 10}
        )

        self.assertEqual(resultado["parametros_simulados"]["ph"], 7.9)
        self.assertEqual(resultado["parametros_simulados"]["temperatura_c"], 30)
        self.assertEqual(resultado["parametros_simulados"]["cya"], 10)
        self.assertEqual(
            resultado["parametros_simulados"]["alcalinidade_total"],
            PARAMETROS["alcalinidade_total"],
        )
        self.assertEqual(
            resultado["parametros_simulados"]["dureza_calcica"],
            PARAMETROS["dureza_calcica"],
        )
        self.assertEqual(resultado["parametros_simulados"]["tds"], PARAMETROS["tds"])

    def test_dicionarios_recebidos_nao_sao_modificados(self):
        atuais = copy.deepcopy(PARAMETROS)
        alteracoes = {"ph": 7.9, "tds": 2100}
        atuais_antes = copy.deepcopy(atuais)
        alteracoes_antes = copy.deepcopy(alteracoes)

        simular_cenario(atuais, alteracoes)

        self.assertEqual(atuais, atuais_antes)
        self.assertEqual(alteracoes, alteracoes_antes)

    def test_chave_desconhecida_gera_erro_em_portugues(self):
        with self.assertRaisesRegex(ValueError, "Chave.*desconhecida"):
            simular_cenario(PARAMETROS, {"cloro": 1})
        with self.assertRaisesRegex(ValueError, "Chave.*desconhecida"):
            analisar_sensibilidade(PARAMETROS, {"cloro": 1})

    def test_entrada_invalida_e_rejeitada_pelo_motor(self):
        parametros = dict(PARAMETROS, tds=0)

        with self.assertRaisesRegex(ValueError, "TDS deve ser maior que zero"):
            simular_cenario(parametros, {})

    def test_sensibilidade_retorna_os_seis_parametros(self):
        resultado = analisar_sensibilidade(PARAMETROS)

        self.assertEqual(
            {item["parametro"] for item in resultado["resultados"]},
            set(PARAMETROS),
        )
        self.assertEqual(len(resultado["resultados"]), 6)

    def test_ranking_e_ordenado_por_impacto_absoluto(self):
        resultado = analisar_sensibilidade(PARAMETROS)
        impactos = [item["impacto_absoluto"] for item in resultado["resultados"]]

        self.assertEqual(impactos, sorted(impactos, reverse=True))
        self.assertEqual(
            resultado["ranking_influencia"],
            [item["parametro"] for item in resultado["resultados"]],
        )

    def test_variacoes_personalizadas_sao_respeitadas(self):
        variacoes = {parametro: 0 for parametro in PARAMETROS}
        variacoes["ph"] = 0.2

        resultado = analisar_sensibilidade(PARAMETROS, variacoes)
        por_parametro = {
            item["parametro"]: item for item in resultado["resultados"]
        }

        self.assertEqual(resultado["variacoes_utilizadas"], variacoes)
        self.assertEqual(por_parametro["ph"]["valor_variado"], 8.0)
        self.assertEqual(por_parametro["tds"]["valor_variado"], PARAMETROS["tds"])

    def test_variacao_invalida_e_rejeitada(self):
        with self.assertRaisesRegex(ValueError, "CYA nao pode ser negativo"):
            analisar_sensibilidade(PARAMETROS, {"cya": -1})
        with self.assertRaisesRegex(ValueError, "deve ser um numero"):
            analisar_sensibilidade(PARAMETROS, {"ph": "0.1"})

    def test_modulos_nao_importam_dependencias_externas(self):
        raiz = Path(__file__).resolve().parents[1]
        codigo = (
            "import json, sys; import simulador_lsi; "
            "print(json.dumps(sorted(name for name in sys.modules "
            "if name == 'streamlit' or name == 'gspread' or name == 'reportlab' "
            "or name.startswith('google.'))))"
        )
        processo = subprocess.run(
            [sys.executable, "-c", codigo],
            cwd=raiz,
            check=True,
            capture_output=True,
            text=True,
        )

        self.assertEqual(json.loads(processo.stdout), [])

    def test_simulador_delega_calculo_ao_motor_sem_reproduzir_formula(self):
        resultado = simular_cenario(PARAMETROS, {"ph": 7.9})
        esperado = calcular_lsi(
            **PARAMETROS,
            limite_inferior=-0.2,
            limite_superior=0.2,
        )
        codigo = Path(__file__).resolve().parents[1] / "simulador_lsi.py"

        self.assertEqual(resultado["cenario_atual"]["lsi"], esperado["lsi"])
        self.assertNotIn("log10", codigo.read_text(encoding="utf-8"))
        self.assertIn("calcular_lsi", codigo.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
