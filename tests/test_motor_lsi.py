import json
import subprocess
import sys
import unittest
from pathlib import Path

from motor_lsi import calcular_lsi


class TestMotorLsi(unittest.TestCase):
    def test_caso_referencia(self):
        resultado = calcular_lsi(7.8, 90, 100, 29, 2000)

        self.assertAlmostEqual(resultado["lsi"], -0.18817, places=5)
        self.assertEqual(resultado["alcalinidade_carbonatada"], 90.0)

    def test_correcao_cya(self):
        resultado = calcular_lsi(7.4, 110, 100, 29, 2000, cya=100)

        self.assertEqual(resultado["fator_correcao_cya"], 0.31)
        self.assertEqual(resultado["alcalinidade_carbonatada"], 79.0)

    def test_interpolacao_cya(self):
        resultado = calcular_lsi(7.5, 110, 100, 29, 2000, cya=100)

        self.assertAlmostEqual(resultado["fator_correcao_cya"], 0.32)
        self.assertAlmostEqual(resultado["alcalinidade_carbonatada"], 78.0)

    def test_classificacoes(self):
        parametros = (7.5, 100, 100, 25, 1000)

        self.assertEqual(
            calcular_lsi(*parametros, limite_inferior=-0.2, limite_superior=0.2)[
                "classificacao"
            ],
            "tendência subsaturada",
        )
        self.assertEqual(
            calcular_lsi(*parametros, limite_inferior=-1, limite_superior=1)[
                "classificacao"
            ],
            "próximo do equilíbrio",
        )
        self.assertEqual(
            calcular_lsi(*parametros, limite_inferior=-10, limite_superior=-1)[
                "classificacao"
            ],
            "tendência incrustante",
        )

    def test_fator_limitado_fora_da_faixa(self):
        resultado = calcular_lsi(8.2, 110, 100, 29, 2000, cya=100)

        self.assertEqual(resultado["fator_correcao_cya"], 0.36)
        self.assertTrue(resultado["avisos"])
        self.assertIn("extremo mais próximo", resultado["avisos"][0])

    def test_entradas_invalidas(self):
        casos = (
            ((7.4, 100, 100, 29, 0), "TDS"),
            ((7.4, 100, 0, 29, 1000), "Dureza calcica"),
            ((7.4, 0, 100, 29, 1000), "Alcalinidade total"),
            ((7.4, 10, 100, 29, 1000, 100), "Alcalinidade carbonatada"),
            ((7.4, 100, 100, 29, 1000, -1), "CYA"),
        )
        for argumentos, mensagem in casos:
            with self.subTest(mensagem=mensagem):
                with self.assertRaisesRegex(ValueError, mensagem):
                    calcular_lsi(*argumentos)

        with self.assertRaisesRegex(ValueError, "limite inferior"):
            calcular_lsi(7.4, 100, 100, 29, 1000, limite_inferior=1, limite_superior=0)

    def test_retorno_e_serializavel(self):
        resultado = calcular_lsi(7.4, 100, 100, 29, 1000)

        json.dumps(resultado, ensure_ascii=False)
        self.assertEqual(resultado["formula_versao"], "LSI_EPA_PHTA_1.0")

    def test_importacao_sem_dependencias_externas(self):
        raiz = Path(__file__).resolve().parents[1]
        codigo = (
            "import json, sys; import motor_lsi; "
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


if __name__ == "__main__":
    unittest.main()