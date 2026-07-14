import unittest
from utils.datas_visitas import filtrar_lancamentos_visitas_rt, gerar_datas_tercas_mes

class TestRTFilters(unittest.TestCase):
    def test_filtrar_inclui_apenas_rt(self):
        lancamentos = [
            {"data": "01/07/2026", "operador": "Operador A", "visita_rt_semanal": False},
            {"data": "07/07/2026", "operador": "Thyago", "visita_rt_semanal": True},
            {"data": "14/07/2026", "operador": "Outro", "tipo_visita": "RT Semanal"},
        ]
        res = filtrar_lancamentos_visitas_rt(lancamentos)
        self.assertEqual(len(res), 2)
        # garante que somente os marcados como RT foram mantidos
        self.assertTrue(all((x.get("visita_rt_semanal") or str(x.get("tipo_visita","")).casefold()=="rt semanal") for x in res))

    def test_operador_na_terca_nao_inclui(self):
        # Mesmo quando a data é terça, um lançamento de operador não deve entrar
        lancamentos = [
            {"data": "07/07/2026", "operador": "Operador A", "visita_rt_semanal": False},
            {"data": "07/07/2026", "operador": "Thyago", "visita_rt_semanal": True},
        ]
        res = filtrar_lancamentos_visitas_rt(lancamentos)
        self.assertEqual(len(res), 1)
        self.assertEqual(res[0].get("operador"), "Thyago")

    def test_gerar_datas_tercas_mes(self):
        # Julho 2026 tem terças; verifica que o util retorna listas válidas
        terças = gerar_datas_tercas_mes("7", "2026")
        self.assertIsInstance(terças, list)
        self.assertTrue(all(isinstance(d, str) for d in terças))
        # Ao menos uma terça no mês
        self.assertGreaterEqual(len(terças), 4)

if __name__ == '__main__':
    unittest.main()
