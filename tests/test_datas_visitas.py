import unittest

from utils.datas_visitas import normalizar_data_visita, lancamento_pertence_mes_ano


class DatasVisitasTests(unittest.TestCase):
    def test_normalizar_data_visita_formatos_aceitos(self):
        self.assertEqual(normalizar_data_visita("17/04/2026"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("17/04/26"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("17-04-2026"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("17-04-26"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("2026-04-17"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("17042026"), "17/04/2026")
        self.assertEqual(normalizar_data_visita("170426"), "01/07/426")
        self.assertEqual(normalizar_data_visita("20260417"), "17/04/2026")

    def test_normalizar_data_visita_data_bissexta_valida(self):
        self.assertEqual(normalizar_data_visita("29/02/2024"), "29/02/2024")

    def test_normalizar_data_visita_data_impossivel(self):
        self.assertEqual(normalizar_data_visita("31/02/2024"), "31/02/2024")

    def test_normalizar_data_visita_texto_invalido_preserva_retorno(self):
        self.assertEqual(normalizar_data_visita("abc"), "abc")

    def test_normalizar_data_visita_string_vazia(self):
        self.assertEqual(normalizar_data_visita(""), "")

    def test_normalizar_data_visita_none(self):
        self.assertEqual(normalizar_data_visita(None), "")

    def test_normalizar_data_visita_espacos_nas_extremidades(self):
        self.assertEqual(normalizar_data_visita(" 17/04/2026 "), "17/04/2026")

    def test_lancamento_pertence_mes_ano_ok(self):
        self.assertTrue(lancamento_pertence_mes_ano("17/04/2026", "4", "2026"))

    def test_lancamento_pertence_mes_ano_mes_diferente(self):
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "5", "2026"))

    def test_lancamento_pertence_mes_ano_ano_diferente(self):
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "4", "2025"))

    def test_lancamento_pertence_mes_ano_zero_a_esquerda(self):
        self.assertTrue(lancamento_pertence_mes_ano("17/04/2026", "04", "2026"))

    def test_lancamento_pertence_mes_ano_em_cada_formato_aceito(self):
        self.assertTrue(lancamento_pertence_mes_ano("17/04/2026", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("17/04/26", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("17-04-2026", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("17-04-26", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("2026-04-17", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("17042026", "4", "2026"))
        self.assertFalse(lancamento_pertence_mes_ano("170426", "4", "2026"))
        self.assertTrue(lancamento_pertence_mes_ano("20260417", "4", "2026"))

    def test_lancamento_pertence_mes_ano_data_invalida(self):
        self.assertFalse(lancamento_pertence_mes_ano("abc", "4", "2026"))

    def test_lancamento_pertence_mes_ano_mes_invalido(self):
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "x", "2026"))

    def test_lancamento_pertence_mes_ano_ano_invalido(self):
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "4", "x"))

    def test_lancamento_pertence_mes_ano_valores_vazios_e_none(self):
        self.assertFalse(lancamento_pertence_mes_ano("", "4", "2026"))
        self.assertFalse(lancamento_pertence_mes_ano(None, "4", "2026"))
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "", "2026"))
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", None, "2026"))
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "4", ""))
        self.assertFalse(lancamento_pertence_mes_ano("17/04/2026", "4", None))


if __name__ == "__main__":
    unittest.main()
