import unittest

from utils.datas_visitas import (
    normalizar_data_visita,
    lancamento_pertence_mes_ano,
    filtrar_lancamentos_rt_tercas,
)


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

    def test_filtrar_lancamentos_rt_tercas_julho_2026_e_casos_variados(self):
        # Terças de julho de 2026 (devem ser mantidas)
        a = {"data": "07/07/2026"}
        b = {"data_visita": "14/07/26"}
        c = {"Data": "2026-07-21"}
        d = {"data": "28/07/2026"}

        # Datas que não são terças (devem ser excluídas)
        e = {"data": "01/07/2026"}
        f = {"data_visita": "06/07/2026"}
        g = {"Data": "10/07/2026"}
        h = {"data": "13/07/2026"}

        # Casos diversos: inválida, vazia, não-dicionário, prioridade de chaves
        i = {"data": "31/02/2026"}  # inválida
        j = {"data": ""}  # vazia
        k = "nao_e_dicionario"
        m = {"data": "01/07/2026", "data_visita": "14/07/2026"}  # usa 'data' primeiro -> excluída
        n = {"outra": "x"}  # sem chaves de data -> excluída

        original_list = [e, a, k, b, i, c, m, d, j, f, g, h, n]

        # preservar cópias para conferir que não são modificadas
        import copy

        originals_copy = copy.deepcopy(original_list)

        filtrados = filtrar_lancamentos_rt_tercas(original_list)

        # Apenas as terças corretas na ordem original: a, b, c, d
        self.assertEqual(filtrados, [a, b, c, d])

        # Ordem preservada: checar que a aparece antes de b etc.
        self.assertTrue(filtrados.index(a) < filtrados.index(b) < filtrados.index(c) < filtrados.index(d))

        # Dicionários originais não foram modificados
        self.assertEqual(original_list, originals_copy)

    def test_filtrar_lancamentos_rt_tercas_none_e_lista_vazia(self):
        self.assertEqual(filtrar_lancamentos_rt_tercas(None), [])
        self.assertEqual(filtrar_lancamentos_rt_tercas([]), [])


if __name__ == "__main__":
    unittest.main()
