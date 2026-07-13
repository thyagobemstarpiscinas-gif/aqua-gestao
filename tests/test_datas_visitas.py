import unittest

from datetime import datetime

from utils.datas_visitas import (
    normalizar_data_visita,
    lancamento_pertence_mes_ano,
    filtrar_lancamentos_rt_tercas,
    filtrar_lancamentos_visitas_rt,
    gerar_datas_tercas_mes,
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

    def test_filtrar_lancamentos_visitas_rt_condicoes_de_rt(self):
        a = {"data": "07/07/2026", "visita_rt_semanal": False, "tipo_visita": "Operador"}
        b = {"data": "14/07/2026", "visita_rt_semanal": True, "tipo_visita": "RT semanal"}
        c = {"data": "15/07/2026", "visita_rt_semanal": "true", "tipo_visita": "Operador"}
        d = {"data": "16/07/2026", "visita_rt_semanal": "1", "tipo_visita": "operador"}
        e = {"data": "17/07/2026", "visita_rt_semanal": "sim", "tipo_visita": "Operador"}
        f = {"data": "18/07/2026", "visita_rt_semanal": "TRUE", "tipo_visita": "Operador"}
        g = {"data": "19/07/2026", "visita_rt_semanal": False, "tipo_visita": "  Rt SeMaNaL  "}
        h = {"data": "20/07/2026", "visita_rt_semanal": False, "tipo_visita": "Operador"}
        i = {"data": "21/07/2026", "visita_rt_semanal": "no", "tipo_visita": "Operador"}
        j = {"data": "22/07/2026", "visita_rt_semanal": None, "tipo_visita": "RT"}
        k = {"data": "23/07/2026", "tipo_visita": "RT semanal"}
        l = {"data": "24/07/2026", "visita_rt_semanal": "0", "tipo_visita": "Operador"}
        m = {"data": "25/07/2026", "visita_rt_semanal": "false", "tipo_visita": "Operador"}
        n = {"data": "26/07/2026", "visita_rt_semanal": "não", "tipo_visita": "Operador"}
        o = "nao_e_dicionario"

        original_list = [a, b, c, d, e, f, g, h, i, j, k, l, m, n, o]
        import copy
        originals_copy = copy.deepcopy(original_list)

        filtrados = filtrar_lancamentos_visitas_rt(original_list)

        self.assertEqual(filtrados, [b, c, d, e, f, g, k])
        self.assertEqual(original_list, originals_copy)
        self.assertTrue(all(isinstance(item, dict) for item in filtrados))
        self.assertEqual(filtrados[0], b)
        self.assertEqual(filtrados[-1], k)

    def test_filtrar_lancamentos_visitas_rt_rejeita_valores_falsos(self):
        casos = [False, "false", "0", "não", ""]
        lista = [{"data": "07/07/2026", "visita_rt_semanal": valor, "tipo_visita": "Operador"} for valor in casos]
        self.assertEqual(filtrar_lancamentos_visitas_rt(lista), [])

    def test_filtrar_lancamentos_visitas_rt_none_e_lista_vazia(self):
        self.assertEqual(filtrar_lancamentos_visitas_rt(None), [])
        self.assertEqual(filtrar_lancamentos_visitas_rt([]), [])

    def test_filtrar_lancamentos_visitas_rt_preserva_ordem_e_objetos(self):
        a = {"data": "07/07/2026", "visita_rt_semanal": True}
        b = {"data": "08/07/2026", "tipo_visita": "rt semanal"}
        c = {"data": "09/07/2026", "visita_rt_semanal": False, "tipo_visita": "Operador"}
        original_list = [a, b, c]
        import copy
        originals_copy = copy.deepcopy(original_list)
        filtrados = filtrar_lancamentos_visitas_rt(original_list)
        self.assertEqual(filtrados, [a, b])
        self.assertEqual(original_list, originals_copy)
        self.assertIs(filtrados[0], a)
        self.assertIs(filtrados[1], b)

    def test_filtrar_lancamentos_visitas_rt_nao_considera_operador_thyago_sem_rt(self):
        item = {"data": "07/07/2026", "visita_rt_semanal": False, "tipo_visita": "Operador", "operador": "Thyago"}
        self.assertEqual(filtrar_lancamentos_visitas_rt([item]), [])

    def test_filtrar_lancamentos_rt_tercas_none_e_lista_vazia(self):
        self.assertEqual(filtrar_lancamentos_rt_tercas(None), [])
        self.assertEqual(filtrar_lancamentos_rt_tercas([]), [])

    def test_gerar_datas_tercas_mes_julho_2026(self):
        esperadas = ["07/07/2026", "14/07/2026", "21/07/2026", "28/07/2026"]
        self.assertEqual(gerar_datas_tercas_mes("7", "2026"), esperadas)
        self.assertEqual(gerar_datas_tercas_mes("07", "2026"), esperadas)

    def test_gerar_datas_tercas_mes_setembro_2026(self):
        esperadas = [
            "01/09/2026",
            "08/09/2026",
            "15/09/2026",
            "22/09/2026",
            "29/09/2026",
        ]
        self.assertEqual(gerar_datas_tercas_mes("9", "2026"), esperadas)

    def test_gerar_datas_tercas_fevereiro_bissexto_e_comum(self):
        # Fevereiro 2021 (comum)
        feb2021 = gerar_datas_tercas_mes("2", "2021")
        # Deve conter 4 terças
        self.assertTrue(all(d.endswith("/2021") for d in feb2021))
        self.assertEqual(len(feb2021), 4)

        # Fevereiro 2024 (bissexto)
        feb2024 = gerar_datas_tercas_mes("02", "2024")
        self.assertTrue(all(d.endswith("/2024") for d in feb2024))
        self.assertEqual(len(feb2024), 4)

    def test_gerar_datas_tercas_meses_invalidos_e_edge_cases(self):
        self.assertEqual(gerar_datas_tercas_mes("0", "2026"), [])
        self.assertEqual(gerar_datas_tercas_mes("13", "2026"), [])
        self.assertEqual(gerar_datas_tercas_mes("", "2026"), [])
        self.assertEqual(gerar_datas_tercas_mes(None, "2026"), [])
        self.assertEqual(gerar_datas_tercas_mes("7", ""), [])
        self.assertEqual(gerar_datas_tercas_mes("7", None), [])
        self.assertEqual(gerar_datas_tercas_mes("7", "26"), [])

    def test_gerar_datas_tercas_ordem_e_formato(self):
        res = gerar_datas_tercas_mes("9", "2026")
        # Ordem cronológica
        self.assertEqual(res, sorted(res, key=lambda s: datetime.strptime(s, "%d/%m/%Y")))
        # Formato dd/mm/aaaa
        for s in res:
            self.assertRegex(s, r"^\d{2}/\d{2}/\d{4}$")


if __name__ == "__main__":
    unittest.main()
