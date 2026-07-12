import unittest

from utils.normalizacao import normalizar_texto_busca, nomes_condominio_equivalentes


class NormalizacaoTests(unittest.TestCase):
    def test_normalizar_texto_busca_texto_simples(self):
        self.assertEqual(normalizar_texto_busca("Condominio Alpha"), "condominio alpha")

    def test_normalizar_texto_busca_maiusculas_e_minusculas(self):
        self.assertEqual(normalizar_texto_busca("CONDOMINIO ALPHA"), "condominio alpha")

    def test_normalizar_texto_busca_acentos(self):
        self.assertEqual(normalizar_texto_busca("Condomínio Águas Claras"), "condominio aguas claras")

    def test_normalizar_texto_busca_pontuacao_e_simbolos(self):
        self.assertEqual(normalizar_texto_busca("Condomínio, Alpha! #123"), "condominio alpha 123")

    def test_normalizar_texto_busca_espacos_duplicados_e_extremidades(self):
        self.assertEqual(normalizar_texto_busca("  Condomínio   Alpha  "), "condominio alpha")

    def test_normalizar_texto_busca_numeros(self):
        self.assertEqual(normalizar_texto_busca("Condomínio 2"), "condominio 2")

    def test_normalizar_texto_busca_string_vazia(self):
        self.assertEqual(normalizar_texto_busca(""), "")

    def test_normalizar_texto_busca_none(self):
        self.assertEqual(normalizar_texto_busca(None), "")

    def test_nomes_condominio_equivalentes_iguais(self):
        self.assertTrue(nomes_condominio_equivalentes("Triad", "Triad"))

    def test_nomes_condominio_equivalentes_diferenca_apenas_de_caixa(self):
        self.assertTrue(nomes_condominio_equivalentes("Triad", "triad"))

    def test_nomes_condominio_equivalentes_diferenca_de_acentos(self):
        self.assertTrue(nomes_condominio_equivalentes("Tríad", "Triad"))

    def test_nomes_condominio_equivalentes_diferenca_de_espacos_e_simbolos(self):
        self.assertTrue(nomes_condominio_equivalentes("Triad Vertical", "Triad-Vertical"))

    def test_nomes_condominio_equivalentes_primeiro_nome_contido_no_segundo(self):
        self.assertTrue(nomes_condominio_equivalentes("Triad", "Triad Vertical"))

    def test_nomes_condominio_equivalentes_segundo_nome_contido_no_primeiro(self):
        self.assertTrue(nomes_condominio_equivalentes("Triad Vertical", "Triad"))

    def test_nomes_condominio_equivalentes_diferentes(self):
        self.assertFalse(nomes_condominio_equivalentes("Triad", "Alphaville"))

    def test_nomes_condominio_equivalentes_ambos_vazios(self):
        self.assertFalse(nomes_condominio_equivalentes("", ""))

    def test_nomes_condominio_equivalentes_primeiro_vazio(self):
        self.assertFalse(nomes_condominio_equivalentes("", "Triad"))

    def test_nomes_condominio_equivalentes_segundo_vazio(self):
        self.assertFalse(nomes_condominio_equivalentes("Triad", ""))

    def test_nomes_condominio_equivalentes_none(self):
        self.assertFalse(nomes_condominio_equivalentes(None, None))


if __name__ == "__main__":
    unittest.main()
