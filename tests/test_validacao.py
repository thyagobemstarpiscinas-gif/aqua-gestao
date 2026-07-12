import unittest

from utils.validacao import validar_cnpj, validar_cpf, validar_data_br


class ValidacaoTests(unittest.TestCase):
    def test_validar_cpf_valido_e_invalido(self):
        self.assertTrue(validar_cpf("52998224725"))
        self.assertFalse(validar_cpf("12345678900"))

    def test_validar_cpf_com_e_sem_pontuacao(self):
        self.assertTrue(validar_cpf("529.982.247-25"))
        self.assertTrue(validar_cpf("52998224725"))

    def test_validar_cpf_todos_digitos_iguais(self):
        self.assertFalse(validar_cpf("11111111111"))

    def test_validar_cnpj_valido_e_invalido(self):
        self.assertTrue(validar_cnpj("12345678000195"))
        self.assertFalse(validar_cnpj("12345678000196"))

    def test_validar_cnpj_com_e_sem_pontuacao(self):
        self.assertTrue(validar_cnpj("12.345.678/0001-95"))
        self.assertTrue(validar_cnpj("12345678000195"))

    def test_validar_cnpj_todos_digitos_iguais(self):
        self.assertFalse(validar_cnpj("11111111111111"))

    def test_validar_data_br_valida_invalida_e_vazia(self):
        self.assertTrue(validar_data_br("31/12/2024"))
        self.assertFalse(validar_data_br("31/02/2024"))
        self.assertFalse(validar_data_br(""))


if __name__ == "__main__":
    unittest.main()
