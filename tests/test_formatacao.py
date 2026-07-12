import unittest

from utils.formatacao import (
    apenas_digitos,
    formatar_cnpj,
    formatar_cpf,
    formatar_telefone,
    humanizar_nome_pasta,
    limpar_nome_arquivo,
    slugify_nome,
    validar_email,
)


class FormatacaoTests(unittest.TestCase):
    def test_slugify_nome_com_entrada_valida(self):
        self.assertEqual(slugify_nome("Condomínio Novo"), "Condomínio_Novo")
        self.assertEqual(slugify_nome("Residencial 01"), "Residencial_01")

    def test_slugify_nome_com_entrada_vazia(self):
        self.assertEqual(slugify_nome(""), "condominio")
        self.assertEqual(slugify_nome(None), "condominio")

    def test_humanizar_nome_pasta_com_entrada_valida(self):
        self.assertEqual(humanizar_nome_pasta("Residencial_01"), "Residencial 01")
        self.assertEqual(humanizar_nome_pasta("Condominio-Novo"), "Condominio Novo")

    def test_humanizar_nome_pasta_vazia(self):
        self.assertEqual(humanizar_nome_pasta(""), "")
        self.assertEqual(humanizar_nome_pasta(None), "")

    def test_limpar_nome_arquivo_com_caracteres_especiais(self):
        self.assertEqual(limpar_nome_arquivo('Meu/Arquivo:Nome?*'), "MeuArquivoNome")
        self.assertEqual(limpar_nome_arquivo("  Nome   com   espaços  "), "Nome_com_espaços")

    def test_limpar_nome_arquivo_vazio(self):
        self.assertEqual(limpar_nome_arquivo(""), "")

    def test_apenas_digitos(self):
        self.assertEqual(apenas_digitos("abc12345xyz"), "12345")
        self.assertEqual(apenas_digitos(""), "")
        self.assertEqual(apenas_digitos(None), "")

    def test_formatar_cpf(self):
        self.assertEqual(formatar_cpf("12345678909"), "123.456.789-09")
        self.assertEqual(formatar_cpf("123"), "123")
        self.assertEqual(formatar_cpf(""), "")

    def test_formatar_cnpj(self):
        self.assertEqual(formatar_cnpj("12345678000195"), "12.345.678/0001-95")
        self.assertEqual(formatar_cnpj("12"), "12")
        self.assertEqual(formatar_cnpj(""), "")

    def test_formatar_telefone(self):
        self.assertEqual(formatar_telefone("11987654321"), "(11) 98765-4321")
        self.assertEqual(formatar_telefone("5511987654321"), "(11) 98765-4321")
        self.assertEqual(formatar_telefone("123"), "(12) 3")

    def test_validar_email(self):
        self.assertTrue(validar_email("usuario@email.com"))
        self.assertTrue(validar_email(""))
        self.assertFalse(validar_email("email-invalido"))
        self.assertFalse(validar_email("usuario@dominio"))


if __name__ == "__main__":
    unittest.main()
