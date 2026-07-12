import unittest

from utils.arquivos import classificar_arquivo, chave_segura


class ArquivosTests(unittest.TestCase):
    def test_classificar_arquivo_categorias_reconhecidas(self):
        self.assertEqual(classificar_arquivo("contrato.pdf"), ("Contrato", "PDF"))
        self.assertEqual(classificar_arquivo("aditivo.docx"), ("Aditivo", "DOCX"))
        self.assertEqual(classificar_arquivo("relatorio.txt"), ("Relatório", "Arquivo"))
        self.assertEqual(classificar_arquivo("documento.docx"), ("Documento", "DOCX"))

    def test_classificar_arquivo_com_letras_maiusculas_e_minusculas(self):
        self.assertEqual(classificar_arquivo("Contrato.PDF"), ("Contrato", "PDF"))
        self.assertEqual(classificar_arquivo("ADITIVO.docx"), ("Aditivo", "DOCX"))
        self.assertEqual(classificar_arquivo("RELATORIO.TxT"), ("Relatório", "Arquivo"))

    def test_classificar_arquivo_desconhecido(self):
        self.assertEqual(classificar_arquivo("foto.jpg"), ("Documento", "Arquivo"))
        self.assertEqual(classificar_arquivo("anexo"), ("Documento", "Arquivo"))

    def test_classificar_arquivo_prioridade_das_condicoes(self):
        self.assertEqual(classificar_arquivo("contrato_aditivo.pdf"), ("Contrato", "PDF"))
        self.assertEqual(classificar_arquivo("relatorio_contrato.docx"), ("Contrato", "DOCX"))
        self.assertEqual(classificar_arquivo("aditivo_relatorio.txt"), ("Aditivo", "Arquivo"))

    def test_chave_segura_textos_simples_e_com_espacos(self):
        self.assertEqual(chave_segura("nome"), "nome")
        self.assertEqual(chave_segura("nome do arquivo"), "nome_do_arquivo")

    def test_chave_segura_acentos_e_pontuacao(self):
        self.assertEqual(chave_segura("café/arquivo"), "caf_arquivo")
        self.assertEqual(chave_segura("Olá, mundo!"), "Ol_mundo_")

    def test_chave_segura_maiusculas_e_string_vazia(self):
        self.assertEqual(chave_segura("NomeArquivo"), "NomeArquivo")
        self.assertEqual(chave_segura(""), "")

    def test_chave_segura_resultado_deterministico(self):
        self.assertEqual(chave_segura("texto"), chave_segura("texto"))
        self.assertEqual(chave_segura("a b c"), "a_b_c")


if __name__ == "__main__":
    unittest.main()
