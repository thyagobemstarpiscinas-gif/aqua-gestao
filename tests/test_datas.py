import unittest
from datetime import date, datetime

from utils.datas import (
    adicionar_um_ano,
    formatar_data_br,
    formatar_data_hora_arquivo,
    hoje_br,
    parse_data_br,
)


class DatasTests(unittest.TestCase):
    def test_hoje_br_no_formato_dd_mm_aaaa(self):
        hoje = hoje_br()
        self.assertRegex(hoje, r"^\d{2}/\d{2}/\d{4}$")

    def test_formatar_data_hora_arquivo_com_timestamp_conhecido(self):
        ts = datetime(2024, 1, 2, 3, 4, 5).timestamp()
        self.assertEqual(formatar_data_hora_arquivo(ts), "02/01/2024 03:04")

    def test_parse_data_br_com_data_valida_invalida_e_vazia(self):
        self.assertEqual(parse_data_br("31/12/2024"), date(2024, 12, 31))
        self.assertIsNone(parse_data_br("31/02/2024"))
        self.assertIsNone(parse_data_br(""))

    def test_formatar_data_br_com_data_conhecida(self):
        self.assertEqual(formatar_data_br(date(2024, 1, 2)), "02/01/2024")

    def test_adicionar_um_ano_em_data_comum(self):
        self.assertEqual(adicionar_um_ano(date(2024, 1, 15)), date(2025, 1, 15))

    def test_adicionar_um_ano_em_29_02_de_ano_bissexto(self):
        self.assertEqual(adicionar_um_ano(date(2024, 2, 29)), date(2025, 2, 28))


if __name__ == "__main__":
    unittest.main()
