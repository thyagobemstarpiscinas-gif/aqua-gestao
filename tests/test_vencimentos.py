import unittest
from datetime import date
from unittest.mock import patch

from utils.vencimentos import (
    calcular_renovacao_anual,
    status_vencimento,
    texto_dias_restantes,
)


class VencimentosTests(unittest.TestCase):
    @patch("utils.vencimentos.date")
    def test_contrato_com_data_final_valida(self, mock_date):
        mock_date.today.return_value = date(2024, 1, 1)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        resultado = status_vencimento("31/12/2024")
        self.assertEqual(resultado["codigo"], "vigente")
        self.assertEqual(resultado["dias"], 365)

    def test_data_final_vazia(self):
        self.assertEqual(
            status_vencimento(""),
            {
                "codigo": "indefinido",
                "rotulo": "Sem vigência válida",
                "mensagem": "Data final ausente ou inválida.",
                "dias": None,
                "css": "status-indefinido",
            },
        )

    def test_data_final_invalida(self):
        self.assertEqual(
            status_vencimento("31/02/2024"),
            {
                "codigo": "indefinido",
                "rotulo": "Sem vigência válida",
                "mensagem": "Data final ausente ou inválida.",
                "dias": None,
                "css": "status-indefinido",
            },
        )

    @patch("utils.vencimentos.date")
    def test_renovacao_anual(self, mock_date):
        mock_date.today.return_value = date(2024, 1, 1)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        inicio, fim = calcular_renovacao_anual("31/12/2024")
        self.assertEqual(inicio, date(2025, 1, 1))
        self.assertEqual(fim, date(2025, 12, 31))

    @patch("utils.vencimentos.date")
    def test_status_vencido(self, mock_date):
        mock_date.today.return_value = date(2024, 1, 2)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        resultado = status_vencimento("01/01/2024")
        self.assertEqual(resultado["codigo"], "vencido")
        self.assertEqual(resultado["dias"], -1)

    @patch("utils.vencimentos.date")
    def test_status_proximo_do_vencimento(self, mock_date):
        mock_date.today.return_value = date(2024, 1, 1)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        resultado = status_vencimento("05/01/2024", alerta_dias=5)
        self.assertEqual(resultado["codigo"], "vencendo")
        self.assertEqual(resultado["dias"], 4)

    @patch("utils.vencimentos.date")
    def test_status_vigente(self, mock_date):
        mock_date.today.return_value = date(2024, 1, 1)
        mock_date.side_effect = lambda *args, **kwargs: date(*args, **kwargs)
        resultado = status_vencimento("10/01/2024", alerta_dias=5)
        self.assertEqual(resultado["codigo"], "vigente")
        self.assertEqual(resultado["dias"], 9)

    def test_textos_de_dias_restantes(self):
        self.assertEqual(texto_dias_restantes({"dias": -3}), "Atrasado há 3 dia(s)")
        self.assertEqual(texto_dias_restantes({"dias": 0}), "Restam 0 dia(s)")
        self.assertEqual(texto_dias_restantes({"dias": 1}), "Restam 1 dia(s)")
        self.assertEqual(texto_dias_restantes({"dias": 5}), "Restam 5 dia(s)")
        self.assertEqual(texto_dias_restantes({"dias": None}), "Dias restantes: não disponível")


if __name__ == "__main__":
    unittest.main()
