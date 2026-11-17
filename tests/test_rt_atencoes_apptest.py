"""Fluxo real de geração do mensal com atenções, planilha fictícia e sem rede."""
import unittest

from test_rt_relatorio_avaliacao import TestRelatorioRTAvaliacaoCompleta as Base, AppTest
from test_rt_atencoes_operacionais import analises


@unittest.skipIf(AppTest is None, "Streamlit obrigatório")
class FluxoAtencoes(unittest.TestCase):
    setUp = Base.setUp
    _preparar = Base._preparar
    _entrar_e_carregar = Base._entrar_e_carregar
    _mes = Base._mes
    _carregar = Base._carregar
    _gerar = Base._gerar
    _texto = staticmethod(Base._texto)

    def test_formulario_pdf_e_docx_com_todas_atencoes(self):
        # Importador legado do escritório seleciona terças; os valores do caso são
        # os mesmos, com datas fictícias compatíveis com esse fluxo.
        visitas = [(f"{dia}/10/2026", [{**a, "nome": "Piscina Adulto"}]) for dia, a in zip(("06", "13", "20", "27"), analises())]
        self._preparar(visitas)
        at = self._entrar_e_carregar()
        self._mes(at, "10", "2026")
        status = at.selectbox(key="rel_status_agua_select")
        self.assertEqual(status.value, "CONFORME")
        self.assertTrue(any("com pontos de atenção operacional" in o for o in status.options))
        pdf = self._gerar(at)
        documentos = [pdf] + list(pdf.parent.glob("*RELATORIO_RT.docx"))
        self.assertEqual(len(documentos), 2)
        for documento in documentos:
            texto = " ".join(self._texto(documento).split())
            for trecho in ("com pontos de atenção operacional", "no limite aceitável de 0,60", "fora da faixa ideal 30–50"):
                self.assertIn(trecho, texto, documento.name)
            for dia in ("06", "13", "20", "27"):
                self.assertIn(f"{dia}/10/2026, Piscina Adulto", texto, documento.name)


if __name__ == "__main__":
    unittest.main()
