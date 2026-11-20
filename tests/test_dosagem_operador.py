import unittest

from streamlit.testing.v1 import AppTest
from test_relatorio_rt_campos import _AppReal
from utils.dosagem import calcular_sugestoes_dosagem


def calcular(**mudancas):
    dados = dict(ph=7.4, crl=3, alc=100, dc=225, cya=40, volume_m3=100)
    return calcular_sugestoes_dosagem(**(dados | mudancas))


class Dosagem(unittest.TestCase):
    def test_print_e_dicloro(self):
        for produto, quantidade in [("Hipoclorito de cálcio 65%", 449), ("Dicloro 56%", 521)]:
            with self.subTest(produto=produto):
                dose, = calcular(crl=1.95, volume_m3=278, dc=180, alc=90, cya=None, produto_cloro=produto)
                self.assertEqual((dose["quantidade"], dose["unidade"]), (quantidade, "g"))
                self.assertIn("1.95", dose["acao"])

    def test_produtos_por_teor(self):
        for dados, esperado in [
            (dict(alc=70, teor_bicarbonato=100), 5.036),
            (dict(alc=70, teor_bicarbonato=99), 5.087),
            (dict(dc=125, teor_cloreto=100), 11.088),
            (dict(dc=125, teor_cloreto=77), 14.4),
        ]:
            with self.subTest(dados=dados):
                dose, = calcular(**dados)
                self.assertEqual(dose["unidade"], "kg")
                self.assertAlmostEqual(dose["quantidade"], esperado, delta=0.002)

    def test_sem_teor_ou_volume_sem_dose(self):
        for teor in (0, None, -1, 101, float("nan")):
            with self.subTest(teor=teor):
                self.assertTrue(all(s["quantidade"] == 0 for s in
                    calcular(alc=60, dc=90, teor_bicarbonato=teor, teor_cloreto=teor)))
        for volume in (None, 0, -1, float("inf")):
            self.assertEqual(calcular(crl=1, volume_m3=volume), [])

    def test_sem_dose_acida_dupla_ou_percentual_arbitrario_de_drenagem(self):
        orientacoes = calcular(ph=8.2, alc=200, dc=450, cya=120)
        self.assertEqual(len(orientacoes), 3)
        self.assertTrue(all(s["quantidade"] == 0 for s in orientacoes))
        self.assertIn("demanda ácida", orientacoes[0]["justificativa"])
        self.assertIn("reposição", orientacoes[1]["justificativa"])
        self.assertIn("reposição", orientacoes[2]["justificativa"])

    def test_limites_e_dicloro_estabilizado(self):
        for cya in (50.01, 100):
            self.assertEqual(calcular(cya=cya)[0]["produto"], "Monitorar CYA")
        for cya in (100, 101):
            self.assertFalse(any(s["quantidade"] for s in calcular(crl=1, cya=cya, produto_cloro="Dicloro 56%")))
        self.assertEqual(calcular(dc=400, cya=50), [])
        self.assertFalse(any(s["quantidade"] for s in calcular(ph=6.9, crl=5.1)))
        for leitura in ({"crl": -1}, {"alc": float("nan")}, {"ph": 15}):
            self.assertFalse(any(s["quantidade"] for s in calcular(**leitura)))


class FormularioDosagens(_AppReal):
    def test_volume_individual_e_precisao_no_preenchimento(self):
        import sys
        aba = sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes")
        aba.linhas[1][3], aba.linhas[1][9] = "300", "278"
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=120)
        at.secrets["operador"] = {"pin_mestre": "teste-dosagem"}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        def clicar(label):
            next(b for b in at.button if b.label == label).click().run()
            self.assertFalse(at.exception, at.exception)
        clicar("📱 Acessar como Operador")
        next(w for w in at.text_input if w.label == "PIN").input("teste-dosagem")
        clicar("Entrar")
        for campo, valor in dict(ph="7,4", crl="1,95", ct="2,09", alc="90", dc="180").items():
            at.text_input(key=f"op_adulto_{campo}").input(valor)
        at.run()
        self.assertTrue(any("449 g" in m.value for m in at.markdown))
        clicar("✅ Aplicar sugestões de Piscina Adulto")
        self.assertEqual(at.text_input(key="op_dos_adulto_qtd_0").value, "449")
        self.assertEqual(at.text_input(key="op_dos_adulto_un_0").value, "g")
        at.text_input(key="op_adulto_alc").input("70").run()
        pct = next(w for w in at.number_input if "NaHCO₃" in w.label)
        pct.set_value(99.0).run()
        clicar("✅ Aplicar sugestões de Piscina Adulto")
        dose, = calcular(alc=70, volume_m3=278, teor_bicarbonato=99)
        self.assertEqual(float(at.text_input(key="op_dos_adulto_qtd_1").value), dose["quantidade"])
        self.assertEqual(at.text_input(key="op_dos_adulto_un_1").value, "kg")
