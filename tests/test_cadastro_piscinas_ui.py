"""Fluxo de edição no app real com planilha fictícia e disco temporário."""
import json
import sys
from streamlit.testing.v1 import AppTest
from test_relatorio_rt_campos import _AppReal, PIN_FICTICIO


class CadastroUI(_AppReal):
    def test_cadastro_bem_star_mantem_empresa(self):
        clientes = sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes")
        clientes.linhas[2][12] = "Bem Star Piscinas"
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)
        botao("🔐 Acesso administrativo").click().run()
        botao("⭐ Bem Star Piscinas").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        at.radio(key="cc_modo_acao").set_value("✏️ Editar cliente existente").run()
        self.assertEqual(at.selectbox(key="cc_sel_editar").options, ["Residencial Beta Fictício"])
        botao("📂 Carregar dados").click().run()
        at.text_input(key="cc_vol_infantil").input("15")
        at.selectbox(key="cc_tipo_piscina_infantil").select("Infantil").run()
        botao("💾 Salvar alterações").click().run()
        self.assertFalse(at.exception, at.exception)
        self.assertEqual(clientes.linhas[2][12], "Bem Star Piscinas")
        self.assertEqual(json.loads(clientes.linhas[2][15])["tipos"]["infantil"], "Infantil")

    def test_edicao_tipos_volumes_e_troca_de_cliente(self):
        at = AppTest.from_file(str(self.pasta / "app" / "app.py"), default_timeout=300)
        at.secrets["admin"] = {"pin": PIN_FICTICIO}
        at.secrets["gcp_service_account"] = {"tipo": "ficticio"}
        at.run()
        botao = lambda rotulo: next(b for b in at.button if b.label == rotulo)
        botao("🔐 Acesso administrativo").click().run()
        botao("🔵 Aqua Gestão").click().run()
        at.text_input(key="pin_admin_input").input(PIN_FICTICIO)
        botao("Entrar no escritório").click().run()
        at.radio(key="cc_modo_acao").set_value("✏️ Editar cliente existente").run()
        self.assertTrue(botao("💾 Salvar alterações").disabled)
        botao("📂 Carregar dados").click().run()
        self.assertFalse(at.exception, at.exception)
        at.text_input(key="cc_vol_adulto").input("100")
        at.text_input(key="cc_pisc_extra1_nome").input("SPA teste")
        at.text_input(key="cc_pisc_extra1_vol").input("12,5")
        at.selectbox(key="cc_tipo_piscina_extra1").select("SPA").run()
        botao("💾 Salvar alterações").click().run()
        self.assertFalse(at.exception, at.exception)
        clientes = sys.modules["gspread"].PLANILHA.worksheet("👥 Clientes")
        row = next(r for r in clientes.linhas if len(r) > 1 and r[1] == "C1")
        self.assertEqual(float(row[3]), 112.5)
        self.assertEqual(json.loads(row[15])["tipos"]["extra1"], "SPA")
        botao("📂 Carregar dados").click().run()
        self.assertEqual(at.selectbox(key="cc_tipo_piscina_extra1").value, "SPA")
        at.selectbox(key="cc_sel_editar").select("Residencial Beta Fictício").run()
        self.assertTrue(botao("💾 Salvar alterações").disabled)
        botao("📂 Carregar dados").click().run()
        self.assertEqual(at.selectbox(key="cc_tipo_piscina_extra1").value, "Não informado")
        self.assertEqual(at.text_input(key="cc_pisc_extra1_vol").value, "")
        at.text_input(key="cc_vol_adulto").input("-10").run()
        self.assertTrue(botao("💾 Salvar alterações").disabled)
        self.assertFalse(at.exception, at.exception)
