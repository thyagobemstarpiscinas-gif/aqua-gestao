import ast
import unittest
from pathlib import Path


class AppStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app_path = Path(__file__).resolve().parents[1] / "app.py"
        cls.source = cls.app_path.read_text(encoding="utf-8")
        cls.tree = ast.parse(cls.source, filename=str(cls.app_path))

    def test_app_python_source_can_be_parsed_and_compiled(self):
        ast.parse(self.source, filename=str(self.app_path))
        compile(self.source, str(self.app_path), "exec")

    def test_expected_functions_exist(self):
        function_names = {
            node.name
            for node in ast.walk(self.tree)
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        }

        expected_functions = {
            "conectar_drive",
            "conectar_sheets",
            "validar_pin_operador",
            "_admin_pin_valido",
            "gerar_documento",
            "gerar_relatorio_mensal",
            "calcular_sugestoes_dosagem",
        }

        self.assertTrue(expected_functions.issubset(function_names))

    def test_no_session_state_clear_call_exists(self):
        for node in ast.walk(self.tree):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                func = node.func
                if (
                    isinstance(func.value, ast.Attribute)
                    and isinstance(func.value.value, ast.Name)
                    and func.value.value.id == "st"
                    and func.value.attr == "session_state"
                    and func.attr == "clear"
                ):
                    self.fail("Encontrado st.session_state.clear() no código")

        self.assertNotIn("st.session_state.clear()", self.source)


if __name__ == "__main__":
    unittest.main()
