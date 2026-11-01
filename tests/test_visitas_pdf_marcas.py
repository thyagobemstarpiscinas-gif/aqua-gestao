"""Gera os três PDFs reais com dados fictícios e confere marca e imagem incorporada."""
import ast
import __future__
import io
import unittest
from datetime import date, datetime
from pathlib import Path
import fitz
from PIL import Image
from docx import Document

ROOT = Path(__file__).resolve().parents[1]


def ambiente():
    tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    ns = {'BASE_DIR': ROOT, 'datetime': datetime, 'date': date,
          'DANCING_SCRIPT_FONT_B64': '', 'LOGO_AQUA_OFICIAL_B64': '', 'CRQ': 'CRQ-MG 2ª Região | CRQ 024025748',
          '_resolver_fotos_visita_para_relatorio': lambda _: {'antes': [], 'depois': [], 'extras': []}}
    constants = {'LOGO_CANDIDATOS', 'LOGO_BEM_STAR_CANDIDATOS', 'NBR_LIM', 'NBR_TXT', 'NBR_CC_MAX'}
    names = {'encontrar_logo', 'encontrar_logo_bem_star', 'gerar_pdf_relatorio_visita_bem_star',
             'gerar_pdf_relatorio_visita', 'gerar_pdf_relatorio_visita_rt', 'valor_float',
             '_normalizar_assinatura_b64', 'atualizar_textos_normativos',
             '_atualizar_tabela_limites_docx', 'normalizar_texto',
             'substituir_placeholders_doc', 'substituir_em_paragrafo'}
    for node in tree.body:
        if (isinstance(node, ast.ImportFrom) and node.module == 'utils.normas_piscinas' or
            isinstance(node, ast.FunctionDef) and node.name in names or
            isinstance(node, ast.Assign) and any(isinstance(x, ast.Name) and x.id in constants for x in node.targets)):
            exec(compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / 'app.py'), 'exec',
                         flags=__future__.annotations.compiler_flag), ns)
    return ns


class VisitasPDF(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = ambiente()
        cls.fixture = {'data': '02/11/2026', 'operador': 'OPERADOR TESTE', 'fotos': [],
                       'piscinas': [{'nome': 'Adulto', 'ph': '7.4', 'cloro_livre': '2.0',
                                     'cloro_total': '2.1', 'alcalinidade': '100', 'dureza': '200'}]}

    def conferir(self, func, logo, marca):
        pdf = self.ns[func](self.fixture, 'CONDOMINIO FICTICIO QA')
        with fitz.open(stream=pdf, filetype='pdf') as doc:
            self.assertIn(marca, '\n'.join(p.get_text() for p in doc))
            # Fontes padrão não incorporadas podem produzir letras sobrepostas
            # em leitores PDF diferentes, apesar da extração textual correta.
            fontes = [item[3] for page in doc for item in page.get_fonts()]
            self.assertTrue(any('BitstreamVeraSans' in fonte for fonte in fontes),
                            'O PDF deve incorporar uma fonte TrueType previsível.')
            # Compara os pixels da imagem incorporada, não só a existência de um logo.
            with Image.open(ROOT / logo) as im:
                esperado = im.convert('RGB').tobytes()
            encontrados = []
            for page in doc:
                for entry in page.get_images():
                    img = doc.extract_image(entry[0])
                    with Image.open(io.BytesIO(img['image'])) as im:
                        encontrados.append(im.convert('RGB').tobytes())
            self.assertIn(esperado, encontrados)

    def test_bem_star_gera_pdf_com_logo(self):
        self.conferir('gerar_pdf_relatorio_visita_bem_star', 'bem_star_logo.png', 'BEM STAR PISCINAS')

    def test_operador_acqua_usa_arte_nova(self):
        self.conferir('gerar_pdf_relatorio_visita', 'acqua_gestao_logo_rt.png', 'ACQUA GESTÃO')

    def test_rt_usa_arte_nova_e_referencia_corrigida(self):
        self.conferir('gerar_pdf_relatorio_visita_rt', 'acqua_gestao_logo_rt.png', '345/2026')
        with fitz.open(stream=self.ns['gerar_pdf_relatorio_visita_rt'](self.fixture, 'QA'), filetype='pdf') as doc:
            self.assertNotIn('Faixa NBR 10339', '\n'.join(p.get_text() for p in doc))

    def test_docx_legado_recebe_referencias_sem_trocar_norma_de_agua_por_projeto(self):
        doc = Document(ROOT / 'relatorio_mensal.docx')
        self.ns['atualizar_textos_normativos'](doc)
        texto = '\n'.join([p.text for p in doc.paragraphs] +
                          [c.text for t in doc.tables for r in t.rows for c in r.cells])
        self.assertIn('345/2026', texto)
        self.assertIn('ABNT NBR 10818:2026', texto)
        self.assertNotIn('NBR 10818:2025', texto)
        for antigo in ('7,2', '0,5', '150'):
            self.assertNotIn(antigo, [c.text for t in doc.tables for r in t.rows for c in r.cells])
        self.assertIn('7,0–7,8', texto)
        self.assertIn('até 100', texto)
        self.assertNotIn('Requisito NBR 11238', texto)
        self.assertNotIn('acreditado pela ANVISA/Inmetro', texto)
        self.assertNotIn('Estabelece os limites físico-químicos aceitáveis', texto)


class VisitasLimites2026(unittest.TestCase):
    """As tabelas das visitas exibem os limites da NBR 10818:2026 e não avaliam CT isolado."""

    @classmethod
    def setUpClass(cls):
        cls.ns = ambiente()

    def texto(self, func, fixture):
        pdf = self.ns[func](fixture, 'CONDOMINIO FICTICIO QA')
        with fitz.open(stream=pdf, filetype='pdf') as doc:
            return '\n'.join(p.get_text() for p in doc)

    def test_tres_visitas_mostram_limites_2026(self):
        fixture = {'data': '02/11/2026', 'operador': 'OPERADOR TESTE', 'fotos': [],
                   'piscinas': [{'nome': 'Adulto', 'ph': '7.1', 'cloro_livre': '4.0',
                                 'cloro_total': '4.2', 'alcalinidade': '70', 'dureza': '120'}]}
        for func in ('gerar_pdf_relatorio_visita_bem_star', 'gerar_pdf_relatorio_visita',
                     'gerar_pdf_relatorio_visita_rt'):
            with self.subTest(func=func):
                texto = self.texto(func, fixture)
                self.assertIn('7,0–7,8', texto)
                self.assertIn('1,0–5,0', texto)
                self.assertNotIn('7,2-7,8', texto)
                self.assertNotIn('7,2–7,8', texto)
                self.assertNotIn('0,5-3,0', texto)
                self.assertNotIn('0,5–3,0', texto)

    def test_rt_sinaliza_ct_inferior_ao_crl_sem_zerar_cc(self):
        fixture = {'data': '02/11/2026', 'operador': 'RT', 'fotos': [],
                   'piscinas': [{'nome': 'Adulto', 'ph': '7.5', 'cloro_livre': '2.5', 'cloro_total': '2.0'}]}
        texto = self.texto('gerar_pdf_relatorio_visita_rt', fixture)
        self.assertIn('CONFERIR MEDICAO', texto)
        self.assertIn('Repetir medicao de cloro livre e total', texto)
