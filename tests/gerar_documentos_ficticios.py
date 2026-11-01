"""Gera documentos fictícios (PDF/DOCX) a partir dos geradores reais, sem rede e sem Sheets.

Uso: python tests/gerar_documentos_ficticios.py [PASTA_SAIDA]
Padrão: generated/inspecao_ficticia (pasta ignorada pelo Git). Somente dados fictícios.
"""
import ast, __future__, builtins, sys, os, io, socket
from pathlib import Path
from datetime import datetime, date

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(sys.argv[1] if len(sys.argv) > 1 else ROOT / "generated" / "inspecao_ficticia").resolve()
OUT.mkdir(parents=True, exist_ok=True)
os.chdir(ROOT); sys.path.insert(0, str(ROOT))
def _negar(*a, **k): raise RuntimeError("rede bloqueada")
socket.socket.connect = _negar; socket.create_connection = _negar

SRC = (ROOT / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SRC)
DEFS = {}
for node in TREE.body:
    if isinstance(node, (ast.FunctionDef, ast.ClassDef)):
        DEFS[node.name] = node
    elif isinstance(node, ast.Assign):
        for t in node.targets:
            if isinstance(t, ast.Name): DEFS.setdefault(t.id, node)
    elif isinstance(node, ast.ImportFrom) and node.module and (node.module.startswith("utils") or node.module == "rt_monthly_report"):
        for a in node.names: DEFS.setdefault(a.asname or a.name, node)

class FakeSt:
    session_state = {}
    secrets = {}
    def __getattr__(self, n): return lambda *a, **k: None

class NS(dict):
    def __missing__(self, nome):
        if nome in DEFS:
            node = DEFS[nome]
            exec(compile(ast.Module(body=[node], type_ignores=[]), str(ROOT / "app.py"), "exec",
                         flags=__future__.annotations.compiler_flag), self)
            if nome in self: return self[nome]
        if hasattr(builtins, nome): return getattr(builtins, nome)
        raise KeyError(nome)

import re, json, hashlib, base64, math
from docx import Document
ns = NS(st=FakeSt(), BASE_DIR=ROOT, GENERATED_DIR=OUT / "_gen", datetime=datetime, date=date, Path=Path,
        re=re, json=json, io=io, hashlib=hashlib, base64=base64, math=math, Document=Document,
        _log_sheets_erro=lambda *a, **k: None,
        _resolver_fotos_visita_para_relatorio=lambda _: {"antes": [], "depois": [], "extras": []})

from PIL import Image, ImageDraw
foto = OUT / "foto_ficticia.jpg"
im = Image.new("RGB", (1200, 800), (40, 140, 200)); d = ImageDraw.Draw(im)
d.rectangle((80, 80, 1120, 720), outline=(255, 255, 255), width=12); d.text((120, 380), "IMAGEM FICTICIA PARA TESTE", fill=(255, 255, 255))
im.save(foto, quality=85)

resultados = []
def registrar(nome, fn):
    try:
        fn(); resultados.append((nome, "OK"))
    except Exception as e:
        import traceback; traceback.print_exc(); resultados.append((nome, f"ERRO {type(e).__name__}: {e}"))

COND = "CONDOMÍNIO FICTÍCIO QA"
visita = {"data": "02/11/2026", "operador": "OPERADOR TESTE", "fotos": [], "observacoes": "Registro fictício para inspeção visual.",
          "piscinas": [
              {"nome": "Adulto", "ph": "7.1", "cloro_livre": "0.8", "cloro_total": "1.6", "alcalinidade": "70", "dureza": "120", "cianurico": "40"},
              {"nome": "Infantil", "ph": "7.5", "cloro_livre": "2.5", "cloro_total": "2.0", "alcalinidade": "100", "dureza": "200", "cianurico": "35"}]}

for func, arq in [("gerar_pdf_relatorio_visita_bem_star", "01_visita_bem_star.pdf"),
                  ("gerar_pdf_relatorio_visita", "02_visita_operador_acqua.pdf"),
                  ("gerar_pdf_relatorio_visita_rt", "03_visita_supervisao_rt.pdf")]:
    registrar(arq, lambda func=func, arq=arq: (OUT / arq).write_bytes(ns[func](visita, COND)))

registros = [
    {"data": "04/11/2026", "condominio": COND, "ph": "7,5", "cloro_livre": "2,5", "cloro_total": "2,8", "alcalinidade": "100", "dureza": "200", "cya": "40", "turbidez": "0,3", "orp": "700", "tds": "800", "temperatura": "27", "is_rt": True, "tipo_visita": "RT"},
    {"data": "11/11/2026", "condominio": COND, "ph": "7,1", "cloro_livre": "0,8", "cloro_total": "1,6", "alcalinidade": "70", "dureza": "120", "cya": "80", "turbidez": "0,5", "orp": "620", "tds": "900", "temperatura": "28", "is_rt": True, "tipo_visita": "RT"},
    {"data": "18/11/2026", "condominio": COND, "ph": "7,6", "cloro_livre": "3,0", "cloro_total": "2,5", "alcalinidade": "110", "dureza": "220", "cya": "45", "turbidez": "0,2", "orp": "710", "tds": "850", "temperatura": "27", "is_rt": True, "tipo_visita": "RT"},
]
# rt_monthly_report.build_monthly_rt_pdf_bytes é sobrescrito pelo app.py (versão em uso).
recs_app = [{"data": r["data"], "ph": r["ph"], "cloro_livre": r["cloro_livre"], "cloro_total": r["cloro_total"], "alcalinidade": r["alcalinidade"],
             "dureza": r["dureza"], "cianurico": r["cya"], "turbidez": r["turbidez"], "orp": r["orp"], "tds": r["tds"], "temperatura": r["temperatura"]} for r in registros]
registrar("05_mensal_rt_app.pdf", lambda: (OUT / "05_mensal_rt_app.pdf").write_bytes(
    ns["build_monthly_rt_pdf_bytes"](recs_app, COND, "11", "2026", "30/11/2026", address="Rua Fictícia, 100")))

from utils import rt_relatorio as rr
from utils.pdf_rt_consolidado import gerar_pdf_consolidado
def consolidado():
    meses = rr.meses_do_periodo("Trimestral", 4, 2026)
    linhas = [{"data": f"0{d}/{m}/2026", "piscina": p, "ph": f"7,{d}", "cloro_livre": "0,8" if d == 1 else "2,5", "cloro_total": "2,9",
               "alcalinidade": "90", "dureza": "200", "cianurico": "40"} for m in (10, 11, 12) for d in (1, 5) for p in ("Adulto", "Infantil")]
    cons = rr.consolidar(linhas, meses)
    gdir = OUT / "_graficos"; graficos = rr.graficos_consolidados(cons, gdir, COND, "4º trimestre de 2026")
    ctx = {"tipo": "Trimestral", "periodo_rotulo": "4º trimestre de 2026", "meses": cons["meses"], "nome_condominio": COND,
           "consolidado": cons, "graficos": graficos, "piscinas_cadastro": rr.piscinas_do_cliente({"vol_adulto": 120, "vol_infantil": 15}),
           "art": {"status": "Em processo de formalização"}, "responsavel": "RT FICTÍCIO", "crq": "CRQ 000", "data_emissao": "31/12/2026"}
    ok, erro = gerar_pdf_consolidado(ctx, OUT / "06_consolidado_rt.pdf", ROOT / "acqua_gestao_logo_rt.png")
    if not ok: raise RuntimeError(erro)
registrar("06_consolidado_rt.pdf", consolidado)

from utils.pdf_bem_star_mensal import gerar_pdf as gerar_bs
from utils.bem_star_mensal import anexar_foto, materializar_fotos
def bs_mensal():
    A = ("id-qa", "CONDOMÍNIO FICTÍCIO BEM STAR", "11", "2026")
    fs = []; anexar_foto(fs, foto.read_bytes(), "foto.jpg", A); fs[0]["piscina"] = "Adulto"; fs[0]["legenda"] = "IMAGEM FICTÍCIA PARA TESTE"
    paths, meta = materializar_fotos(fs, A, str(OUT / "_bs"))
    lanc = [{"data": f"{d:02d}/11/2026", "piscinas": [{"nome": "Adulto", "tipo": "Adulto", "ph": "7.4", "cloro_livre": "2.1", "cloro_total": "2.4", "alcalinidade": "95"},
                                                     {"nome": "Infantil", "tipo": "Infantil", "ph": "7.6", "cloro_livre": "2.8"}]} for d in (3, 10, 17, 24)]
    ctx = {"cliente": A[1], "mes": "11", "ano": "2026", "lancamentos": lanc, "fotos": paths, "fotos_metadados": meta}
    ok, err = gerar_bs(OUT / "07_mensal_bem_star.pdf", ctx, cnpj="00.000.000/0000-00", responsavel="OPERADOR TESTE", logo_path=ROOT / "bem_star_logo.png")
    if not ok: raise RuntimeError(err)
registrar("07_mensal_bem_star.pdf", bs_mensal)

def docx_legado():
    doc = Document(ROOT / "relatorio_mensal.docx"); ns["atualizar_textos_normativos"](doc); doc.save(OUT / "08_relatorio_mensal_textos.docx")
registrar("08_relatorio_mensal_textos.docx", docx_legado)

pass

def premium_e_docx():
    analises = [
        {"data": "04/11/2026", "ph": "7,5", "cloro_livre": "2,5", "cloro_total": "2,8", "alcalinidade": "100", "dureza": "200", "cianurico": "40", "operador": "RT", "piscina": "Adulto"},
        {"data": "11/11/2026", "ph": "7,1", "cloro_livre": "0,8", "cloro_total": "1,6", "alcalinidade": "70", "dureza": "120", "cianurico": "80", "operador": "RT", "piscina": "Adulto"},
        {"data": "18/11/2026", "ph": "7,6", "cloro_livre": "3,0", "cloro_total": "2,5", "alcalinidade": "110", "dureza": "220", "cianurico": "45", "operador": "RT", "piscina": "Infantil"},
    ]
    aval = ns["avaliar_conformidade_analises"](analises)
    textos = ns["gerar_textos_automaticos_relatorio"](analises, aval)
    print("AVALIACAO:", aval["status"], aval["detalhes"])
    dados = {
        "empresa_rt": "ACQUA GESTÃO", "responsavel_tecnico": "Thyago Fernando da Silveira", "assinatura_rt_texto": "Thyago Fernando da Silveira",
        "crq": "CRQ-MG 2ª Região | CRQ 024025748", "qualificacao": "Técnico em Química", "certificacoes": "NR-26 | NR-6",
        "nome_condominio": COND, "cnpj_condominio": "00.000.000/0000-00", "endereco_condominio": "Rua Fictícia, 100",
        "representante": "SÍNDICO FICTÍCIO", "cpf_cnpj_representante": "000.000.000-00", "tipo_atendimento": "Contrato ativo",
        "mes_referencia": "11", "ano_referencia": "2026", "art_status": "Em processo de emissão", "art_numero": "N/A", "art_inicio": "N/A",
        "art_fim": "N/A", "art_texto": "", "data_emissao": "30/11/2026", "status_agua": aval["status"], "diagnostico": textos["diagnostico"],
        "analises": analises, "dosagens": [{"produto": "Cloro granulado", "fabricante_lote": "FICTÍCIO/L01", "quantidade": "1,2", "unidade": "kg", "finalidade": "Desinfecção"}],
        "recomendacoes": [{"recomendacao": t, "prazo": "Imediato", "responsavel": "Operação / RT"} for t in textos["recomendacoes"]],
        "observacoes": textos["observacoes"], "conformidades": {}, "epis": {}, "avaliacao_automatica": aval,
        "piscinas_cadastro": {"piscinas": [{"nome": "Adulto", "tipo": "Adulto", "volume": 120}, {"nome": "Infantil", "tipo": "Infantil", "volume": 15}], "volume_total": 135, "sem_distribuicao": False, "fonte": "cadastro"},
        "fotos_meta": {},
    }
    ok, err = ns["gerar_pdf_relatorio_rt_premium_reportlab"](dados, [foto], OUT / "09_mensal_rt_premium.pdf")
    if not ok: raise RuntimeError(err)
    ns["preencher_relatorio_mensal_docx"](ROOT / "relatorio_mensal.docx", OUT / "10_mensal_rt.docx", dados, fotos=[foto])
registrar("09_mensal_rt_premium.pdf + 10_mensal_rt.docx", premium_e_docx)
for n, r in resultados: print(f"{r:6} {n}")
sys.exit(0 if all(r == "OK" for _, r in resultados) else 1)
