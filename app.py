import os
import re
import json
import shutil
import platform
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

import streamlit as st
from docx import Document

# =========================================
# CONFIGURAÇÃO GERAL
# =========================================

APP_TITLE = "Aqua Gestão – Controle Técnico de Piscinas"
RESPONSAVEL_TECNICO = "Thyago Fernando da Silveira"
CRQ = "CRQ 024025748"
QUALIFICACAO_RT = "Técnico em Química"
CERTIFICACOES_RT = "NR-26 e NR-6"
EMPRESA_RT = "Aqua Gestão – Controle Técnico de Piscinas"

BASE_DIR = Path(__file__).resolve().parent
GENERATED_DIR = BASE_DIR / "generated"
TEMPLATE_CONTRATO = BASE_DIR / "template.docx"
TEMPLATE_ADITIVO = BASE_DIR / "aditivo.docx"
TEMPLATE_RELATORIO = BASE_DIR / "relatorio_mensal.docx"
DADOS_JSON_NAME = "dados_condominio.json"
MANIFEST_JSON_NAME = "manifest.json"

LOGO_CANDIDATOS = [
    BASE_DIR / "aqua_gestao_logo.png",
    BASE_DIR / "aqua_gestao_logo.jpg",
    BASE_DIR / "aqua_gestao_logo.jpeg",
    BASE_DIR / "logo.png",
    BASE_DIR / "logo.jpg",
    BASE_DIR / "logo.jpeg",
    BASE_DIR / "Aqua Gestão Logo.png",
    BASE_DIR / "Aqua_Gestao_Logo.png",
    BASE_DIR / "assets" / "aqua_gestao_logo.png",
    BASE_DIR / "assets" / "logo.png",
    BASE_DIR / "images" / "aqua_gestao_logo.png",
    BASE_DIR / "images" / "logo.png",
]

GENERATED_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title="Aqua Gestão RT",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================
# ESTILO VISUAL
# =========================================

st.markdown(
    """
    <style>
        .main .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1360px;
        }

        .top-card {
            border: 1px solid rgba(20, 85, 160, 0.18);
            border-radius: 18px;
            padding: 18px 22px;
            background: linear-gradient(135deg, #ffffff 0%, #f6fbff 100%);
            box-shadow: 0 10px 25px rgba(10, 50, 100, 0.07);
            margin-bottom: 18px;
        }

        .top-title {
            font-size: 1.7rem;
            font-weight: 700;
            color: #0d3d75;
            margin: 0;
            line-height: 1.15;
        }

        .top-subtitle {
            font-size: 0.95rem;
            color: #3b5d85;
            margin-top: 8px;
        }

        .info-badge {
            display: inline-block;
            padding: 7px 12px;
            border-radius: 999px;
            background: #edf5ff;
            color: #134b8a;
            border: 1px solid #d3e6ff;
            font-size: 0.88rem;
            margin-right: 8px;
            margin-top: 8px;
        }

        .section-card {
            border: 1px solid rgba(20, 85, 160, 0.16);
            border-radius: 18px;
            padding: 18px;
            background: #ffffff;
            box-shadow: 0 8px 20px rgba(10, 50, 100, 0.05);
            margin-bottom: 16px;
        }

        .history-meta {
            font-size: 0.84rem;
            color: #6a7d92;
            margin-top: 2px;
        }

        .confirm-box {
            border: 1px solid rgba(220, 80, 80, 0.35);
            border-radius: 12px;
            padding: 10px;
            margin-top: 8px;
            margin-bottom: 10px;
            background: rgba(220, 80, 80, 0.08);
        }

        .quick-mode-box {
            border: 1px solid rgba(20, 85, 160, 0.18);
            border-radius: 14px;
            padding: 12px 14px;
            background: #f6fbff;
            margin-bottom: 14px;
        }

        .alert-vigente {
            border: 1px solid rgba(40, 140, 80, 0.25);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(40, 140, 80, 0.08);
            margin-bottom: 14px;
        }

        .alert-vencendo {
            border: 1px solid rgba(220, 150, 20, 0.35);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(255, 190, 40, 0.12);
            margin-bottom: 14px;
        }

        .alert-vencido {
            border: 1px solid rgba(220, 60, 60, 0.35);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(220, 60, 60, 0.10);
            margin-bottom: 14px;
        }

        .status-badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 0.75rem;
            margin-top: 6px;
            border: 1px solid rgba(0,0,0,0.08);
        }

        .status-vigente { background: #ecfff3; color: #1d7a43; }
        .status-vencendo { background: #fff8e8; color: #9c6200; }
        .status-vencido { background: #fff0f0; color: #b42318; }
        .status-indefinido { background: #f4f6f8; color: #52606d; }

        .venc-row {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 16px;
            padding: 14px;
            background: linear-gradient(135deg, #ffffff 0%, #fbfdff 100%);
            margin-bottom: 12px;
        }

        .venc-nome {
            font-size: 1rem;
            font-weight: 700;
            color: #153f73;
            margin-bottom: 4px;
        }

        .venc-meta {
            font-size: 0.88rem;
            color: #5d7288;
            margin-bottom: 3px;
        }

        .venc-empty {
            border: 1px dashed rgba(20, 85, 160, 0.20);
            border-radius: 14px;
            padding: 14px;
            background: #f9fcff;
            color: #59708b;
        }

        .legacy-note {
            border: 1px solid rgba(120, 120, 120, 0.18);
            border-radius: 12px;
            padding: 10px 12px;
            background: #fafbfd;
            color: #5e6f82;
            font-size: 0.87rem;
            margin-top: 8px;
            margin-bottom: 8px;
        }

        .dash-mini {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 16px;
            padding: 14px;
            background: #fbfdff;
            min-height: 120px;
        }

        .dash-title {
            font-size: 0.88rem;
            color: #5f7590;
            margin-bottom: 6px;
        }

        .dash-value {
            font-size: 1.55rem;
            font-weight: 700;
            color: #103f78;
            line-height: 1.1;
        }

        .dash-sub {
            font-size: 0.82rem;
            color: #6d8197;
            margin-top: 8px;
        }

        .docs-note {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 12px;
            padding: 10px;
            background: #fbfdff;
            color: #5c7188;
            font-size: 0.84rem;
            margin-top: 8px;
        }

        .doc-chip {
            display: inline-block;
            padding: 5px 9px;
            border-radius: 999px;
            font-size: 0.75rem;
            margin-right: 6px;
            margin-top: 6px;
            background: #eef6ff;
            color: #174d87;
            border: 1px solid #d5e6fb;
        }

        .health-ok {
            color: #1d7a43;
            font-weight: 600;
        }

        .health-no {
            color: #b42318;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================
# FUNÇÕES AUXILIARES GERAIS
# =========================================

def encontrar_logo() -> Path | None:
    for caminho in LOGO_CANDIDATOS:
        if caminho.exists() and caminho.is_file():
            return caminho

    for extensao in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
        for pasta in [BASE_DIR, BASE_DIR / "assets", BASE_DIR / "images"]:
            if pasta.exists():
                encontrados = list(pasta.glob(extensao))
                for arq in encontrados:
                    if "logo" in arq.name.lower():
                        return arq
    return None


def slugify_nome(texto: str) -> str:
    texto = (texto or "").strip()
    texto = re.sub(r"[^\w\s-]", "", texto, flags=re.UNICODE)
    texto = re.sub(r"\s+", "_", texto)
    return texto[:120] if texto else "condominio"


def humanizar_nome_pasta(texto: str) -> str:
    texto = (texto or "").strip()
    texto = texto.replace("_", " ").replace("-", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def limpar_nome_arquivo(texto: str) -> str:
    texto = re.sub(r'[<>:"/\\\\|?*]+', "", texto)
    texto = re.sub(r"\s+", "_", texto.strip())
    return texto[:150]


def hoje_br() -> str:
    return date.today().strftime("%d/%m/%Y")


def agora_br() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def apenas_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def formatar_data_hora_arquivo(ts: float) -> str:
    dt = datetime.fromtimestamp(ts)
    return dt.strftime("%d/%m/%Y %H:%M")


def classificar_arquivo(nome_arquivo: str) -> tuple[str, str]:
    nome_lower = nome_arquivo.lower()

    if "relatorio" in nome_lower:
        tipo_doc = "Relatório"
    elif "contrato" in nome_lower:
        tipo_doc = "Contrato"
    elif "aditivo" in nome_lower:
        tipo_doc = "Aditivo"
    else:
        tipo_doc = "Documento"

    if nome_lower.endswith(".pdf"):
        tipo_ext = "PDF"
    elif nome_lower.endswith(".docx"):
        tipo_ext = "DOCX"
    else:
        tipo_ext = "Arquivo"

    return tipo_doc, tipo_ext


def chave_segura(texto: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_]+", "_", texto)


def parse_data_br(texto: str):
    try:
        return datetime.strptime((texto or "").strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def formatar_data_br(dt: date) -> str:
    return dt.strftime("%d/%m/%Y")


def adicionar_um_ano(dt: date) -> date:
    try:
        return dt.replace(year=dt.year + 1)
    except ValueError:
        return dt.replace(month=2, day=28, year=dt.year + 1)


def calcular_renovacao_anual(data_fim_texto: str):
    fim_atual = parse_data_br(data_fim_texto)
    if not fim_atual:
        return None, None

    novo_inicio = fim_atual + timedelta(days=1)
    novo_fim = adicionar_um_ano(novo_inicio) - timedelta(days=1)
    return novo_inicio, novo_fim


def status_vencimento(data_fim_texto: str, alerta_dias: int = 30):
    fim = parse_data_br(data_fim_texto)
    if not fim:
        return {
            "codigo": "indefinido",
            "rotulo": "Sem vigência válida",
            "mensagem": "Data final ausente ou inválida.",
            "dias": None,
            "css": "status-indefinido",
        }

    hoje = date.today()
    dias = (fim - hoje).days

    if dias < 0:
        return {
            "codigo": "vencido",
            "rotulo": "Vencido",
            "mensagem": f"Contrato vencido há {abs(dias)} dia(s).",
            "dias": dias,
            "css": "status-vencido",
        }

    if dias <= alerta_dias:
        return {
            "codigo": "vencendo",
            "rotulo": "Vence em breve",
            "mensagem": f"Contrato vence em {dias} dia(s).",
            "dias": dias,
            "css": "status-vencendo",
        }

    return {
        "codigo": "vigente",
        "rotulo": "Vigente",
        "mensagem": f"Contrato vigente. Restam {dias} dia(s) para o vencimento.",
        "dias": dias,
        "css": "status-vigente",
    }


def texto_dias_restantes(status: dict) -> str:
    dias = status.get("dias")
    if dias is None:
        return "Dias restantes: não disponível"
    if dias < 0:
        return f"Atrasado há {abs(dias)} dia(s)"
    return f"Restam {dias} dia(s)"


def sistema_e_windows() -> bool:
    return platform.system().lower().startswith("win")


def diagnostico_sistema() -> dict:
    return {
        "template_contrato_ok": TEMPLATE_CONTRATO.exists(),
        "template_aditivo_ok": TEMPLATE_ADITIVO.exists(),
        "template_relatorio_ok": TEMPLATE_RELATORIO.exists(),
        "generated_ok": GENERATED_DIR.exists(),
        "logo_ok": encontrar_logo() is not None,
        "windows_ok": sistema_e_windows(),
    }

# =========================================
# MÁSCARAS / FORMATAÇÃO
# =========================================

def formatar_cpf(texto: str) -> str:
    dig = apenas_digitos(texto)[:11]
    if len(dig) <= 3:
        return dig
    if len(dig) <= 6:
        return f"{dig[:3]}.{dig[3:]}"
    if len(dig) <= 9:
        return f"{dig[:3]}.{dig[3:6]}.{dig[6:]}"
    return f"{dig[:3]}.{dig[3:6]}.{dig[6:9]}-{dig[9:]}"


def formatar_cnpj(texto: str) -> str:
    dig = apenas_digitos(texto)[:14]
    if len(dig) <= 2:
        return dig
    if len(dig) <= 5:
        return f"{dig[:2]}.{dig[2:]}"
    if len(dig) <= 8:
        return f"{dig[:2]}.{dig[2:5]}.{dig[5:]}"
    if len(dig) <= 12:
        return f"{dig[:2]}.{dig[2:5]}.{dig[5:8]}/{dig[8:]}"
    return f"{dig[:2]}.{dig[2:5]}.{dig[5:8]}/{dig[8:12]}-{dig[12:]}"


def formatar_telefone(texto: str) -> str:
    dig = apenas_digitos(texto)

    if dig.startswith("55") and len(dig) > 11:
        dig = dig[2:]

    dig = dig[:11]

    if len(dig) <= 2:
        return dig
    if len(dig) <= 6:
        return f"({dig[:2]}) {dig[2:]}"
    if len(dig) <= 10:
        return f"({dig[:2]}) {dig[2:6]}-{dig[6:]}"
    return f"({dig[:2]}) {dig[2:7]}-{dig[7:]}"


def formatar_data_digitada(texto: str) -> str:
    dig = apenas_digitos(texto)[:8]
    if len(dig) <= 2:
        return dig
    if len(dig) <= 4:
        return f"{dig[:2]}/{dig[2:]}"
    return f"{dig[:2]}/{dig[2:4]}/{dig[4:]}"


def moeda_br_sem_prefixo(texto: str) -> str:
    if not texto:
        return ""

    dig = apenas_digitos(str(texto))
    if not dig:
        return ""

    if len(dig) == 1:
        valor = float(f"0.0{dig}")
    elif len(dig) == 2:
        valor = float(f"0.{dig}")
    else:
        valor = float(f"{dig[:-2]}.{dig[-2:]}")

    inteiro = int(valor)
    centavos = int(round((valor - inteiro) * 100))
    inteiro_fmt = f"{inteiro:,}".replace(",", ".")
    return f"{inteiro_fmt},{centavos:02d}"


def moeda_br(texto: str) -> str:
    fmt = moeda_br_sem_prefixo(texto)
    return f"R$ {fmt}" if fmt else ""


def valor_para_template(texto: str) -> str:
    texto = (texto or "").strip()
    if texto.startswith("R$"):
        return texto
    fmt = moeda_br_sem_prefixo(texto)
    return f"R$ {fmt}" if fmt else ""


def on_change_cpf():
    st.session_state.cpf_sindico = formatar_cpf(st.session_state.get("cpf_sindico", ""))


def on_change_cnpj():
    st.session_state.cnpj_condominio = formatar_cnpj(st.session_state.get("cnpj_condominio", ""))


def on_change_whatsapp():
    st.session_state.whatsapp_cliente = formatar_telefone(st.session_state.get("whatsapp_cliente", ""))


def on_change_valor_mensal():
    st.session_state.valor_mensal = moeda_br(st.session_state.get("valor_mensal", ""))


def on_change_valor_aditivo():
    st.session_state.valor_aditivo = moeda_br(st.session_state.get("valor_aditivo", ""))


def on_change_data_inicio():
    st.session_state.data_inicio = formatar_data_digitada(st.session_state.get("data_inicio", ""))


def on_change_data_fim():
    st.session_state.data_fim = formatar_data_digitada(st.session_state.get("data_fim", ""))


def on_change_data_assinatura():
    st.session_state.data_assinatura = formatar_data_digitada(st.session_state.get("data_assinatura", ""))

# =========================================
# VALIDAÇÕES REAIS
# =========================================

def validar_cpf(cpf: str) -> bool:
    cpf = apenas_digitos(cpf)

    if len(cpf) != 11:
        return False
    if cpf == cpf[0] * 11:
        return False

    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1
    if dig1 != int(cpf[9]):
        return False

    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2
    return dig2 == int(cpf[10])


def validar_cnpj(cnpj: str) -> bool:
    cnpj = apenas_digitos(cnpj)

    if len(cnpj) != 14:
        return False
    if cnpj == cnpj[0] * 14:
        return False

    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma1 = sum(int(cnpj[i]) * pesos1[i] for i in range(12))
    resto1 = soma1 % 11
    dig1 = 0 if resto1 < 2 else 11 - resto1
    if dig1 != int(cnpj[12]):
        return False

    pesos2 = [6] + pesos1
    soma2 = sum(int(cnpj[i]) * pesos2[i] for i in range(13))
    resto2 = soma2 % 11
    dig2 = 0 if resto2 < 2 else 11 - resto2
    return dig2 == int(cnpj[13])


def validar_data_br(texto: str) -> bool:
    try:
        datetime.strptime(texto.strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False


def validar_email(email: str) -> bool:
    email = (email or "").strip()
    if not email:
        return True
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return re.match(padrao, email) is not None


def validar_campos_obrigatorios(dados: dict) -> list[str]:
    mapa = {
        "DATA_ASSINATURA": "Data de assinatura",
        "NOME_CONDOMINIO": "Nome do condomínio",
        "CNPJ_CONDOMINIO": "CNPJ do condomínio",
        "ENDERECO_CONDOMINIO": "Endereço do condomínio",
        "NOME_SINDICO": "Nome do síndico / representante",
        "CPF_SINDICO": "CPF do síndico / representante",
        "VALOR_MENSAL": "Valor mensal",
        "VALOR_ADITIVO": "Valor com desconto/aditivo",
        "DATA_INICIO": "Data de início",
        "DATA_FIM": "Data de fim",
    }

    faltando = []
    for chave, rotulo in mapa.items():
        if not (dados.get(chave) or "").strip():
            faltando.append(rotulo)

    return faltando


def validar_campos_formato(dados: dict, email_cliente: str) -> list[str]:
    erros = []

    if not validar_cpf(dados["CPF_SINDICO"]):
        erros.append("CPF do síndico/representante inválido.")
    if not validar_cnpj(dados["CNPJ_CONDOMINIO"]):
        erros.append("CNPJ do condomínio inválido.")
    if not validar_data_br(dados["DATA_ASSINATURA"]):
        erros.append("Data de assinatura inválida. Use o formato dd/mm/aaaa.")
    if not validar_data_br(dados["DATA_INICIO"]):
        erros.append("Data de início inválida. Use o formato dd/mm/aaaa.")
    if not validar_data_br(dados["DATA_FIM"]):
        erros.append("Data de fim inválida. Use o formato dd/mm/aaaa.")
    if email_cliente.strip() and not validar_email(email_cliente):
        erros.append("E-mail inválido.")

    dt_inicio = parse_data_br(dados["DATA_INICIO"])
    dt_fim = parse_data_br(dados["DATA_FIM"])
    if dt_inicio and dt_fim and dt_fim < dt_inicio:
        erros.append("A data de fim não pode ser anterior à data de início.")

    return erros

# =========================================
# MANIFEST / AUDITORIA DOCUMENTAL
# =========================================

def caminho_manifest(pasta_condominio: Path) -> Path:
    return pasta_condominio / MANIFEST_JSON_NAME


def carregar_manifest(pasta_condominio: Path) -> dict:
    caminho = caminho_manifest(pasta_condominio)
    if not caminho.exists():
        return {
            "condominio": "",
            "criado_em": agora_br(),
            "atualizado_em": agora_br(),
            "documentos": [],
        }
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if "documentos" not in dados or not isinstance(dados["documentos"], list):
            dados["documentos"] = []
        return dados
    except Exception:
        return {
            "condominio": "",
            "criado_em": agora_br(),
            "atualizado_em": agora_br(),
            "documentos": [],
        }


def salvar_manifest(pasta_condominio: Path, manifest: dict):
    manifest["atualizado_em"] = agora_br()
    with open(caminho_manifest(pasta_condominio), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)


def registrar_documento_manifest(
    pasta_condominio: Path,
    nome_condominio: str,
    tipo: str,
    arquivo_docx: Path | None,
    arquivo_pdf: Path | None,
    pdf_gerado: bool,
    erro_pdf: str | None,
    dados_utilizados: dict,
):
    manifest = carregar_manifest(pasta_condominio)
    manifest["condominio"] = nome_condominio

    registro = {
        "registrado_em": agora_br(),
        "tipo": tipo,
        "arquivo_docx": arquivo_docx.name if arquivo_docx and arquivo_docx.exists() else "",
        "arquivo_pdf": arquivo_pdf.name if arquivo_pdf and arquivo_pdf.exists() else "",
        "pdf_gerado": pdf_gerado,
        "erro_pdf": erro_pdf or "",
        "data_inicio": dados_utilizados.get("DATA_INICIO", ""),
        "data_fim": dados_utilizados.get("DATA_FIM", ""),
        "valor_mensal": dados_utilizados.get("VALOR_MENSAL", ""),
        "valor_aditivo": dados_utilizados.get("VALOR_ADITIVO", ""),
    }
    manifest["documentos"].append(registro)
    salvar_manifest(pasta_condominio, manifest)


def resumo_manifest(pasta_condominio: Path) -> dict:
    manifest = carregar_manifest(pasta_condominio)
    docs = manifest.get("documentos", [])

    contratos = [d for d in docs if d.get("tipo") == "Contrato"]
    aditivos = [d for d in docs if d.get("tipo") == "Aditivo"]
    relatorios = [d for d in docs if d.get("tipo") == "Relatório"]

    ultimo = docs[-1] if docs else None

    return {
        "total": len(docs),
        "contratos": len(contratos),
        "aditivos": len(aditivos),
        "relatorios": len(relatorios),
        "ultimo_registro": ultimo,
        "manifest_atualizado_em": manifest.get("atualizado_em", ""),
    }

# =========================================
# PERSISTÊNCIA DE DADOS DO CONDOMÍNIO
# =========================================

def salvar_snapshot_formulario() -> dict:
    return {
        "nome_condominio": (st.session_state.get("nome_condominio") or "").strip(),
        "cnpj_condominio": (st.session_state.get("cnpj_condominio") or "").strip(),
        "endereco_condominio": (st.session_state.get("endereco_condominio") or "").strip(),
        "nome_sindico": (st.session_state.get("nome_sindico") or "").strip(),
        "cpf_sindico": (st.session_state.get("cpf_sindico") or "").strip(),
        "valor_mensal": valor_para_template((st.session_state.get("valor_mensal") or "").strip()),
        "valor_aditivo": valor_para_template((st.session_state.get("valor_aditivo") or "").strip()),
        "data_inicio": (st.session_state.get("data_inicio") or "").strip(),
        "data_fim": (st.session_state.get("data_fim") or "").strip(),
        "data_assinatura": (st.session_state.get("data_assinatura") or "").strip(),
        "whatsapp_cliente": (st.session_state.get("whatsapp_cliente") or "").strip(),
        "email_cliente": (st.session_state.get("email_cliente") or "").strip(),
        "observacoes_internas": (st.session_state.get("observacoes_internas") or "").strip(),
        "salvo_em": agora_br(),
        "responsavel_tecnico": RESPONSAVEL_TECNICO,
        "crq": CRQ,
        "marca": APP_TITLE,
    }


def salvar_dados_condominio(pasta_condominio: Path, dados_para_salvar: dict):
    caminho_json = pasta_condominio / DADOS_JSON_NAME
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados_para_salvar, f, ensure_ascii=False, indent=2)


def carregar_dados_condominio(pasta_condominio: Path) -> dict | None:
    caminho_json = pasta_condominio / DADOS_JSON_NAME
    if not caminho_json.exists():
        return None
    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def aplicar_dados_no_formulario(dados_salvos: dict):
    st.session_state.nome_condominio = dados_salvos.get("nome_condominio", "")
    st.session_state.cnpj_condominio = dados_salvos.get("cnpj_condominio", "")
    st.session_state.endereco_condominio = dados_salvos.get("endereco_condominio", "")
    st.session_state.nome_sindico = dados_salvos.get("nome_sindico", "")
    st.session_state.cpf_sindico = dados_salvos.get("cpf_sindico", "")
    st.session_state.valor_mensal = dados_salvos.get("valor_mensal", "")
    st.session_state.valor_aditivo = dados_salvos.get("valor_aditivo", "")
    st.session_state.data_inicio = dados_salvos.get("data_inicio", "")
    st.session_state.data_fim = dados_salvos.get("data_fim", "")
    st.session_state.data_assinatura = dados_salvos.get("data_assinatura", hoje_br())
    st.session_state.whatsapp_cliente = dados_salvos.get("whatsapp_cliente", "")
    st.session_state.email_cliente = dados_salvos.get("email_cliente", "")
    st.session_state.observacoes_internas = dados_salvos.get("observacoes_internas", "")
    st.session_state.origem_dados_carregados = dados_salvos.get("nome_condominio", "")


def preparar_cadastro_legado(nome_pasta: str):
    limpar_formulario()
    st.session_state.nome_condominio = humanizar_nome_pasta(nome_pasta)
    st.session_state.origem_dados_carregados = f"{humanizar_nome_pasta(nome_pasta)} (cadastro criado de pasta antiga)"


def preparar_renovacao_no_formulario(dados_salvos: dict) -> tuple[bool, str]:
    data_fim_atual = dados_salvos.get("data_fim", "")
    novo_inicio, novo_fim = calcular_renovacao_anual(data_fim_atual)

    if not novo_inicio or not novo_fim:
        return False, "Não foi possível renovar. A data final salva está ausente ou inválida."

    aplicar_dados_no_formulario(dados_salvos)
    st.session_state.data_inicio = formatar_data_br(novo_inicio)
    st.session_state.data_fim = formatar_data_br(novo_fim)
    st.session_state.data_assinatura = hoje_br()
    st.session_state.origem_dados_carregados = dados_salvos.get("nome_condominio", "")
    return True, "Nova vigência anual preenchida no formulário."


def obter_pasta_atual_do_formulario() -> Path | None:
    nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    if not nome_condominio:
        return None
    return GENERATED_DIR / slugify_nome(nome_condominio)


def salvar_cadastro_sem_gerar_documentos() -> tuple[bool, str]:
    nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    if not nome_condominio:
        return False, "Informe ao menos o nome do condomínio para salvar o cadastro."

    dados_base = {
        "DATA_ASSINATURA": (st.session_state.get("data_assinatura") or "").strip(),
        "NOME_CONDOMINIO": nome_condominio,
        "CNPJ_CONDOMINIO": (st.session_state.get("cnpj_condominio") or "").strip(),
        "ENDERECO_CONDOMINIO": (st.session_state.get("endereco_condominio") or "").strip(),
        "NOME_SINDICO": (st.session_state.get("nome_sindico") or "").strip(),
        "CPF_SINDICO": (st.session_state.get("cpf_sindico") or "").strip(),
        "VALOR_MENSAL": valor_para_template((st.session_state.get("valor_mensal") or "").strip()),
        "VALOR_ADITIVO": valor_para_template((st.session_state.get("valor_aditivo") or "").strip()),
        "DATA_INICIO": (st.session_state.get("data_inicio") or "").strip(),
        "DATA_FIM": (st.session_state.get("data_fim") or "").strip(),
    }

    erros = validar_para_geracao(dados_base, (st.session_state.get("email_cliente") or "").strip())
    if erros:
        return False, "Corrija os campos antes de salvar o cadastro: " + " | ".join(erros)

    pasta_condominio = obter_pasta_atual_do_formulario()
    assert pasta_condominio is not None
    pasta_condominio.mkdir(parents=True, exist_ok=True)
    salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

    manifest = carregar_manifest(pasta_condominio)
    manifest["condominio"] = nome_condominio
    salvar_manifest(pasta_condominio, manifest)

    return True, f"Cadastro salvo/atualizado com sucesso em '{pasta_condominio.name}'."

# =========================================
# HISTÓRICO / PAINEL
# =========================================

def listar_arquivos_pasta(pasta: Path):
    if not pasta.exists():
        return []

    arquivos = []
    for p in pasta.iterdir():
        if p.is_file():
            if p.name in (DADOS_JSON_NAME, MANIFEST_JSON_NAME):
                continue

            tipo_doc, tipo_ext = classificar_arquivo(p.name)
            arquivos.append(
                {
                    "path": p,
                    "name": p.name,
                    "tipo_doc": tipo_doc,
                    "tipo_ext": tipo_ext,
                    "modificado_em": formatar_data_hora_arquivo(p.stat().st_mtime),
                    "ts": p.stat().st_mtime,
                }
            )

    return sorted(arquivos, key=lambda x: x["ts"], reverse=True)


def localizar_ultimo_documento(arquivos: list, tipo_doc: str, extensao_preferida: str = "PDF"):
    preferidos = [a for a in arquivos if a["tipo_doc"] == tipo_doc and a["tipo_ext"] == extensao_preferida]
    if preferidos:
        return preferidos[0]
    alternativos = [a for a in arquivos if a["tipo_doc"] == tipo_doc]
    return alternativos[0] if alternativos else None


def listar_historico():
    if not GENERATED_DIR.exists():
        return []

    pastas = [p for p in GENERATED_DIR.iterdir() if p.is_dir()]
    pastas = sorted(pastas, key=lambda p: p.stat().st_mtime, reverse=True)

    historico = []
    for pasta in pastas:
        arquivos = listar_arquivos_pasta(pasta)
        dados_salvos = carregar_dados_condominio(pasta)
        data_fim = dados_salvos.get("data_fim", "") if dados_salvos else ""
        status = status_vencimento(data_fim)
        resumo_docs = resumo_manifest(pasta)

        historico.append(
            {
                "nome": pasta.name,
                "pasta": pasta,
                "arquivos": arquivos,
                "total_arquivos": len(arquivos),
                "tem_dados_salvos": (pasta / DADOS_JSON_NAME).exists(),
                "status_vencimento": status,
                "data_fim": data_fim,
                "resumo_docs": resumo_docs,
            }
        )
    return historico


def listar_painel_vencimentos(alerta_dias: int):
    if not GENERATED_DIR.exists():
        return []

    pastas = [p for p in GENERATED_DIR.iterdir() if p.is_dir()]
    itens = []

    for pasta in pastas:
        dados_salvos = carregar_dados_condominio(pasta)
        arquivos = listar_arquivos_pasta(pasta)
        resumo_docs = resumo_manifest(pasta)

        nome_exibicao = pasta.name
        data_fim = ""
        status = {
            "codigo": "indefinido",
            "rotulo": "Sem vigência válida",
            "mensagem": "Data final ausente ou inválida.",
            "dias": None,
            "css": "status-indefinido",
        }
        origem = "legado_sem_json"

        if dados_salvos:
            nome_exibicao = dados_salvos.get("nome_condominio", pasta.name)
            data_fim = dados_salvos.get("data_fim", "")
            status = status_vencimento(data_fim, alerta_dias)
            origem = "json"

        ultimo_contrato = localizar_ultimo_documento(arquivos, "Contrato")
        ultimo_aditivo = localizar_ultimo_documento(arquivos, "Aditivo")
        ultimo_relatorio = localizar_ultimo_documento(arquivos, "Relatório")

        itens.append(
            {
                "nome_exibicao": nome_exibicao,
                "slug": pasta.name,
                "pasta": pasta,
                "dados": dados_salvos,
                "arquivos": arquivos,
                "data_fim": data_fim,
                "data_fim_dt": parse_data_br(data_fim),
                "status": status,
                "origem": origem,
                "tem_json": dados_salvos is not None,
                "ultimo_contrato": ultimo_contrato,
                "ultimo_aditivo": ultimo_aditivo,
                "ultimo_relatorio": ultimo_relatorio,
                "resumo_docs": resumo_docs,
            }
        )

    def ordem_item(item):
        status_codigo = item["status"]["codigo"]
        dt_fim = item["data_fim_dt"] or date.max
        dias = item["status"]["dias"]
        dias_sort = dias if dias is not None else 999999

        prioridade = {
            "vencido": 0,
            "vencendo": 1,
            "vigente": 2,
            "indefinido": 3,
        }.get(status_codigo, 9)

        return (prioridade, dt_fim, dias_sort, item["nome_exibicao"].lower())

    return sorted(itens, key=ordem_item)


def filtrar_itens_painel(itens: list, termo: str, filtro_status: str):
    resultado = itens

    termo = (termo or "").strip().lower()
    if termo:
        filtrados = []
        for item in resultado:
            dados = item["dados"] or {}
            campos = [
                item["nome_exibicao"],
                item["slug"],
                dados.get("nome_condominio", ""),
                dados.get("cnpj_condominio", ""),
                dados.get("nome_sindico", ""),
                item["status"]["rotulo"],
            ]
            texto_unico = " ".join(campos).lower()
            if termo in texto_unico:
                filtrados.append(item)
        resultado = filtrados

    mapa_status = {
        "Todos": None,
        "Vigente": "vigente",
        "Vence em breve": "vencendo",
        "Vencido": "vencido",
        "Sem vigência válida": "indefinido",
    }

    codigo = mapa_status.get(filtro_status)
    if codigo:
        resultado = [i for i in resultado if i["status"]["codigo"] == codigo]

    return resultado


def abrir_pasta_windows(caminho: Path):
    try:
        if not caminho.exists():
            st.error("A pasta não foi encontrada.")
            return
        if not sistema_e_windows():
            st.error("Esta função de abertura direta foi preparada para ambiente Windows.")
            return
        os.startfile(str(caminho))
    except Exception as e:
        st.error(f"Não foi possível abrir a pasta no Windows: {e}")


def abrir_arquivo_windows(caminho: Path):
    try:
        if not caminho.exists():
            st.error("O arquivo não foi encontrado.")
            return
        if not sistema_e_windows():
            st.error("Esta função de abertura direta foi preparada para ambiente Windows.")
            return
        os.startfile(str(caminho))
    except Exception as e:
        st.error(f"Não foi possível abrir o arquivo: {e}")


def deletar_pasta_condominio(pasta: Path):
    if pasta.exists() and pasta.is_dir():
        shutil.rmtree(pasta)


def deletar_arquivo_individual(arquivo: Path):
    pasta = arquivo.parent
    if arquivo.exists() and arquivo.is_file():
        arquivo.unlink()

    if pasta.exists() and pasta.is_dir():
        restantes = list(pasta.iterdir())
        if len(restantes) == 0:
            pasta.rmdir()

# =========================================
# DOCX / PLACEHOLDERS
# =========================================

def substituir_em_paragrafo(paragraph, mapa: dict[str, str]):
    texto_original = paragraph.text
    texto_novo = texto_original
    alterou = False

    for chave, valor in mapa.items():
        if chave in texto_novo:
            texto_novo = texto_novo.replace(chave, valor)
            alterou = True

    if alterou:
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = texto_novo
        else:
            paragraph.add_run(texto_novo)


def substituir_placeholders_doc(doc: Document, mapa: dict[str, str]):
    for paragraph in doc.paragraphs:
        substituir_em_paragrafo(paragraph, mapa)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    substituir_em_paragrafo(paragraph, mapa)

    for section in doc.sections:
        header = section.header
        footer = section.footer

        for paragraph in header.paragraphs:
            substituir_em_paragrafo(paragraph, mapa)

        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        substituir_em_paragrafo(paragraph, mapa)

        for paragraph in footer.paragraphs:
            substituir_em_paragrafo(paragraph, mapa)

        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        substituir_em_paragrafo(paragraph, mapa)


def adicionar_espaco(doc: Document, qtd: int = 1):
    for _ in range(qtd):
        doc.add_paragraph("")


def adicionar_bloco_assinaturas(doc: Document, nome_sindico: str):
    adicionar_espaco(doc, 2)

    p_local = doc.add_paragraph()
    p_local.alignment = 1
    p_local.add_run(f"Uberlândia/MG, {hoje_br()}.")

    adicionar_espaco(doc, 2)

    tabela = doc.add_table(rows=2, cols=2)
    tabela.autofit = True

    cell_00 = tabela.cell(0, 0)
    cell_01 = tabela.cell(0, 1)
    cell_10 = tabela.cell(1, 0)
    cell_11 = tabela.cell(1, 1)

    cell_00.text = "__________________________________\nAQUA GESTÃO – CONTROLE TÉCNICO DE PISCINAS\nCONTRATADA"
    cell_01.text = f"__________________________________\n{nome_sindico}\nCONTRATANTE"

    cell_10.text = "__________________________________\nTestemunha 1\nNome:\nCPF:"
    cell_11.text = "__________________________________\nTestemunha 2\nNome:\nCPF:"

    for row in tabela.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = 1


def converter_docx_para_pdf(docx_path: Path, pdf_path: Path):
    try:
        if not sistema_e_windows():
            return False, "Conversão DOCX -> PDF configurada para uso no Windows com Word."

        if not docx_path.exists():
            return False, "Arquivo DOCX não encontrado para conversão."

        import pythoncom
        from docx2pdf import convert

        pythoncom.CoInitialize()
        convert(str(docx_path), str(pdf_path))

        if not pdf_path.exists():
            return False, "Conversão executada, mas o PDF não foi localizado ao final."

        return True, None
    except Exception as e:
        return False, str(e)


def gerar_documento(
    template_path: Path,
    output_docx: Path,
    placeholders: dict[str, str],
    incluir_assinaturas: bool = True,
    nome_sindico: str = "",
):
    if not template_path.exists():
        raise FileNotFoundError(f"Template não encontrado: {template_path.name}")

    doc = Document(str(template_path))
    substituir_placeholders_doc(doc, placeholders)

    if incluir_assinaturas:
        adicionar_bloco_assinaturas(doc, nome_sindico=nome_sindico)

    doc.save(str(output_docx))

# =========================================
# RELATÓRIO MENSAL DE RT
# =========================================

def normalizar_texto(texto: str) -> str:
    texto = (texto or "").lower()
    mapa = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüç", "aaaaaeeeeiiiiooooouuuuc")
    return texto.translate(mapa)


def set_cell_text(cell, texto: str):
    texto = texto if texto is not None else ""
    if cell.paragraphs:
        primeira = True
        for p in cell.paragraphs:
            for run in p.runs:
                run.text = ""
            if primeira:
                if p.runs:
                    p.runs[0].text = str(texto)
                else:
                    p.add_run(str(texto))
                primeira = False
    else:
        cell.text = str(texto)


def encontrar_tabela_por_keywords(doc: Document, keywords: list[str]):
    kws = [normalizar_texto(k) for k in keywords]
    melhor = None
    melhor_score = -1
    for table in doc.tables:
        conteudo = " ".join(normalizar_texto(c.text) for row in table.rows[:3] for c in row.cells)
        score = sum(1 for k in kws if k in conteudo)
        if score > melhor_score:
            melhor = table
            melhor_score = score
    return melhor if melhor_score > 0 else None


def preencher_tabela_generica(table, rows_data: list[list[str]], start_row: int = 1):
    if table is None:
        return False
    total_cols = max((len(r.cells) for r in table.rows), default=0)
    if total_cols == 0:
        return False
    for idx, row_data in enumerate(rows_data, start=start_row):
        if idx >= len(table.rows):
            table.add_row()
        row = table.rows[idx]
        for c in range(total_cols):
            valor = row_data[c] if c < len(row_data) else ""
            set_cell_text(row.cells[c], valor)
    return True


def substituir_rotulo_em_paragrafos(container, mapa_rotulos: dict[str, str]) -> int:
    alteracoes = 0
    paragrafos = []
    if hasattr(container, "paragraphs"):
        paragrafos.extend(container.paragraphs)
    if hasattr(container, "tables"):
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    paragrafos.extend(cell.paragraphs)
    for p in paragrafos:
        txt_norm = normalizar_texto(p.text)
        for rotulo, valor in mapa_rotulos.items():
            if normalizar_texto(rotulo) in txt_norm and valor:
                novo = f"{rotulo}: {valor}"
                for run in p.runs:
                    run.text = ""
                if p.runs:
                    p.runs[0].text = novo
                else:
                    p.add_run(novo)
                alteracoes += 1
                break
    return alteracoes


def coletar_analises_relatorio() -> list[dict]:
    itens = []
    for i in range(10):
        itens.append({
            "data": (st.session_state.get(f"rel_analise_data_{i}") or "").strip(),
            "ph": (st.session_state.get(f"rel_analise_ph_{i}") or "").strip(),
            "cloro_livre": (st.session_state.get(f"rel_analise_cl_{i}") or "").strip(),
            "cloro_total": (st.session_state.get(f"rel_analise_ct_{i}") or "").strip(),
            "alcalinidade": (st.session_state.get(f"rel_analise_alc_{i}") or "").strip(),
            "dureza": (st.session_state.get(f"rel_analise_dc_{i}") or "").strip(),
            "cianurico": (st.session_state.get(f"rel_analise_cya_{i}") or "").strip(),
            "operador": (st.session_state.get(f"rel_analise_operador_{i}") or "").strip(),
        })
    return itens


def coletar_dosagens_relatorio() -> list[dict]:
    itens = []
    for i in range(7):
        itens.append({
            "produto": (st.session_state.get(f"rel_dos_produto_{i}") or "").strip(),
            "fabricante_lote": (st.session_state.get(f"rel_dos_lote_{i}") or "").strip(),
            "quantidade": (st.session_state.get(f"rel_dos_qtd_{i}") or "").strip(),
            "unidade": (st.session_state.get(f"rel_dos_un_{i}") or "").strip(),
            "finalidade": (st.session_state.get(f"rel_dos_finalidade_{i}") or "").strip(),
        })
    return itens


def coletar_recomendacoes_relatorio() -> list[dict]:
    itens = []
    for i in range(5):
        itens.append({
            "recomendacao": (st.session_state.get(f"rel_rec_texto_{i}") or "").strip(),
            "prazo": (st.session_state.get(f"rel_rec_prazo_{i}") or "").strip(),
            "responsavel": (st.session_state.get(f"rel_rec_resp_{i}") or "").strip(),
        })
    return itens


def coletar_conformidades_relatorio() -> dict:
    return {
        "nbr_11238": (st.session_state.get("rel_nbr_11238") or "").strip(),
        "nr_26": (st.session_state.get("rel_nr_26") or "").strip(),
        "nr_06": (st.session_state.get("rel_nr_06") or "").strip(),
    }


def montar_dados_relatorio() -> dict:
    nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    representante = (st.session_state.get("nome_sindico") or "").strip()
    dados_base = salvar_snapshot_formulario()
    return {
        "empresa_rt": EMPRESA_RT,
        "responsavel_tecnico": RESPONSAVEL_TECNICO,
        "crq": CRQ,
        "qualificacao": QUALIFICACAO_RT,
        "certificacoes": CERTIFICACOES_RT,
        "nome_condominio": nome_condominio,
        "cnpj_condominio": dados_base.get("cnpj_condominio", ""),
        "endereco_condominio": dados_base.get("endereco_condominio", ""),
        "representante": representante,
        "mes_referencia": (st.session_state.get("rel_mes_referencia") or "").strip(),
        "ano_referencia": (st.session_state.get("rel_ano_referencia") or "").strip(),
        "art_numero": (st.session_state.get("rel_art_numero") or "").strip(),
        "art_inicio": (st.session_state.get("rel_art_inicio") or "").strip(),
        "art_fim": (st.session_state.get("rel_art_fim") or "").strip(),
        "data_emissao": (st.session_state.get("rel_data_emissao") or hoje_br()).strip(),
        "status_agua": (st.session_state.get("rel_status_agua") or "CONFORME").strip(),
        "diagnostico": (st.session_state.get("rel_diagnostico") or "").strip(),
        "analises": coletar_analises_relatorio(),
        "dosagens": coletar_dosagens_relatorio(),
        "recomendacoes": coletar_recomendacoes_relatorio(),
        "conformidades": coletar_conformidades_relatorio(),
    }


def validar_relatorio_mensal(dados_relatorio: dict) -> list[str]:
    erros = []
    if not dados_relatorio.get("nome_condominio"):
        erros.append("Informe o nome do condomínio no cadastro antes de gerar o relatório.")
    if not dados_relatorio.get("mes_referencia"):
        erros.append("Informe o mês de referência do relatório.")
    if not dados_relatorio.get("ano_referencia"):
        erros.append("Informe o ano de referência do relatório.")
    if not TEMPLATE_RELATORIO.exists():
        erros.append("O arquivo relatorio_mensal.docx não foi localizado na pasta do projeto.")
    return erros


def append_relatorio_fallback(doc: Document, dados_relatorio: dict):
    doc.add_page_break()
    doc.add_paragraph("COMPLEMENTO AUTOMÁTICO – DADOS ESTRUTURADOS DO RELATÓRIO MENSAL")
    doc.add_paragraph(f"Condomínio: {dados_relatorio['nome_condominio']}")
    doc.add_paragraph(f"Mês/Ano de referência: {dados_relatorio['mes_referencia']}/{dados_relatorio['ano_referencia']}")
    doc.add_paragraph(f"ART nº: {dados_relatorio['art_numero']}")
    doc.add_paragraph(f"Vigência ART: {dados_relatorio['art_inicio']} até {dados_relatorio['art_fim']}")
    doc.add_paragraph(f"Data de emissão: {dados_relatorio['data_emissao']}")
    doc.add_paragraph(f"Responsável técnico: {dados_relatorio['responsavel_tecnico']} – {dados_relatorio['qualificacao']} – {dados_relatorio['crq']}")
    doc.add_paragraph(f"Certificações relevantes: {dados_relatorio['certificacoes']}")
    doc.add_paragraph(f"Status geral da água: {dados_relatorio['status_agua']}")
    doc.add_paragraph(f"Diagnóstico técnico: {dados_relatorio['diagnostico']}")

    doc.add_paragraph("ANÁLISES FÍSICO-QUÍMICAS")
    t1 = doc.add_table(rows=1, cols=8)
    headers1 = ["Data", "pH", "Cloro Livre", "Cloro Total", "Alcalinidade", "Dureza Cálcica", "Ácido Cianúrico", "Operador"]
    for i, h in enumerate(headers1):
        set_cell_text(t1.rows[0].cells[i], h)
    linhas = []
    for a in dados_relatorio["analises"]:
        linhas.append([a["data"], a["ph"], a["cloro_livre"], a["cloro_total"], a["alcalinidade"], a["dureza"], a["cianurico"], a["operador"]])
    preencher_tabela_generica(t1, linhas, start_row=1)

    doc.add_paragraph("DOSAGENS DE PRODUTOS QUÍMICOS")
    t2 = doc.add_table(rows=1, cols=5)
    headers2 = ["Produto químico", "Fabricante/Lote", "Quantidade", "Unidade", "Finalidade técnica"]
    for i, h in enumerate(headers2):
        set_cell_text(t2.rows[0].cells[i], h)
    linhas = []
    for d in dados_relatorio["dosagens"]:
        linhas.append([d["produto"], d["fabricante_lote"], d["quantidade"], d["unidade"], d["finalidade"]])
    preencher_tabela_generica(t2, linhas, start_row=1)

    doc.add_paragraph("RECOMENDAÇÕES TÉCNICAS")
    t3 = doc.add_table(rows=1, cols=3)
    headers3 = ["Recomendação", "Prazo", "Responsável"]
    for i, h in enumerate(headers3):
        set_cell_text(t3.rows[0].cells[i], h)
    linhas = []
    for r in dados_relatorio["recomendacoes"]:
        linhas.append([r["recomendacao"], r["prazo"], r["responsavel"]])
    preencher_tabela_generica(t3, linhas, start_row=1)

    doc.add_paragraph("CONFORMIDADE E SEGURANÇA")
    doc.add_paragraph(f"NBR 11238: {dados_relatorio['conformidades']['nbr_11238']}")
    doc.add_paragraph(f"NR-26 / GHS: {dados_relatorio['conformidades']['nr_26']}")
    doc.add_paragraph(f"NR-06 / EPI: {dados_relatorio['conformidades']['nr_06']}")


def preencher_relatorio_mensal_docx(template_path: Path, output_docx: Path, dados_relatorio: dict):
    doc = Document(str(template_path))

    placeholders = {
        "{{NOME_CONDOMINIO}}": dados_relatorio["nome_condominio"],
        "{{CNPJ_CONDOMINIO}}": dados_relatorio["cnpj_condominio"],
        "{{ENDERECO_CONDOMINIO}}": dados_relatorio["endereco_condominio"],
        "{{NOME_SINDICO}}": dados_relatorio["representante"],
        "{{RESPONSAVEL_TECNICO}}": dados_relatorio["responsavel_tecnico"],
        "{{CRQ}}": dados_relatorio["crq"],
        "{{QUALIFICACAO_RT}}": dados_relatorio["qualificacao"],
        "{{CERTIFICACOES_RT}}": dados_relatorio["certificacoes"],
        "{{EMPRESA_RT}}": dados_relatorio["empresa_rt"],
        "{{MES_REFERENCIA}}": dados_relatorio["mes_referencia"],
        "{{ANO_REFERENCIA}}": dados_relatorio["ano_referencia"],
        "{{ART_NUMERO}}": dados_relatorio["art_numero"],
        "{{ART_INICIO}}": dados_relatorio["art_inicio"],
        "{{ART_FIM}}": dados_relatorio["art_fim"],
        "{{DATA_EMISSAO}}": dados_relatorio["data_emissao"],
        "{{STATUS_AGUA}}": dados_relatorio["status_agua"],
        "{{DIAGNOSTICO_TECNICO}}": dados_relatorio["diagnostico"],
        "{{NBR_11238}}": dados_relatorio["conformidades"]["nbr_11238"],
        "{{NR_26}}": dados_relatorio["conformidades"]["nr_26"],
        "{{NR_06}}": dados_relatorio["conformidades"]["nr_06"],
    }
    substituir_placeholders_doc(doc, placeholders)

    substituir_rotulo_em_paragrafos(doc, {
        "Mês de referência": dados_relatorio["mes_referencia"],
        "Ano de referência": dados_relatorio["ano_referencia"],
        "ART nº": dados_relatorio["art_numero"],
        "Vigência da ART": f"{dados_relatorio['art_inicio']} a {dados_relatorio['art_fim']}",
        "Data de emissão": dados_relatorio["data_emissao"],
        "Status geral da água": dados_relatorio["status_agua"],
        "Parecer técnico": dados_relatorio["diagnostico"],
        "Diagnóstico técnico": dados_relatorio["diagnostico"],
        "NBR 11238": dados_relatorio["conformidades"]["nbr_11238"],
        "NR-26": dados_relatorio["conformidades"]["nr_26"],
        "NR-06": dados_relatorio["conformidades"]["nr_06"],
        "Responsável Técnico": f"{dados_relatorio['responsavel_tecnico']} - {dados_relatorio['crq']}",
        "Representante do estabelecimento": dados_relatorio["representante"],
    })

    analises_rows = [
        [a["data"], a["ph"], a["cloro_livre"], a["cloro_total"], a["alcalinidade"], a["dureza"], a["cianurico"], a["operador"]]
        for a in dados_relatorio["analises"]
    ]
    dosagens_rows = [
        [d["produto"], d["fabricante_lote"], d["quantidade"], d["unidade"], d["finalidade"]]
        for d in dados_relatorio["dosagens"]
    ]
    recomendacoes_rows = [
        [r["recomendacao"], r["prazo"], r["responsavel"]]
        for r in dados_relatorio["recomendacoes"]
    ]

    tabela_analises = encontrar_tabela_por_keywords(doc, ["ph", "cloro", "alcalinidade", "cianurico", "operador"])
    tabela_dosagens = encontrar_tabela_por_keywords(doc, ["produto", "lote", "quantidade", "unidade", "finalidade"])
    tabela_recomendacoes = encontrar_tabela_por_keywords(doc, ["recomendacao", "prazo", "responsavel"])

    preencheu_alguma = False
    if tabela_analises:
        preencheu_alguma = preencher_tabela_generica(tabela_analises, analises_rows, start_row=1) or preencheu_alguma
    if tabela_dosagens:
        preencheu_alguma = preencher_tabela_generica(tabela_dosagens, dosagens_rows, start_row=1) or preencheu_alguma
    if tabela_recomendacoes:
        preencheu_alguma = preencher_tabela_generica(tabela_recomendacoes, recomendacoes_rows, start_row=1) or preencheu_alguma

    if not preencheu_alguma:
        append_relatorio_fallback(doc, dados_relatorio)

    doc.save(str(output_docx))


def gerar_relatorio_mensal() -> tuple[bool, str]:
    dados_relatorio = montar_dados_relatorio()
    erros = validar_relatorio_mensal(dados_relatorio)
    if erros:
        return False, " | ".join(erros)

    nome_condominio = dados_relatorio["nome_condominio"]
    pasta_condominio = GENERATED_DIR / slugify_nome(nome_condominio)
    pasta_condominio.mkdir(parents=True, exist_ok=True)
    salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_nome = limpar_nome_arquivo(
        f"Relatorio_Mensal_RT_{nome_condominio}_{dados_relatorio['mes_referencia']}_{dados_relatorio['ano_referencia']}_{timestamp}"
    )
    relatorio_docx = pasta_condominio / f"{base_nome}.docx"
    relatorio_pdf = pasta_condominio / f"{base_nome}.pdf"

    preencher_relatorio_mensal_docx(TEMPLATE_RELATORIO, relatorio_docx, dados_relatorio)
    ok_pdf, erro_pdf = converter_docx_para_pdf(relatorio_docx, relatorio_pdf)

    registrar_documento_manifest(
        pasta_condominio=pasta_condominio,
        nome_condominio=nome_condominio,
        tipo="Relatório",
        arquivo_docx=relatorio_docx,
        arquivo_pdf=relatorio_pdf,
        pdf_gerado=ok_pdf,
        erro_pdf=erro_pdf,
        dados_utilizados={
            "DATA_INICIO": (st.session_state.get("data_inicio") or "").strip(),
            "DATA_FIM": (st.session_state.get("data_fim") or "").strip(),
            "VALOR_MENSAL": valor_para_template((st.session_state.get("valor_mensal") or "").strip()),
            "VALOR_ADITIVO": valor_para_template((st.session_state.get("valor_aditivo") or "").strip()),
        },
    )

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Relatório mensal gerado")
    c1, c2, c3 = st.columns(3)
    with c1:
        if relatorio_docx.exists():
            with open(relatorio_docx, "rb") as f:
                st.download_button(
                    "Baixar DOCX do relatório",
                    data=f,
                    file_name=relatorio_docx.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
    with c2:
        if ok_pdf and relatorio_pdf.exists():
            with open(relatorio_pdf, "rb") as f:
                st.download_button(
                    "Baixar PDF do relatório",
                    data=f,
                    file_name=relatorio_pdf.name,
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.warning(f"PDF do relatório não gerado. Erro: {erro_pdf}")
    with c3:
        if st.button("Abrir pasta do condomínio", key="abrir_pasta_relatorio", use_container_width=True):
            abrir_pasta_windows(pasta_condominio)
    st.markdown('</div>', unsafe_allow_html=True)

    return True, f"Relatório mensal registrado com sucesso para {nome_condominio}."

# =========================================
# EXPORTAÇÃO DE CADASTRO
# =========================================

def gerar_html_resumo_cadastro(item: dict) -> str:
    dados = item["dados"] or {}
    status = item["status"]
    nome = item["nome_exibicao"]
    resumo_docs = item.get("resumo_docs", {})

    def val(chave):
        return dados.get(chave, "Não informado") or "Não informado"

    ultimo_registro = resumo_docs.get("ultimo_registro") or {}

    html = f"""
    <html>
    <head>
        <meta charset="utf-8" />
        <title>Resumo de Cadastro - {nome}</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 28px;
                color: #15395f;
            }}
            h1 {{
                font-size: 24px;
                color: #0d3d75;
                margin-bottom: 6px;
            }}
            h2 {{
                font-size: 16px;
                color: #25598d;
                margin-top: 24px;
                margin-bottom: 10px;
            }}
            .sub {{
                color: #5e7691;
                margin-bottom: 20px;
            }}
            .box {{
                border: 1px solid #d7e6f7;
                border-radius: 12px;
                padding: 14px 16px;
                margin-bottom: 12px;
                background: #fbfdff;
            }}
            .line {{
                margin-bottom: 6px;
            }}
            .label {{
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <h1>{nome}</h1>
        <div class="sub">Resumo de cadastro exportado pelo sistema Aqua Gestão</div>

        <div class="box">
            <div class="line"><span class="label">Status:</span> {status['rotulo']}</div>
            <div class="line"><span class="label">Mensagem:</span> {status['mensagem']}</div>
            <div class="line"><span class="label">Data final:</span> {item.get('data_fim') or 'Não informada'}</div>
        </div>

        <h2>Dados do condomínio</h2>
        <div class="box">
            <div class="line"><span class="label">Nome do condomínio:</span> {val('nome_condominio')}</div>
            <div class="line"><span class="label">CNPJ:</span> {val('cnpj_condominio')}</div>
            <div class="line"><span class="label">Endereço:</span> {val('endereco_condominio')}</div>
            <div class="line"><span class="label">Síndico / representante:</span> {val('nome_sindico')}</div>
            <div class="line"><span class="label">CPF:</span> {val('cpf_sindico')}</div>
        </div>

        <h2>Dados contratuais</h2>
        <div class="box">
            <div class="line"><span class="label">Valor mensal:</span> {val('valor_mensal')}</div>
            <div class="line"><span class="label">Valor aditivo:</span> {val('valor_aditivo')}</div>
            <div class="line"><span class="label">Data de início:</span> {val('data_inicio')}</div>
            <div class="line"><span class="label">Data de fim:</span> {val('data_fim')}</div>
            <div class="line"><span class="label">Data de assinatura:</span> {val('data_assinatura')}</div>
        </div>

        <h2>Contato</h2>
        <div class="box">
            <div class="line"><span class="label">WhatsApp:</span> {val('whatsapp_cliente')}</div>
            <div class="line"><span class="label">E-mail:</span> {val('email_cliente')}</div>
            <div class="line"><span class="label">Última atualização cadastral:</span> {val('salvo_em')}</div>
        </div>

        <h2>Controle documental</h2>
        <div class="box">
            <div class="line"><span class="label">Total de registros documentais:</span> {resumo_docs.get('total', 0)}</div>
            <div class="line"><span class="label">Contratos gerados:</span> {resumo_docs.get('contratos', 0)}</div>
            <div class="line"><span class="label">Aditivos gerados:</span> {resumo_docs.get('aditivos', 0)}</div>
            <div class="line"><span class="label">Relatórios gerados:</span> {resumo_docs.get('relatorios', 0)}</div>
            <div class="line"><span class="label">Último registro:</span> {ultimo_registro.get('registrado_em', 'Não informado')}</div>
            <div class="line"><span class="label">Último tipo:</span> {ultimo_registro.get('tipo', 'Não informado')}</div>
        </div>

        <h2>Observações internas</h2>
        <div class="box">
            <div class="line">{val('observacoes_internas')}</div>
        </div>
    </body>
    </html>
    """
    return html

# =========================================
# MENSAGENS / LINKS
# =========================================

def montar_mensagem_envio(
    nome_condominio: str,
    nome_sindico: str,
    caminho_contrato_pdf: Path | None = None,
    caminho_aditivo_pdf: Path | None = None,
) -> str:
    partes = [
        f"Prezado(a) {nome_sindico},",
        "",
        f"Segue em anexo a documentação referente ao condomínio {nome_condominio}:",
        "",
    ]

    if caminho_contrato_pdf and caminho_contrato_pdf.exists():
        partes.append("- Contrato de Responsabilidade Técnica (PDF)")
    if caminho_aditivo_pdf and caminho_aditivo_pdf.exists():
        partes.append("- Aditivo contratual (PDF)")

    partes += [
        "",
        "Permaneço à disposição para quaisquer esclarecimentos.",
        "",
        "Atenciosamente,",
        f"{RESPONSAVEL_TECNICO}",
        CRQ,
        "Aqua Gestão – Controle Técnico de Piscinas",
    ]

    return "\n".join(partes)


def link_whatsapp(telefone: str, mensagem: str) -> str:
    somente_numeros = apenas_digitos(telefone or "")
    if not somente_numeros.startswith("55") and somente_numeros:
        somente_numeros = "55" + somente_numeros
    return f"https://wa.me/{somente_numeros}?text={quote(mensagem)}"


def link_email(email: str, assunto: str, corpo: str) -> str:
    return f"mailto:{email}?subject={quote(assunto)}&body={quote(corpo)}"


def componente_copiar(texto: str):
    escaped = (
        texto.replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("${", "\\${")
    )

    st.components.v1.html(
        f"""
        <div style="margin-top:6px;">
            <button
                onclick="navigator.clipboard.writeText(`{escaped}`); this.innerText='Mensagem copiada';"
                style="
                    background:#0d5db8;
                    color:white;
                    border:none;
                    padding:10px 14px;
                    border-radius:10px;
                    cursor:pointer;
                    font-weight:600;
                "
            >
                Copiar mensagem
            </button>
        </div>
        """,
        height=55,
    )

# =========================================
# AÇÕES DO PAINEL
# =========================================

def gerar_aditivo_renovacao_por_painel(pasta: Path, alerta_dias: int) -> tuple[bool, str]:
    dados_salvos = carregar_dados_condominio(pasta)
    if not dados_salvos:
        return False, "Esta pasta ainda não possui dados_condominio.json."

    data_fim_atual = dados_salvos.get("data_fim", "")
    novo_inicio, novo_fim = calcular_renovacao_anual(data_fim_atual)
    if not novo_inicio or not novo_fim:
        return False, "Não foi possível calcular a renovação porque a data final salva está inválida."

    dados_atualizados = dict(dados_salvos)
    dados_atualizados["data_inicio"] = formatar_data_br(novo_inicio)
    dados_atualizados["data_fim"] = formatar_data_br(novo_fim)
    dados_atualizados["data_assinatura"] = hoje_br()
    dados_atualizados["salvo_em"] = agora_br()

    placeholders = {
        "{{DATA_ASSINATURA}}": dados_atualizados.get("data_assinatura", ""),
        "{{NOME_CONDOMINIO}}": dados_atualizados.get("nome_condominio", ""),
        "{{CNPJ_CONDOMINIO}}": dados_atualizados.get("cnpj_condominio", ""),
        "{{ENDERECO_CONDOMINIO}}": dados_atualizados.get("endereco_condominio", ""),
        "{{NOME_SINDICO}}": dados_atualizados.get("nome_sindico", ""),
        "{{CPF_SINDICO}}": dados_atualizados.get("cpf_sindico", ""),
        "{{VALOR_MENSAL}}": valor_para_template(dados_atualizados.get("valor_mensal", "")),
        "{{VALOR_ADITIVO}}": valor_para_template(dados_atualizados.get("valor_aditivo", "")),
        "{{DATA_INICIO}}": dados_atualizados.get("data_inicio", ""),
        "{{DATA_FIM}}": dados_atualizados.get("data_fim", ""),
    }

    nome_condominio = dados_atualizados.get("nome_condominio", pasta.name)
    nome_sindico = dados_atualizados.get("nome_sindico", "")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")
    aditivo_docx = pasta / f"{base_nome_aditivo}.docx"
    aditivo_pdf = pasta / f"{base_nome_aditivo}.pdf"

    gerar_documento(
        template_path=TEMPLATE_ADITIVO,
        output_docx=aditivo_docx,
        placeholders=placeholders,
        incluir_assinaturas=st.session_state.get("incluir_assinaturas", True),
        nome_sindico=nome_sindico,
    )

    ok_pdf, erro_pdf = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)
    salvar_dados_condominio(pasta, dados_atualizados)
    aplicar_dados_no_formulario(dados_atualizados)

    registrar_documento_manifest(
        pasta_condominio=pasta,
        nome_condominio=nome_condominio,
        tipo="Aditivo",
        arquivo_docx=aditivo_docx,
        arquivo_pdf=aditivo_pdf,
        pdf_gerado=ok_pdf,
        erro_pdf=erro_pdf,
        dados_utilizados=placeholders,
    )

    if ok_pdf:
        return True, f"Aditivo de renovação gerado para '{nome_condominio}'. Nova vigência: {dados_atualizados['data_inicio']} até {dados_atualizados['data_fim']}."
    return True, f"Aditivo DOCX de renovação gerado para '{nome_condominio}', mas o PDF falhou: {erro_pdf}"

# =========================================
# PROCESSAMENTO DE VALIDAÇÃO
# =========================================

def validar_para_geracao(dados_base: dict, email_cliente: str) -> list[str]:
    faltando = validar_campos_obrigatorios(dados_base)
    if faltando:
        return [f"Preencha o campo obrigatório: {item}" for item in faltando]

    erros_formato = validar_campos_formato(dados_base, email_cliente)
    return erros_formato

# =========================================
# SESSÃO
# =========================================

def inicializar_campos():
    defaults = {
        "nome_condominio": "",
        "cnpj_condominio": "",
        "endereco_condominio": "",
        "nome_sindico": "",
        "cpf_sindico": "",
        "valor_mensal": "",
        "valor_aditivo": "",
        "data_inicio": "",
        "data_fim": "",
        "data_assinatura": hoje_br(),
        "whatsapp_cliente": "",
        "email_cliente": "",
        "observacoes_internas": "",
        "filtro_historico": "",
        "ultima_pasta_gerada": None,
        "confirm_delete_folder": "",
        "confirm_delete_file": "",
        "origem_dados_carregados": "",
        "alerta_vencimento_dias": 30,
        "painel_acao_msg": "",
        "busca_rapida": "",
        "filtro_status_central": "Todos",
        "incluir_assinaturas": True,
        "rel_mes_referencia": datetime.now().strftime("%m"),
        "rel_ano_referencia": str(datetime.now().year),
        "rel_art_numero": "",
        "rel_art_inicio": "",
        "rel_art_fim": "",
        "rel_data_emissao": hoje_br(),
        "rel_status_agua": "CONFORME",
        "rel_diagnostico": "",
    }
    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def limpar_formulario():
    st.session_state.nome_condominio = ""
    st.session_state.cnpj_condominio = ""
    st.session_state.endereco_condominio = ""
    st.session_state.nome_sindico = ""
    st.session_state.cpf_sindico = ""
    st.session_state.valor_mensal = ""
    st.session_state.valor_aditivo = ""
    st.session_state.data_inicio = ""
    st.session_state.data_fim = ""
    st.session_state.data_assinatura = hoje_br()
    st.session_state.whatsapp_cliente = ""
    st.session_state.email_cliente = ""
    st.session_state.observacoes_internas = ""
    st.session_state.origem_dados_carregados = ""
    st.session_state.rel_mes_referencia = datetime.now().strftime("%m")
    st.session_state.rel_ano_referencia = str(datetime.now().year)
    st.session_state.rel_art_numero = ""
    st.session_state.rel_art_inicio = ""
    st.session_state.rel_art_fim = ""
    st.session_state.rel_data_emissao = hoje_br()
    st.session_state.rel_status_agua = "CONFORME"
    st.session_state.rel_diagnostico = ""


inicializar_campos()

# =========================================
# TOPO
# =========================================

logo = encontrar_logo()

col_top1, col_top2 = st.columns([1, 5])

with col_top1:
    if logo:
        st.image(str(logo), width=150)

with col_top2:
    st.markdown(
        f"""
        <div class="top-card">
            <div class="top-title">{APP_TITLE}</div>
            <div class="top-subtitle">
                Sistema profissional para geração automatizada de contrato, aditivo e relatório mensal de RT
            </div>
            <div>
                <span class="info-badge">{RESPONSAVEL_TECNICO}</span>
                <span class="info-badge">{CRQ}</span>
                <span class="info-badge">{QUALIFICACAO_RT}</span>
                <span class="info-badge">Windows + Word + DOCX/PDF</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================
# SIDEBAR – CONFIG + HISTÓRICO
# =========================================

with st.sidebar:
    st.header("Histórico recente")

    st.number_input(
        "Lembrete de vencimento (dias)",
        min_value=1,
        max_value=180,
        step=1,
        key="alerta_vencimento_dias",
    )

    st.checkbox(
        "Adicionar bloco extra de assinaturas ao final",
        key="incluir_assinaturas",
        help="Mantenha ativado se seus templates não possuem assinatura pronta.",
    )

    st.text_input(
        "Filtrar condomínio",
        key="filtro_historico",
        placeholder="Digite parte do nome...",
    )

    historico = listar_historico()
    filtro = st.session_state.filtro_historico.strip().lower()

    if filtro:
        historico = [h for h in historico if filtro in h["nome"].lower()]

    if not historico:
        st.caption("Nenhum histórico encontrado.")
    else:
        for item in historico:
            nome_cond = item["nome"]
            pasta = item["pasta"]
            arquivos = item["arquivos"]
            folder_key = chave_segura(str(pasta))
            status = status_vencimento(item["data_fim"], st.session_state.alerta_vencimento_dias)
            resumo_docs = item.get("resumo_docs", {})

            titulo = f"{nome_cond} ({item['total_arquivos']})"

            with st.expander(titulo, expanded=False):
                st.caption(str(pasta))
                st.markdown(
                    f"<span class='status-badge {status['css']}'>{status['rotulo']}</span>",
                    unsafe_allow_html=True,
                )
                st.caption(status["mensagem"])
                st.caption(
                    f"Docs registrados: {resumo_docs.get('total', 0)} | "
                    f"Contratos: {resumo_docs.get('contratos', 0)} | "
                    f"Aditivos: {resumo_docs.get('aditivos', 0)} | "
                    f"Relatórios: {resumo_docs.get('relatorios', 0)}"
                )

                col_h1, col_h2 = st.columns(2)
                with col_h1:
                    if st.button(
                        "Carregar dados",
                        key=f"carregar_dados_{folder_key}",
                        use_container_width=True,
                    ):
                        dados_salvos = carregar_dados_condominio(pasta)
                        if dados_salvos:
                            aplicar_dados_no_formulario(dados_salvos)
                            st.success("Dados carregados no formulário.")
                            st.rerun()
                        else:
                            st.warning("Nenhum cadastro salvo encontrado para este condomínio.")

                with col_h2:
                    if st.button(
                        "Abrir pasta",
                        key=f"abrir_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        abrir_pasta_windows(pasta)

                if st.session_state.confirm_delete_folder == folder_key:
                    if st.button(
                        "Confirmar pasta",
                        key=f"confirmar_del_pasta_{folder_key}",
                        type="primary",
                        use_container_width=True,
                    ):
                        try:
                            deletar_pasta_condominio(pasta)
                            st.session_state.confirm_delete_folder = ""
                            st.success(f"Pasta de '{nome_cond}' excluída.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao excluir pasta: {e}")
                else:
                    if st.button(
                        "Excluir pasta",
                        key=f"pedir_del_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        st.session_state.confirm_delete_folder = folder_key
                        st.rerun()

                if st.session_state.confirm_delete_folder == folder_key:
                    st.markdown(
                        "<div class='confirm-box'><strong>Confirma excluir toda a pasta deste condomínio?</strong></div>",
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        "Cancelar exclusão da pasta",
                        key=f"cancelar_del_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        st.session_state.confirm_delete_folder = ""
                        st.rerun()

                st.markdown("**Arquivos:**")
                if not arquivos:
                    st.caption("Sem arquivos.")
                else:
                    for arq in arquivos:
                        caminho = arq["path"]
                        nome = arq["name"]
                        tipo_doc = arq["tipo_doc"]
                        tipo_ext = arq["tipo_ext"]
                        modificado = arq["modificado_em"]

                        file_key = chave_segura(str(caminho))

                        st.markdown(
                            f"**{tipo_doc} • {tipo_ext}**\n\n"
                            f"<div class='history-meta'>{nome}</div>"
                            f"<div class='history-meta'>Atualizado em: {modificado}</div>",
                            unsafe_allow_html=True,
                        )

                        col_a1, col_a2, col_a3 = st.columns([1.15, 1.0, 0.85])

                        with col_a1:
                            if st.button(
                                "Abrir arquivo",
                                key=f"abrir_arq_{file_key}",
                                use_container_width=True,
                            ):
                                abrir_arquivo_windows(caminho)

                        with col_a2:
                            if st.button(
                                "Abrir pasta",
                                key=f"abrir_pasta_arq_{file_key}",
                                use_container_width=True,
                            ):
                                abrir_pasta_windows(caminho.parent)

                        with col_a3:
                            if st.session_state.confirm_delete_file == file_key:
                                if st.button(
                                    "OK",
                                    key=f"confirmar_del_arq_{file_key}",
                                    type="primary",
                                    use_container_width=True,
                                ):
                                    try:
                                        deletar_arquivo_individual(caminho)
                                        st.session_state.confirm_delete_file = ""
                                        st.success(f"Arquivo excluído: {nome}")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Erro ao excluir arquivo: {e}")
                            else:
                                if st.button(
                                    "🗑️",
                                    key=f"pedir_del_arq_{file_key}",
                                    help="Excluir somente este arquivo",
                                    use_container_width=True,
                                ):
                                    st.session_state.confirm_delete_file = file_key
                                    st.rerun()

                        if st.session_state.confirm_delete_file == file_key:
                            st.markdown(
                                "<div class='confirm-box'>Excluir este arquivo?</div>",
                                unsafe_allow_html=True,
                            )
                            if st.button(
                                "Cancelar",
                                key=f"cancelar_del_arq_{file_key}",
                                use_container_width=True,
                            ):
                                st.session_state.confirm_delete_file = ""
                                st.rerun()

                        st.markdown("---")

# =========================================
# MODO DE OPERAÇÃO
# =========================================

modo = st.radio(
    "Modo de operação",
    ["Modo Escritório", "Modo Campo"],
    horizontal=True,
)

if modo == "Modo Campo":
    st.info("Fluxo compacto habilitado para operação em campo no Windows / tablet Windows.")

if st.session_state.origem_dados_carregados:
    st.markdown(
        f"""
        <div class="quick-mode-box">
            <strong>Dados carregados:</strong> {st.session_state.origem_dados_carregados}<br>
            Agora você pode ajustar valor do aditivo, vigência ou renovar automaticamente.
        </div>
        """,
        unsafe_allow_html=True,
    )

if st.session_state.painel_acao_msg:
    st.success(st.session_state.painel_acao_msg)
    st.session_state.painel_acao_msg = ""

status_formulario = status_vencimento(
    st.session_state.get("data_fim", ""),
    st.session_state.alerta_vencimento_dias,
)

if st.session_state.get("data_fim", "").strip():
    if status_formulario["codigo"] == "vencido":
        st.markdown(
            f"<div class='alert-vencido'><strong>Alerta de vencimento:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )
    elif status_formulario["codigo"] == "vencendo":
        st.markdown(
            f"<div class='alert-vencendo'><strong>Atenção:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )
    elif status_formulario["codigo"] == "vigente":
        st.markdown(
            f"<div class='alert-vigente'><strong>Status:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )

# =========================================
# DADOS DE PAINEL
# =========================================

painel_vencimentos = listar_painel_vencimentos(st.session_state.alerta_vencimento_dias)
painel_filtrado = filtrar_itens_painel(
    painel_vencimentos,
    st.session_state.busca_rapida,
    st.session_state.filtro_status_central,
)

total_monitorado = len(painel_vencimentos)
total_vencidos = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vencido"])
total_vencendo = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vencendo"])
total_vigentes = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vigente"])
total_indefinidos = len([i for i in painel_vencimentos if i["status"]["codigo"] == "indefinido"])
total_com_json = len([i for i in painel_vencimentos if i["tem_json"]])

itens_vencidos = [i for i in painel_filtrado if i["status"]["codigo"] == "vencido"]
itens_vencendo = [i for i in painel_filtrado if i["status"]["codigo"] == "vencendo"]
itens_indefinidos = [i for i in painel_filtrado if i["status"]["codigo"] == "indefinido"]

# =========================================
# DASHBOARD EXECUTIVO
# =========================================

taxa_estrutura = (total_com_json / total_monitorado * 100) if total_monitorado else 0
criticos = [i for i in painel_vencimentos if i["status"]["codigo"] in ("vencido", "vencendo")][:5]

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Dashboard executivo")

d1, d2, d3, d4, d5 = st.columns(5)
with d1:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Total monitorado</div><div class='dash-value'>{total_monitorado}</div><div class='dash-sub'>Pastas acompanhadas pelo sistema</div></div>",
        unsafe_allow_html=True,
    )
with d2:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vigentes</div><div class='dash-value'>{total_vigentes}</div><div class='dash-sub'>Dentro da vigência regular</div></div>",
        unsafe_allow_html=True,
    )
with d3:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vencem em breve</div><div class='dash-value'>{total_vencendo}</div><div class='dash-sub'>Dentro da faixa de alerta</div></div>",
        unsafe_allow_html=True,
    )
with d4:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vencidos</div><div class='dash-value'>{total_vencidos}</div><div class='dash-sub'>Exigem ação imediata</div></div>",
        unsafe_allow_html=True,
    )
with d5:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Cadastros estruturados</div><div class='dash-value'>{taxa_estrutura:.0f}%</div><div class='dash-sub'>{total_com_json} de {total_monitorado} com JSON salvo</div></div>",
        unsafe_allow_html=True,
    )

cx1, cx2 = st.columns([1.15, 1])
with cx1:
    st.markdown("**Resumo crítico**")
    if not criticos:
        st.success("Nenhum condomínio em faixa crítica no momento.")
    else:
        for item in criticos:
            st.markdown(
                f"- **{item['nome_exibicao']}** — {item['status']['rotulo']} — {item['status']['mensagem']}"
            )

with cx2:
    st.markdown("**Leitura rápida**")
    st.write(f"Sem vigência válida: **{total_indefinidos}**")
    st.write(f"Com ação prioritária agora: **{total_vencidos + total_vencendo}**")
    st.write(f"Faixa de lembrete atual: **{st.session_state.alerta_vencimento_dias} dias**")

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# SAÚDE DO SISTEMA
# =========================================

diag = diagnostico_sistema()

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Saúde do sistema")

s1, s2, s3, s4, s5, s6 = st.columns(6)
with s1:
    st.markdown(
        f"Template do contrato<br><span class='{'health-ok' if diag['template_contrato_ok'] else 'health-no'}'>{'OK' if diag['template_contrato_ok'] else 'Ausente'}</span>",
        unsafe_allow_html=True,
    )
with s2:
    st.markdown(
        f"Template do aditivo<br><span class='{'health-ok' if diag['template_aditivo_ok'] else 'health-no'}'>{'OK' if diag['template_aditivo_ok'] else 'Ausente'}</span>",
        unsafe_allow_html=True,
    )
with s3:
    st.markdown(
        f"Template do relatório<br><span class='{'health-ok' if diag['template_relatorio_ok'] else 'health-no'}'>{'OK' if diag['template_relatorio_ok'] else 'Ausente'}</span>",
        unsafe_allow_html=True,
    )
with s4:
    st.markdown(
        f"Pasta de documentos<br><span class='{'health-ok' if diag['generated_ok'] else 'health-no'}'>{'OK' if diag['generated_ok'] else 'Ausente'}</span>",
        unsafe_allow_html=True,
    )
with s5:
    st.markdown(
        f"Logo institucional<br><span class='{'health-ok' if diag['logo_ok'] else 'health-no'}'>{'OK' if diag['logo_ok'] else 'Não localizada'}</span>",
        unsafe_allow_html=True,
    )
with s5:
    st.markdown(
        f"Logo institucional<br><span class='{'health-ok' if diag['logo_ok'] else 'health-no'}'>{'OK' if diag['logo_ok'] else 'Não localizada'}</span>",
        unsafe_allow_html=True,
    )
with s6:
    st.markdown(
        f"Ambiente Windows<br><span class='{'health-ok' if diag['windows_ok'] else 'health-no'}'>{'OK' if diag['windows_ok'] else 'Fora do padrão'}</span>",
        unsafe_allow_html=True,
    )

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# BUSCA RÁPIDA
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Busca rápida profissional")

b1, b2, b3 = st.columns([2.2, 1.2, 1])
with b1:
    st.text_input(
        "Buscar por nome, CNPJ, síndico ou status",
        key="busca_rapida",
        placeholder="Ex.: terra nova, 12.345.678/0001-90, Marcelo, vencido...",
    )
with b2:
    st.selectbox(
        "Filtrar por status",
        options=["Todos", "Vigente", "Vence em breve", "Vencido", "Sem vigência válida"],
        key="filtro_status_central",
    )
with b3:
    st.metric("Resultado da busca", len(painel_filtrado))

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# PAINEL DE VENCIMENTOS
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Painel de vencimentos")

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Total monitorado", len(painel_filtrado))
with m2:
    st.metric("Vencidos", len([i for i in painel_filtrado if i["status"]["codigo"] == "vencido"]))
with m3:
    st.metric("Vencem em breve", len([i for i in painel_filtrado if i["status"]["codigo"] == "vencendo"]))
with m4:
    st.metric("Vigentes", len([i for i in painel_filtrado if i["status"]["codigo"] == "vigente"]))
with m5:
    st.metric("Sem vigência", len([i for i in painel_filtrado if i["status"]["codigo"] == "indefinido"]))

st.caption(
    "Painel operacional separado do histórico comum, com filtro central, exportação de cadastro e visualização dos últimos documentos."
)

aba1, aba2, aba3 = st.tabs(
    ["Condomínios vencidos", "Condomínios que vencem em breve", "Sem vigência válida"]
)

def render_exportacao_e_docs(item: dict, item_key: str):
    dados = item["dados"]
    ultimo_contrato = item["ultimo_contrato"]
    ultimo_aditivo = item["ultimo_aditivo"]
    ultimo_relatorio = item.get("ultimo_relatorio")
    resumo_docs = item.get("resumo_docs", {})
    ultimo_registro = resumo_docs.get("ultimo_registro") or {}

    st.markdown("<div class='docs-note'><strong>Documentos e controle</strong></div>", unsafe_allow_html=True)

    chips = []
    chips.append(f"<span class='doc-chip'>Registros: {resumo_docs.get('total', 0)}</span>")
    chips.append(f"<span class='doc-chip'>Contratos: {resumo_docs.get('contratos', 0)}</span>")
    chips.append(f"<span class='doc-chip'>Aditivos: {resumo_docs.get('aditivos', 0)}</span>")
    chips.append(f"<span class='doc-chip'>Relatórios: {resumo_docs.get('relatorios', 0)}</span>")
    if ultimo_registro.get("registrado_em"):
        chips.append(f"<span class='doc-chip'>Último: {ultimo_registro.get('registrado_em')}</span>")
    st.markdown("".join(chips), unsafe_allow_html=True)

    if dados and dados.get("salvo_em"):
        st.caption(f"Última atualização cadastral: {dados.get('salvo_em')}")

    dc1, dc2, dc3, dc4, dc5 = st.columns(5)
    with dc1:
        if ultimo_contrato:
            if st.button("Abrir último contrato", key=f"abrir_contrato_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_contrato["path"])
        else:
            st.button("Sem contrato", key=f"sem_contrato_{item_key}", disabled=True, use_container_width=True)

    with dc2:
        if ultimo_aditivo:
            if st.button("Abrir último aditivo", key=f"abrir_aditivo_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_aditivo["path"])
        else:
            st.button("Sem aditivo", key=f"sem_aditivo_{item_key}", disabled=True, use_container_width=True)

    with dc3:
        if ultimo_relatorio:
            if st.button("Abrir último relatório", key=f"abrir_relatorio_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_relatorio["path"])
        else:
            st.button("Sem relatório", key=f"sem_relatorio_{item_key}", disabled=True, use_container_width=True)

    with dc4:
        if dados:
            json_bytes = json.dumps(dados, ensure_ascii=False, indent=2).encode("utf-8")
            st.download_button(
                "Exportar JSON backup",
                data=json_bytes,
                file_name=f"{item['slug']}_cadastro.json",
                mime="application/json",
                key=f"download_json_{item_key}",
                use_container_width=True,
            )
        else:
            st.button("Sem JSON", key=f"sem_json_{item_key}", disabled=True, use_container_width=True)

    with dc5:
        if dados:
            html = gerar_html_resumo_cadastro(item).encode("utf-8")
            st.download_button(
                "Exportar resumo HTML",
                data=html,
                file_name=f"{item['slug']}_resumo_cadastro.html",
                mime="text/html",
                key=f"download_html_{item_key}",
                use_container_width=True,
            )
        else:
            st.button("Sem resumo", key=f"sem_resumo_{item_key}", disabled=True, use_container_width=True)

with aba1:
    if not itens_vencidos:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio vencido no momento.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_vencidos:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            data_fim = item["data_fim"] or "Não informada"
            item_key = chave_segura(f"painel_vencido_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Data final:</strong> {data_fim}</div>
                    <div class="venc-meta"><strong>Situação:</strong> {texto_dias_restantes(status)}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("Editar cadastro", key=f"editar_vencido_{item_key}", use_container_width=True):
                    aplicar_dados_no_formulario(dados_salvos)
                    st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                    st.rerun()

            with c2:
                if st.button("Abrir pasta", key=f"abrir_vencido_{item_key}", use_container_width=True):
                    abrir_pasta_windows(pasta)

            with c3:
                if st.button("Renovar no formulário", key=f"renovar_vencido_{item_key}", use_container_width=True):
                    ok, msg = preparar_renovacao_no_formulario(dados_salvos)
                    if ok:
                        st.session_state.painel_acao_msg = f"{msg} Condomínio: {nome_exibicao}."
                        st.rerun()
                    else:
                        st.error(msg)

            with c4:
                if st.button("Gerar aditivo renovação", key=f"aditivo_vencido_{item_key}", use_container_width=True):
                    ok, msg = gerar_aditivo_renovacao_por_painel(pasta, st.session_state.alerta_vencimento_dias)
                    if ok:
                        st.session_state.painel_acao_msg = msg
                        st.rerun()
                    else:
                        st.error(msg)

            render_exportacao_e_docs(item, item_key)

with aba2:
    if not itens_vencendo:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio dentro da faixa de alerta no momento.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_vencendo:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            data_fim = item["data_fim"] or "Não informada"
            item_key = chave_segura(f"painel_vencendo_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Data final:</strong> {data_fim}</div>
                    <div class="venc-meta"><strong>Situação:</strong> {texto_dias_restantes(status)}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("Editar cadastro", key=f"editar_vencendo_{item_key}", use_container_width=True):
                    aplicar_dados_no_formulario(dados_salvos)
                    st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                    st.rerun()

            with c2:
                if st.button("Abrir pasta", key=f"abrir_vencendo_{item_key}", use_container_width=True):
                    abrir_pasta_windows(pasta)

            with c3:
                if st.button("Renovar no formulário", key=f"renovar_vencendo_{item_key}", use_container_width=True):
                    ok, msg = preparar_renovacao_no_formulario(dados_salvos)
                    if ok:
                        st.session_state.painel_acao_msg = f"{msg} Condomínio: {nome_exibicao}."
                        st.rerun()
                    else:
                        st.error(msg)

            with c4:
                if st.button("Gerar aditivo renovação", key=f"aditivo_vencendo_{item_key}", use_container_width=True):
                    ok, msg = gerar_aditivo_renovacao_por_painel(pasta, st.session_state.alerta_vencimento_dias)
                    if ok:
                        st.session_state.painel_acao_msg = msg
                        st.rerun()
                    else:
                        st.error(msg)

            render_exportacao_e_docs(item, item_key)

with aba3:
    if not itens_indefinidos:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio sem vigência válida encontrado.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_indefinidos:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            origem = item["origem"]
            total_arquivos = len(item["arquivos"])
            item_key = chave_segura(f"painel_indefinido_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Status:</strong> {status['rotulo']}</div>
                    <div class="venc-meta"><strong>Arquivos encontrados:</strong> {total_arquivos}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if origem == "legado_sem_json":
                st.markdown(
                    "<div class='legacy-note'>Histórico antigo sem <strong>dados_condominio.json</strong>. "
                    "Agora você pode criar um cadastro inicial diretamente a partir desta pasta.</div>",
                    unsafe_allow_html=True,
                )

                c1, c2, c3 = st.columns(3)
                with c1:
                    if st.button("Criar cadastro desta pasta", key=f"legado_{item_key}", use_container_width=True):
                        preparar_cadastro_legado(pasta.name)
                        st.session_state.painel_acao_msg = f"Cadastro inicial preparado a partir da pasta '{pasta.name}'."
                        st.rerun()

                with c2:
                    if st.button("Abrir pasta", key=f"abrir_legado_{item_key}", use_container_width=True):
                        abrir_pasta_windows(pasta)

                with c3:
                    st.button("Gerar aditivo renovação", key=f"aditivo_legado_{item_key}", disabled=True, use_container_width=True)
            else:
                st.markdown(
                    "<div class='legacy-note'>Existe cadastro salvo, porém sem data final válida. "
                    "Carregue os dados no formulário e corrija a vigência.</div>",
                    unsafe_allow_html=True,
                )

                c1, c2, c3 = st.columns(3)
                with c1:
                    if st.button("Editar cadastro", key=f"editar_indefinido_{item_key}", use_container_width=True):
                        aplicar_dados_no_formulario(dados_salvos)
                        st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                        st.rerun()
                with c2:
                    if st.button("Abrir pasta", key=f"abrir_indefinido_{item_key}", use_container_width=True):
                        abrir_pasta_windows(pasta)
                with c3:
                    st.button("Gerar aditivo renovação", key=f"aditivo_indefinido_{item_key}", disabled=True, use_container_width=True)

            render_exportacao_e_docs(item, item_key)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# FORMULÁRIO
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Dados do contrato e aditivo")

col1, col2 = st.columns(2)

with col1:
    st.text_input("Nome do condomínio", key="nome_condominio")
    st.text_input(
        "CNPJ do condomínio",
        key="cnpj_condominio",
        on_change=on_change_cnpj,
        placeholder="00.000.000/0000-00",
    )
    st.text_area("Endereço do condomínio", key="endereco_condominio", height=100)
    st.text_input("Nome do síndico / representante", key="nome_sindico")
    st.text_input(
        "CPF do síndico / representante",
        key="cpf_sindico",
        on_change=on_change_cpf,
        placeholder="000.000.000-00",
    )

with col2:
    st.text_input(
        "Valor mensal",
        key="valor_mensal",
        on_change=on_change_valor_mensal,
        placeholder="R$ 1.621,00",
    )
    st.text_input(
        "Valor com desconto/aditivo",
        key="valor_aditivo",
        on_change=on_change_valor_aditivo,
        placeholder="R$ 810,50",
    )
    st.text_input(
        "Data de início",
        key="data_inicio",
        placeholder="01/04/2026",
        on_change=on_change_data_inicio,
    )
    st.text_input(
        "Data de fim",
        key="data_fim",
        placeholder="31/03/2027",
        on_change=on_change_data_fim,
    )
    st.text_input(
        "Data de assinatura",
        key="data_assinatura",
        placeholder="dd/mm/aaaa",
        on_change=on_change_data_assinatura,
    )

st.markdown("---")

if modo == "Modo Campo":
    col_campo1, col_campo2 = st.columns(2)
    with col_campo1:
        st.text_input(
            "WhatsApp do cliente (opcional)",
            key="whatsapp_cliente",
            on_change=on_change_whatsapp,
            placeholder="(34) 99999-9999",
        )
    with col_campo2:
        st.text_input("E-mail do cliente (opcional)", key="email_cliente")
else:
    col_cont1, col_cont2 = st.columns(2)
    with col_cont1:
        st.text_input(
            "WhatsApp do cliente",
            key="whatsapp_cliente",
            on_change=on_change_whatsapp,
            placeholder="(34) 99999-9999",
        )
    with col_cont2:
        st.text_input("E-mail do cliente", key="email_cliente")

st.text_area(
    "Observações internas (não vai para contrato/aditivo)",
    key="observacoes_internas",
    height=110,
    placeholder="Ex.: condição comercial específica, histórico jurídico, observação de cobrança, particularidades do condomínio...",
)

st.markdown('</div>', unsafe_allow_html=True)

# =========================================
# AÇÕES DE CADASTRO / RENOVAÇÃO
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Cadastro e renovação")

r1, r2, r3 = st.columns([1.15, 1.15, 2])

with r1:
    if st.button("Salvar cadastro / atualizar JSON", use_container_width=True):
        ok, msg = salvar_cadastro_sem_gerar_documentos()
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

with r2:
    if st.button("Renovação anual rápida", use_container_width=True):
        novo_inicio, novo_fim = calcular_renovacao_anual(st.session_state.data_fim)
        if not novo_inicio or not novo_fim:
            st.error("Não foi possível renovar. Verifique se a data final atual está válida.")
        else:
            st.session_state.data_inicio = formatar_data_br(novo_inicio)
            st.session_state.data_fim = formatar_data_br(novo_fim)
            st.session_state.data_assinatura = hoje_br()
            st.success("Nova vigência anual preenchida automaticamente.")
            st.rerun()

with r3:
    st.caption(
        "Salvar cadastro atualiza apenas o JSON do condomínio. "
        "Renovação anual rápida ajusta datas no formulário sem gerar documento automaticamente."
    )

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# RELATÓRIO MENSAL DE RESPONSABILIDADE TÉCNICA
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Relatório mensal de responsabilidade técnica")
st.caption("Módulo integrado ao histórico do condomínio, com geração em DOCX/PDF e registro no manifest.json.")

ra1, ra2, ra3, ra4 = st.columns(4)
with ra1:
    st.text_input("Mês de referência", key="rel_mes_referencia", placeholder="04")
with ra2:
    st.text_input("Ano de referência", key="rel_ano_referencia", placeholder="2026")
with ra3:
    st.text_input("ART nº", key="rel_art_numero", placeholder="Ex.: 123456")
with ra4:
    st.text_input("Data de emissão", key="rel_data_emissao", placeholder="dd/mm/aaaa")

rb1, rb2, rb3, rb4 = st.columns(4)
with rb1:
    st.text_input("Vigência da ART - início", key="rel_art_inicio", placeholder="01/04/2026")
with rb2:
    st.text_input("Vigência da ART - fim", key="rel_art_fim", placeholder="31/03/2027")
with rb3:
    st.selectbox("Status geral da água", ["CONFORME", "NÃO CONFORME", "EM CORREÇÃO"], key="rel_status_agua")
with rb4:
    st.text_input("Representante (usa cadastro)", value=st.session_state.get("nome_sindico", ""), disabled=True)

st.text_area(
    "Diagnóstico / parecer técnico",
    key="rel_diagnostico",
    height=120,
    placeholder="Descreva o parecer técnico mensal, condição operacional, não conformidades e medidas orientadas.",
)

with st.expander("Análises físico-químicas (10 linhas)", expanded=False):
    cab = st.columns([1.1, 0.7, 0.9, 0.9, 0.9, 0.9, 0.9, 1.2])
    rot = ["Data", "pH", "Cl livre", "Cl total", "Alcalin.", "Dureza", "Cianúr.", "Operador"]
    for c, r in zip(cab, rot):
        c.markdown(f"**{r}**")
    for i in range(10):
        cols = st.columns([1.1, 0.7, 0.9, 0.9, 0.9, 0.9, 0.9, 1.2])
        cols[0].text_input(f"Data {i+1}", key=f"rel_analise_data_{i}", label_visibility="collapsed", placeholder="dd/mm")
        cols[1].text_input(f"pH {i+1}", key=f"rel_analise_ph_{i}", label_visibility="collapsed")
        cols[2].text_input(f"Cl livre {i+1}", key=f"rel_analise_cl_{i}", label_visibility="collapsed")
        cols[3].text_input(f"Cl total {i+1}", key=f"rel_analise_ct_{i}", label_visibility="collapsed")
        cols[4].text_input(f"Alcal {i+1}", key=f"rel_analise_alc_{i}", label_visibility="collapsed")
        cols[5].text_input(f"Dureza {i+1}", key=f"rel_analise_dc_{i}", label_visibility="collapsed")
        cols[6].text_input(f"CYA {i+1}", key=f"rel_analise_cya_{i}", label_visibility="collapsed")
        cols[7].text_input(f"Operador {i+1}", key=f"rel_analise_operador_{i}", label_visibility="collapsed")

with st.expander("Dosagens químicas (7 linhas)", expanded=False):
    cab = st.columns([1.4, 1.3, 0.8, 0.7, 1.6])
    rot = ["Produto químico", "Fabricante/Lote", "Quantidade", "Unidade", "Finalidade técnica"]
    for c, r in zip(cab, rot):
        c.markdown(f"**{r}**")
    for i in range(7):
        cols = st.columns([1.4, 1.3, 0.8, 0.7, 1.6])
        cols[0].text_input(f"Produto {i+1}", key=f"rel_dos_produto_{i}", label_visibility="collapsed")
        cols[1].text_input(f"Lote {i+1}", key=f"rel_dos_lote_{i}", label_visibility="collapsed")
        cols[2].text_input(f"Qtd {i+1}", key=f"rel_dos_qtd_{i}", label_visibility="collapsed")
        cols[3].text_input(f"Un {i+1}", key=f"rel_dos_un_{i}", label_visibility="collapsed")
        cols[4].text_input(f"Finalidade {i+1}", key=f"rel_dos_finalidade_{i}", label_visibility="collapsed")

with st.expander("Recomendações técnicas (5 linhas)", expanded=False):
    cab = st.columns([2.2, 1.0, 1.1])
    rot = ["Recomendação", "Prazo", "Responsável"]
    for c, r in zip(cab, rot):
        c.markdown(f"**{r}**")
    for i in range(5):
        cols = st.columns([2.2, 1.0, 1.1])
        cols[0].text_input(f"Recomendação {i+1}", key=f"rel_rec_texto_{i}", label_visibility="collapsed")
        cols[1].text_input(f"Prazo {i+1}", key=f"rel_rec_prazo_{i}", label_visibility="collapsed")
        cols[2].text_input(f"Responsável {i+1}", key=f"rel_rec_resp_{i}", label_visibility="collapsed")

rc1, rc2, rc3 = st.columns(3)
with rc1:
    st.text_area("NBR 11238", key="rel_nbr_11238", height=90, placeholder="Informar condição de conformidade, observações e medidas necessárias.")
with rc2:
    st.text_area("NR-26 / GHS", key="rel_nr_26", height=90, placeholder="Sinalização, rotulagem, armazenamento e comunicação de risco.")
with rc3:
    st.text_area("NR-06 / EPI", key="rel_nr_06", height=90, placeholder="Condição de EPI, uso, exigências e orientações emitidas.")

relatorio_btn = st.button("Gerar relatório mensal DOCX + PDF", key="gerar_relatorio_mensal_btn", use_container_width=True)

st.markdown('</div>', unsafe_allow_html=True)

# =========================================
# AÇÕES PRINCIPAIS
# =========================================

dados = {
    "DATA_ASSINATURA": st.session_state.data_assinatura.strip(),
    "NOME_CONDOMINIO": st.session_state.nome_condominio.strip(),
    "CNPJ_CONDOMINIO": st.session_state.cnpj_condominio.strip(),
    "ENDERECO_CONDOMINIO": st.session_state.endereco_condominio.strip(),
    "NOME_SINDICO": st.session_state.nome_sindico.strip(),
    "CPF_SINDICO": st.session_state.cpf_sindico.strip(),
    "VALOR_MENSAL": valor_para_template(st.session_state.valor_mensal.strip()),
    "VALOR_ADITIVO": valor_para_template(st.session_state.valor_aditivo.strip()),
    "DATA_INICIO": st.session_state.data_inicio.strip(),
    "DATA_FIM": st.session_state.data_fim.strip(),
}

placeholders = {
    "{{DATA_ASSINATURA}}": dados["DATA_ASSINATURA"],
    "{{NOME_CONDOMINIO}}": dados["NOME_CONDOMINIO"],
    "{{CNPJ_CONDOMINIO}}": dados["CNPJ_CONDOMINIO"],
    "{{ENDERECO_CONDOMINIO}}": dados["ENDERECO_CONDOMINIO"],
    "{{NOME_SINDICO}}": dados["NOME_SINDICO"],
    "{{CPF_SINDICO}}": dados["CPF_SINDICO"],
    "{{VALOR_MENSAL}}": dados["VALOR_MENSAL"],
    "{{VALOR_ADITIVO}}": dados["VALOR_ADITIVO"],
    "{{DATA_INICIO}}": dados["DATA_INICIO"],
    "{{DATA_FIM}}": dados["DATA_FIM"],
}

col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1.5, 1.5, 1, 1])

with col_btn1:
    gerar = st.button(
        "Gerar contrato + aditivo",
        type="primary",
        use_container_width=True,
    )

with col_btn2:
    gerar_aditivo_rapido = st.button(
        "Gerar somente aditivo rápido",
        use_container_width=True,
    )

with col_btn3:
    if st.button("Limpar formulário", use_container_width=True):
        limpar_formulario()
        st.rerun()

with col_btn4:
    if st.button("Abrir pasta gerada", use_container_width=True):
        abrir_pasta_windows(GENERATED_DIR)

# =========================================
# FUNÇÕES DE PROCESSAMENTO DE DOCUMENTOS
# =========================================

def exibir_bloco_envio(
    nome_condominio: str,
    pasta_condominio: Path,
    whatsapp_cliente: str,
    email_cliente: str,
    mensagem: str,
):
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Mensagem pronta para envio")
    st.text_area(
        "Texto",
        value=mensagem,
        height=220,
        label_visibility="collapsed",
    )

    componente_copiar(mensagem)

    col_env1, col_env2, col_env3 = st.columns(3)

    with col_env1:
        if whatsapp_cliente.strip():
            url_wa = link_whatsapp(whatsapp_cliente, mensagem)
            st.link_button("Abrir WhatsApp", url_wa, use_container_width=True)
        else:
            st.button("Abrir WhatsApp", disabled=True, use_container_width=True)

    with col_env2:
        if email_cliente.strip():
            assunto = f"Documentação contratual – {nome_condominio}"
            url_mail = link_email(email_cliente, assunto, mensagem)
            st.link_button("Abrir e-mail", url_mail, use_container_width=True)
        else:
            st.button("Abrir e-mail", disabled=True, use_container_width=True)

    with col_env3:
        if st.button("Abrir pasta deste condomínio", use_container_width=True):
            abrir_pasta_windows(pasta_condominio)

    st.markdown("</div>", unsafe_allow_html=True)


def gerar_contrato_e_aditivo():
    email_cliente = st.session_state.email_cliente.strip()
    erros = validar_para_geracao(dados, email_cliente)

    if erros:
        st.error("Corrija os campos antes de gerar os documentos:")
        for erro in erros:
            st.write(f"- {erro}")
        return

    try:
        nome_condominio = st.session_state.nome_condominio.strip()
        nome_sindico = st.session_state.nome_sindico.strip()
        whatsapp_cliente = st.session_state.whatsapp_cliente.strip()

        nome_pasta = slugify_nome(nome_condominio)
        pasta_condominio = GENERATED_DIR / nome_pasta
        pasta_condominio.mkdir(parents=True, exist_ok=True)

        salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_nome_contrato = limpar_nome_arquivo(f"Contrato_RT_{nome_condominio}_{timestamp}")
        base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")

        contrato_docx = pasta_condominio / f"{base_nome_contrato}.docx"
        contrato_pdf = pasta_condominio / f"{base_nome_contrato}.pdf"

        aditivo_docx = pasta_condominio / f"{base_nome_aditivo}.docx"
        aditivo_pdf = pasta_condominio / f"{base_nome_aditivo}.pdf"

        with st.spinner("Gerando contrato..."):
            gerar_documento(
                template_path=TEMPLATE_CONTRATO,
                output_docx=contrato_docx,
                placeholders=placeholders,
                incluir_assinaturas=st.session_state.get("incluir_assinaturas", True),
                nome_sindico=nome_sindico,
            )

        with st.spinner("Gerando aditivo..."):
            gerar_documento(
                template_path=TEMPLATE_ADITIVO,
                output_docx=aditivo_docx,
                placeholders=placeholders,
                incluir_assinaturas=st.session_state.get("incluir_assinaturas", True),
                nome_sindico=nome_sindico,
            )

        ok_contrato, erro_contrato = converter_docx_para_pdf(contrato_docx, contrato_pdf)
        ok_aditivo, erro_aditivo = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)

        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Contrato",
            arquivo_docx=contrato_docx,
            arquivo_pdf=contrato_pdf,
            pdf_gerado=ok_contrato,
            erro_pdf=erro_contrato,
            dados_utilizados=placeholders,
        )

        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Aditivo",
            arquivo_docx=aditivo_docx,
            arquivo_pdf=aditivo_pdf,
            pdf_gerado=ok_aditivo,
            erro_pdf=erro_aditivo,
            dados_utilizados=placeholders,
        )

        st.session_state.ultima_pasta_gerada = str(pasta_condominio)

        st.success("Contrato e aditivo gerados com sucesso.")

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Arquivos gerados")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("**Contrato**")
            if contrato_docx.exists():
                with open(contrato_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do contrato",
                        data=f,
                        file_name=contrato_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            if ok_contrato and contrato_pdf.exists():
                with open(contrato_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do contrato",
                        data=f,
                        file_name=contrato_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do contrato não gerado. Erro: {erro_contrato}")

        with c2:
            st.markdown("**Aditivo**")
            if aditivo_docx.exists():
                with open(aditivo_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do aditivo",
                        data=f,
                        file_name=aditivo_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            if ok_aditivo and aditivo_pdf.exists():
                with open(aditivo_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do aditivo",
                        data=f,
                        file_name=aditivo_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do aditivo não gerado. Erro: {erro_aditivo}")

        st.markdown("</div>", unsafe_allow_html=True)

        mensagem = montar_mensagem_envio(
            nome_condominio=nome_condominio,
            nome_sindico=nome_sindico,
            caminho_contrato_pdf=contrato_pdf if contrato_pdf.exists() else None,
            caminho_aditivo_pdf=aditivo_pdf if aditivo_pdf.exists() else None,
        )

        exibir_bloco_envio(
            nome_condominio=nome_condominio,
            pasta_condominio=pasta_condominio,
            whatsapp_cliente=whatsapp_cliente,
            email_cliente=email_cliente,
            mensagem=mensagem,
        )

    except Exception as e:
        st.error(f"Erro na geração dos documentos: {e}")


def gerar_somente_aditivo_rapido():
    email_cliente = st.session_state.email_cliente.strip()
    erros = validar_para_geracao(dados, email_cliente)

    if erros:
        st.error("Corrija os campos antes de gerar o aditivo rápido:")
        for erro in erros:
            st.write(f"- {erro}")
        return

    try:
        nome_condominio = st.session_state.nome_condominio.strip()
        nome_sindico = st.session_state.nome_sindico.strip()
        whatsapp_cliente = st.session_state.whatsapp_cliente.strip()

        nome_pasta = slugify_nome(nome_condominio)
        pasta_condominio = GENERATED_DIR / nome_pasta
        pasta_condominio.mkdir(parents=True, exist_ok=True)

        salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")

        aditivo_docx = pasta_condominio / f"{base_nome_aditivo}.docx"
        aditivo_pdf = pasta_condominio / f"{base_nome_aditivo}.pdf"

        with st.spinner("Gerando aditivo rápido..."):
            gerar_documento(
                template_path=TEMPLATE_ADITIVO,
                output_docx=aditivo_docx,
                placeholders=placeholders,
                incluir_assinaturas=st.session_state.get("incluir_assinaturas", True),
                nome_sindico=nome_sindico,
            )

        ok_aditivo, erro_aditivo = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)

        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Aditivo",
            arquivo_docx=aditivo_docx,
            arquivo_pdf=aditivo_pdf,
            pdf_gerado=ok_aditivo,
            erro_pdf=erro_aditivo,
            dados_utilizados=placeholders,
        )

        st.session_state.ultima_pasta_gerada = str(pasta_condominio)

        st.success("Aditivo rápido gerado com sucesso.")

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Arquivo gerado")

        c1, c2 = st.columns(2)

        with c1:
            if aditivo_docx.exists():
                with open(aditivo_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do aditivo",
                        data=f,
                        file_name=aditivo_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

        with c2:
            if ok_aditivo and aditivo_pdf.exists():
                with open(aditivo_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do aditivo",
                        data=f,
                        file_name=aditivo_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do aditivo não gerado. Erro: {erro_aditivo}")

        st.markdown("</div>", unsafe_allow_html=True)

        mensagem = montar_mensagem_envio(
            nome_condominio=nome_condominio,
            nome_sindico=nome_sindico,
            caminho_contrato_pdf=None,
            caminho_aditivo_pdf=aditivo_pdf if aditivo_pdf.exists() else None,
        )

        exibir_bloco_envio(
            nome_condominio=nome_condominio,
            pasta_condominio=pasta_condominio,
            whatsapp_cliente=whatsapp_cliente,
            email_cliente=email_cliente,
            mensagem=mensagem,
        )

    except Exception as e:
        st.error(f"Erro na geração do aditivo rápido: {e}")

# =========================================
# PROCESSAMENTO
# =========================================

if gerar:
    gerar_contrato_e_aditivo()

if gerar_aditivo_rapido:
    gerar_somente_aditivo_rapido()

if relatorio_btn:
    ok_rel, msg_rel = gerar_relatorio_mensal()
    if ok_rel:
        st.success(msg_rel)
    else:
        st.error(msg_rel)

# =========================================
# RODAPÉ
# =========================================

st.markdown("---")
st.caption(
    f"{APP_TITLE} • {RESPONSAVEL_TECNICO} • {CRQ} • Ambiente prioritário: Windows"
)