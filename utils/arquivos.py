import re


def classificar_arquivo(nome_arquivo: str) -> tuple[str, str]:
    nome_lower = nome_arquivo.lower()

    if "contrato" in nome_lower:
        tipo_doc = "Contrato"
    elif "aditivo" in nome_lower:
        tipo_doc = "Aditivo"
    elif "relatorio" in nome_lower:
        tipo_doc = "Relatório"
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
