import re


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
    texto = re.sub(r'[<>:"/\\|?*]+', "", texto)
    texto = re.sub(r"\s+", "_", texto.strip())
    return texto[:150]


def apenas_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


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


def validar_email(email: str) -> bool:
    email = (email or "").strip()
    if not email:
        return True
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return re.match(padrao, email) is not None
