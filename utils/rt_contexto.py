"""Contexto RT serializável; nenhuma conexão externa ou dependência de Streamlit.

UploadFile não é persistível. Guardamos bytes, nome e digest dentro do rascunho
do próprio contexto, e só materializamos arquivos após validar sua identidade.
"""
import base64
import hashlib
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ART_PADRAO = "Em processo de formalização"
TRANSITORIOS = {"rel_fotos_upload", "rel_pdf_visita_importador", "rel_regerar_pdf_visita_idx",
                "rel_art_modo_rt2025", "rel_art_numero_rt2025", "rel_status_agua_select"}
IDENTIFICACAO = ("rel_nome_condominio", "rel_mes_referencia", "rel_ano_referencia")


def normalizar_rascunho(dados, empresa, minimo):
    if not isinstance(dados, dict) or not isinstance(dados.get("campos"), dict) or not dados["campos"]:
        raise ValueError("Rascunho sem campos válidos.")
    valores = dict(dados["campos"])
    origem = valores.get("rel_origem_dados") or ""
    if not isinstance(origem, str) or dados.get("empresa_ativa", "aqua_gestao") != empresa or (origem and not origem.startswith(empresa + "|")):
        raise ValueError("Rascunho de outra empresa: restauração bloqueada.")
    for k, v in valores.items():
        if not isinstance(k, str) or (k.startswith("rel_") and k not in TRANSITORIOS and not isinstance(v, (str, int, float, bool, type(None)))):
            raise ValueError("Campo de rascunho inválido.")
    def contador(v):
        if isinstance(v, bool) or not isinstance(v, (int, str)) or not str(v).strip().isdigit():
            raise ValueError("Total de análises inválido.")
        return max(int(v), minimo)
    # Valida ambas as fontes, mesmo quando o metadado de topo tem precedência.
    if "rel_analises_total" in valores:
        contador(valores["rel_analises_total"])
    total = contador(dados.get("rel_analises_total", valores.get("rel_analises_total", minimo)))
    valores["rel_analises_total"] = total
    ler_fotos(valores, origem)
    return valores, total


def identidade_ultimos(chave):
    partes = str(chave).split("|")
    if len(partes) != 3 or partes[0] not in ("aqua_gestao", "bem_star") or not partes[1].startswith("id:") or not partes[1][3:].strip():
        raise ValueError("Últimos valores exigem empresa e ID estável validados.")
    periodo = partes[2].split("/")
    if len(periodo) != 2 or not all(p.isdigit() for p in periodo) or not 1 <= int(periodo[0]) <= 12 or len(periodo[1]) != 4 or not 2000 <= int(periodo[1]) <= 2100:
        raise ValueError("Período de origem inválido.")
    return partes[0], partes[1][3:], f"{int(periodo[0]):02d}/{periodo[1]}"


def caminho_ultimos(raiz, chave):
    empresa, cliente, _ = identidade_ultimos(chave)
    digest = hashlib.sha256(json.dumps([empresa, cliente], ensure_ascii=False).encode()).hexdigest()
    return Path(raiz) / "_rt_ultimos" / digest / "ultimo_final.json"


def validar_ultimos(dados, chave):
    empresa, cliente, _ = identidade_ultimos(chave)
    if not isinstance(dados, dict) or dados.get("versao") != 1 or dados.get("empresa") != empresa or dados.get("id_cliente") != cliente:
        raise ValueError("Registro de últimos valores inválido ou de outra identidade.")
    identidade_ultimos(f"{empresa}|id:{cliente}|{dados.get('periodo', '')}")
    datetime.fromisoformat(dados["gerado_em"])
    for nome in ("parametros_ultimos", "dosagens_ultimas"):
        itens = dados.get(nome)
        if not isinstance(itens, list) or not all(isinstance(i, dict) and all(isinstance(k, str) and isinstance(v, str) for k, v in i.items()) for i in itens):
            raise ValueError("Valores recuperáveis inválidos.")
        limite = 40 if nome == "parametros_ultimos" else 7
        if not itens or len(itens) > limite:
            raise ValueError("Quantidade de valores recuperáveis inválida.")
        obrigatorio = "ph" if nome == "parametros_ultimos" else "produto"
        if any(obrigatorio not in item for item in itens):
            raise ValueError("Registro incompleto de últimos valores.")
    return dados


def montar_ultimos(chave, snapshot):
    empresa, cliente, periodo = identidade_ultimos(chave)
    dados = {"versao": 1, "empresa": empresa, "id_cliente": cliente, "periodo": periodo,
             "gerado_em": datetime.now(timezone.utc).isoformat(),
             "parametros_ultimos": [{k: v for k, v in item.items() if k != "data"} for item in snapshot["parametros_ultimos"]],
             "dosagens_ultimas": snapshot["dosagens_ultimas"]}
    validar_ultimos(dados, chave)
    return dados


def registrar_ultimos(raiz, chave, snapshot):
    dados = montar_ultimos(chave, snapshot)
    destino = caminho_ultimos(raiz, chave)
    destino.parent.mkdir(parents=True, exist_ok=True)
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=destino.parent, delete=False) as f:
            temporario = Path(f.name)
            json.dump(dados, f, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporario, destino)
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)


def carregar_ultimos(raiz, chave):
    return validar_ultimos(json.loads(caminho_ultimos(raiz, chave).read_text(encoding="utf-8")), chave)
# Campos de escolha (seletores e opções): vazio não é opção válida. Ao limpar, voltam ao
# padrão que o próprio widget usaria; senão o widget corrige o valor na execução seguinte
# e o salvamento automático confunde a correção com uma edição do usuário.
PADROES_ESCOLHA = {
    "rel_tipo_atendimento": "Contrato ativo",
    "rel_status_agua": "CONFORME",
    **{f"rel_epi_{e}_status": "Conforme" for e in ("luvas", "oculos", "respirador", "botas")},
    **{f"rel_nbr11238_{c}": "Sim" for c in ("profundidade", "retrolavagem", "skimmers", "circulacao", "chuveiro")},
}


def campos(ss):
    return {k: v for k, v in ss.items() if str(k).startswith("rel_")
            and k not in TRANSITORIOS and isinstance(v, (str, int, float, bool, type(None)))}


def limpar(ss):
    for k, v in campos(ss).items():
        ss[k] = "" if isinstance(v, str) else (False if isinstance(v, bool) else 0)
    ss.update(PADROES_ESCOLHA)
    ss["rel_art_status"] = ss["rel_art_status_widget"] = ART_PADRAO
    ss["rel_fotos_preservadas"] = ""
    for k in TRANSITORIOS - {"rel_fotos_upload", "rel_pdf_visita_importador"}:
        ss.pop(k, None)
    ss["_rt_fotos_nonce"] = ss.get("_rt_fotos_nonce", 0) + 1
    ss.pop("_rel_fotos_origem", None)
    ss.pop("_rt_upload_hashes", None)


def ler_fotos(ss, chave):
    raw = ss.get("rel_fotos_preservadas") or ""
    if not raw:
        return []
    data = json.loads(raw)
    if data.get("identidade") != chave:
        raise ValueError("Fotografias pertencem a outro condomínio ou período.")
    for f in data["fotos"]:
        conteudo = base64.b64decode(f["bytes"], validate=True)
        if hashlib.sha256(conteudo).hexdigest() != f["sha256"]:
            raise ValueError("Fotografia do rascunho corrompida.")
    return data["fotos"]


def sincronizar_fotos(ss, chave, uploads):
    fotos = {f["sha256"]: f for f in ler_fotos(ss, chave)}
    novos = {}
    for up in uploads or []:
        b = bytes(up.getvalue())
        digest = hashlib.sha256(b).hexdigest()
        novos[digest] = {"nome": Path(up.name).name, "sha256": digest,
                         "bytes": base64.b64encode(b).decode("ascii")}
    # Remoção no uploader vale para os anexos ativos, sem apagar fotos restauradas.
    for digest in set(ss.get("_rt_upload_hashes") or []) - set(novos):
        fotos.pop(digest, None)
    fotos.update(novos)
    ss["_rt_upload_hashes"] = list(novos)
    if fotos or ss.get("rel_fotos_preservadas"):
        ss["rel_fotos_preservadas"] = json.dumps({"identidade": chave, "fotos": list(fotos.values())})
    return list(fotos.values())


def materializar_fotos(ss, chave, raiz):
    pasta = Path(raiz) / "_rt_contextos" / hashlib.sha256(chave.encode()).hexdigest() / "fotos"
    resultado = []
    for f in ler_fotos(ss, chave):
        # O nome de origem é só metadado; a persistência não depende dele.
        pasta.mkdir(parents=True, exist_ok=True)
        ext = Path(f["nome"]).suffix.lower()
        if ext not in (".png", ".jpg", ".jpeg", ".webp"):
            raise ValueError("Formato de fotografia não suportado no relatório RT.")
        p = pasta / (f["sha256"] + ext)
        p.write_bytes(base64.b64decode(f["bytes"], validate=True))
        resultado.append(p)
    return resultado
