"""Contexto RT serializável; nenhuma conexão externa ou dependência de Streamlit.

UploadFile não é persistível. Guardamos bytes, nome e digest dentro do rascunho
do próprio contexto, e só materializamos arquivos após validar sua identidade.
"""
import base64
import hashlib
import json
from pathlib import Path

ART_PADRAO = "Em processo de formalização"
TRANSITORIOS = {"rel_fotos_upload", "rel_pdf_visita_importador", "rel_regerar_pdf_visita_idx",
                "rel_art_modo_rt2025", "rel_art_numero_rt2025", "rel_status_agua_select"}
IDENTIFICACAO = ("rel_nome_condominio", "rel_mes_referencia", "rel_ano_referencia")
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
