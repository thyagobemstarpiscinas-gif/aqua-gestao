"""Cópia persistente das pastas de condomínio (generated/<pasta>) no Google Drive.

No Streamlit Cloud o disco local é apagado a cada reinício. Os dados contratuais
(dados_condominio.json), o histórico de documentos (manifest.json) e os documentos
gerados (DOCX/PDF) de cada condomínio alimentam o painel de vencimentos, os contadores
e o histórico. Este módulo mantém uma cópia deles no Drive já usado pelo app:

    <pasta de fotos do app>/_Pastas_Condominios/<nome da pasta local>/<arquivo>

Regras:
- só pastas de condomínio (filhas diretas de generated/, nome sem "_" inicial) e só os
  arquivos do nível principal com nome sem "_" inicial (rascunhos, temporários e
  subpastas de fotos têm armazenamento próprio e ficam de fora);
- cada arquivo tem uma única cópia por nome: gravações seguintes atualizam a mesma
  (sem duplicar); a data de modificação original vai em appProperties.mtime;
- restauração: baixa só o que falta (ou está mais antigo) no disco local e devolve a
  data original ao arquivo; nunca apaga nem sobrescreve arquivo local mais novo;
- exclusão feita pelo app envia a cópia para a lixeira do Drive (reversível).
"""
from __future__ import annotations

import hashlib
import io
import mimetypes
import os
from pathlib import Path

NOME_BASE = "_Pastas_Condominios"
NOME_MANIFESTO = "manifest.json"
MIME_PASTA = "application/vnd.google-apps.folder"
TOLERANCIA_S = 2  # diferença de mtime abaixo disso não é alteração


def pasta_persistivel(pasta: Path, raiz: Path) -> bool:
    try:
        return pasta.parent.resolve() == raiz.resolve() and not pasta.name.startswith(("_", "."))
    except OSError:
        return False


def arquivos_persistiveis(pasta: Path) -> list[Path]:
    if not pasta.is_dir():
        return []
    return sorted(p for p in pasta.iterdir() if p.is_file() and not p.name.startswith(("_", ".")))


def _esc(texto: str) -> str:
    return str(texto).replace("\\", "\\\\").replace("'", "\\'")


def _md5(caminho: Path) -> str:
    return hashlib.md5(caminho.read_bytes()).hexdigest()


def _md5_remoto(item: dict) -> str:
    return str((item.get("appProperties") or {}).get("md5", ""))


def _mtime_remoto(item: dict) -> float:
    try:
        return float((item.get("appProperties") or {}).get("mtime", 0))
    except (TypeError, ValueError):
        return 0.0


class DrivePastas:
    """Operações mínimas no Drive (service = googleapiclient drive v3)."""

    def __init__(self, service, raiz_id: str, nome_base: str = NOME_BASE):
        self.s, self.raiz_id, self.nome_base = service, raiz_id, nome_base
        self._ids: dict[tuple[str, str], str] = {}

    def _listar(self, q: str) -> list[dict]:
        itens, token = [], None
        while True:
            r = self.s.files().list(
                q=q, pageSize=1000, pageToken=token,
                fields="nextPageToken, files(id, name, mimeType, appProperties, modifiedTime)").execute()
            itens.extend(r.get("files", []))
            token = r.get("nextPageToken")
            if not token:
                return itens

    def _pasta(self, nome: str, pai_id: str, criar: bool) -> str | None:
        chave = (pai_id, nome)
        if chave in self._ids:
            return self._ids[chave]
        achadas = self._listar(f"name='{_esc(nome)}' and '{_esc(pai_id)}' in parents "
                               f"and mimeType='{MIME_PASTA}' and trashed=false")
        if achadas:
            self._ids[chave] = sorted(achadas, key=lambda a: a["id"])[0]["id"]  # escolha estável
        elif criar:
            self._ids[chave] = self.s.files().create(
                body={"name": nome, "mimeType": MIME_PASTA, "parents": [pai_id]}, fields="id").execute()["id"]
        else:
            return None
        return self._ids[chave]

    def base(self, criar: bool = True) -> str | None:
        return self._pasta(self.nome_base, self.raiz_id, criar)

    def pasta_condominio(self, nome: str, criar: bool = True) -> str | None:
        base = self.base(criar)
        return self._pasta(nome, base, criar) if base else None

    def subpastas(self) -> list[dict]:
        base = self.base(criar=False)
        if not base:
            return []
        return self._listar(f"'{_esc(base)}' in parents and mimeType='{MIME_PASTA}' and trashed=false")

    def arquivos(self, pasta_id: str) -> dict[str, dict]:
        """nome -> item mais recente (se houver nomes repetidos, vale o mais novo)."""
        por_nome: dict[str, dict] = {}
        for item in self._listar(f"'{_esc(pasta_id)}' in parents and mimeType!='{MIME_PASTA}' and trashed=false"):
            atual = por_nome.get(item["name"])
            if atual is None or _mtime_remoto(item) > _mtime_remoto(atual):
                por_nome[item["name"]] = item
        return por_nome

    def enviar(self, caminho: Path, pasta_id: str, existente: dict | None) -> None:
        from googleapiclient.http import MediaIoBaseUpload
        mime = mimetypes.guess_type(caminho.name)[0] or "application/octet-stream"
        media = MediaIoBaseUpload(io.BytesIO(caminho.read_bytes()), mimetype=mime, resumable=False)
        props = {"appProperties": {"mtime": str(int(caminho.stat().st_mtime)), "md5": _md5(caminho)}}
        if existente:
            self.s.files().update(fileId=existente["id"], body=props, media_body=media, fields="id").execute()
        else:
            self.s.files().create(body={"name": caminho.name, "parents": [pasta_id], **props},
                                  media_body=media, fields="id").execute()

    def baixar(self, file_id: str) -> bytes:
        return self.s.files().get_media(fileId=file_id).execute()

    def lixeira(self, file_id: str) -> None:
        self.s.files().update(fileId=file_id, body={"trashed": True}, fields="id").execute()


def _mesclar_manifesto(destino: Path, remoto_bytes: bytes) -> None:
    """Manifesto local mais novo + registros que só existem no Drive (nenhum registro se perde).

    Caso de uso: o Drive falhou na leitura logo após um reinício e um documento foi
    registrado num manifesto local vazio; o histórico anterior continua no Drive."""
    import json
    try:
        local = json.loads(destino.read_text(encoding="utf-8"))
        remoto = json.loads(remoto_bytes.decode("utf-8"))
    except (ValueError, OSError, UnicodeDecodeError):
        return
    docs_local = list(local.get("documentos") or [])
    chave = lambda d: json.dumps(d, sort_keys=True, ensure_ascii=False)  # noqa: E731
    vistos = {chave(d) for d in docs_local}
    faltantes = [d for d in (remoto.get("documentos") or []) if chave(d) not in vistos]
    if faltantes:
        local["documentos"] = faltantes + docs_local  # anteriores primeiro, como na gravação original
        destino.write_text(json.dumps(local, ensure_ascii=False, indent=2), encoding="utf-8")


def sincronizar_pasta(drive: DrivePastas, pasta: Path) -> int:
    """Envia ao Drive os arquivos novos ou alterados da pasta. Retorna quantos enviou."""
    locais = arquivos_persistiveis(pasta)
    if not locais:
        return 0
    pasta_id = drive.pasta_condominio(pasta.name, criar=True)
    remotos = drive.arquivos(pasta_id)
    enviados = 0
    for arq in locais:
        rem = remotos.get(arq.name)
        # Conteúdo decide (gravações seguidas no mesmo segundo também são enviadas).
        if rem is None or _md5_remoto(rem) != _md5(arq):
            drive.enviar(arq, pasta_id, rem)
            enviados += 1
    return enviados


def restaurar(drive: DrivePastas, raiz: Path) -> dict:
    """Traz do Drive o que falta no disco e envia o que só existe (ou é mais novo) no disco."""
    raiz.mkdir(parents=True, exist_ok=True)
    resumo = {"pastas": 0, "baixados": 0, "enviados": 0}
    remotas = {p["name"]: p["id"] for p in drive.subpastas() if not p["name"].startswith(("_", "."))}
    for nome, pasta_id in remotas.items():
        local = raiz / nome
        if not pasta_persistivel(local, raiz):
            continue
        resumo["pastas"] += 1
        drive._ids[(drive.base(criar=False), nome)] = pasta_id
        mais_recente = None
        for nome_arq, item in drive.arquivos(pasta_id).items():
            if nome_arq.startswith(("_", ".")) or "/" in nome_arq or "\\" in nome_arq:
                continue
            destino = local / nome_arq
            mt = _mtime_remoto(item)
            if destino.exists():
                if _md5_remoto(item) and _md5_remoto(item) == _md5(destino):
                    continue  # mesmo conteúdo
                if destino.stat().st_mtime + TOLERANCIA_S >= mt:
                    if nome_arq == NOME_MANIFESTO:
                        _mesclar_manifesto(destino, drive.baixar(item["id"]))
                    continue  # local mais novo: preservado (enviado em seguida)
            local.mkdir(exist_ok=True)
            temporario = destino.with_name(f".{nome_arq}.restaurando")
            temporario.write_bytes(drive.baixar(item["id"]))
            os.replace(temporario, destino)
            if mt:
                os.utime(destino, (mt, mt))
            resumo["baixados"] += 1
        datas = [p.stat().st_mtime for p in arquivos_persistiveis(local)]
        if datas:
            mais_recente = max(datas)
            os.utime(local, (mais_recente, mais_recente))
    for pasta in sorted(p for p in raiz.iterdir() if p.is_dir()):
        if pasta_persistivel(pasta, raiz):
            resumo["enviados"] += sincronizar_pasta(drive, pasta)
    return resumo


def excluir_arquivo(drive: DrivePastas, pasta: Path, nome_arquivo: str) -> None:
    pasta_id = drive.pasta_condominio(pasta.name, criar=False)
    if pasta_id:
        item = drive.arquivos(pasta_id).get(nome_arquivo)
        if item:
            drive.lixeira(item["id"])


def excluir_pasta(drive: DrivePastas, pasta: Path) -> None:
    pasta_id = drive.pasta_condominio(pasta.name, criar=False)
    if pasta_id:
        drive.lixeira(pasta_id)
        drive._ids.pop((drive.base(criar=False), pasta.name), None)
