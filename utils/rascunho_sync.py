"""Pendência local durável. Sem Streamlit, credenciais ou chamadas de rede próprias."""
import hashlib
import json
import os
from pathlib import Path
import tempfile


class ConflitoRascunho(ValueError):
    pass


def empresa_dados(dados):
    empresa = dados.get("empresa_ativa", "aqua_gestao")
    origem = dados.get("campos", {}).get("rel_origem_dados") or ""
    if empresa not in ("aqua_gestao", "bem_star") or (origem and not str(origem).startswith(empresa + "|")):
        raise ValueError("Identidade de empresa inválida no rascunho")
    return empresa


def hash_conteudo(dados):
    from utils.rt_contexto import TRANSITORIOS
    campos = dados.get("campos")
    if not isinstance(campos, dict):
        raise ValueError("Conteúdo de rascunho ilegível")
    conteudo = {"empresa": empresa_dados(dados),
                "total": dados.get("rel_analises_total"),
                "campos": {k: v for k, v in campos.items() if k not in TRANSITORIOS}}
    return hashlib.sha256(json.dumps(conteudo, sort_keys=True, ensure_ascii=False,
                                    allow_nan=False).encode("utf-8")).hexdigest()


def escrever_atomico(destino, dados):
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=destino.parent,
                                         prefix=destino.stem + ".", suffix=".tmp", delete=False) as f:
            temporario = Path(f.name)
            json.dump(dados, f, ensure_ascii=False, allow_nan=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temporario, destino)
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)


class GestorRascunhos:
    def __init__(self, raiz, empresa, ss, arquivo, desserializar, gravar, eh_chave, ordem, serializar):
        if empresa not in ("aqua_gestao", "bem_star"):
            raise ValueError("Empresa inválida")
        self.raiz, self.empresa, self.ss = Path(raiz), empresa, ss
        self.arquivo, self.desserializar, self.gravar = arquivo, desserializar, gravar
        self.eh_chave, self.ordem = eh_chave, ordem
        self.serializar = serializar
        self.pasta = self.raiz / "_rascunhos_rt_sync" / empresa

    def caminho(self, chave):
        return self.pasta / (hashlib.sha256(chave.encode()).hexdigest() + ".json")

    def ler(self, chave):
        p = self.caminho(chave)
        if not p.exists():
            return None
        estado = json.loads(p.read_text(encoding="utf-8"))
        if (estado.get("versao") != 1 or estado.get("chave") != chave
                or estado.get("empresa") != self.empresa
                or empresa_dados(estado["dados"]) != self.empresa
                or estado.get("hash") != hash_conteudo(estado["dados"])
                or estado.get("estado") not in ("pendente", "conflito", "confirmado", "descartado")):
            raise ValueError("Registro local de sincronização inválido")
        return estado

    def preparar(self, chave, dados, local_apenas=False):
        if empresa_dados(dados) != self.empresa:
            raise ValueError("Rascunho de outra empresa")
        digest = hash_conteudo(dados)
        anterior = self.ler(chave)
        if anterior and anterior["estado"] == "descartado":
            raise ConflitoRascunho("Rascunho descartado; continue em nova chave.")
        if anterior and anterior["estado"] in ("pendente", "conflito"):
            if anterior["hash"] != digest:
                raise ConflitoRascunho("A cópia pendente foi preservada. Restaure a versão desejada para continuar em nova chave.")
            return anterior
        registro = {"versao": 1, "empresa": self.empresa, "chave": chave,
                    "dados": dados, "hash": digest,
                    "base_hash": anterior["hash"] if anterior else None,
                    "estado": "confirmado" if local_apenas else "pendente"}
        # Este arquivo é a cópia autoritativa, gravada antes de qualquer envio.
        escrever_atomico(self.caminho(chave), registro)
        return registro

    def estado(self, chave, estado, erro=None):
        registro = self.ler(chave)
        if registro is None:
            raise ValueError("Cópia persistida não encontrada")
        registro["estado"] = estado
        registro["erro"] = erro
        escrever_atomico(self.caminho(chave), registro)

    def reconstruir(self):
        pendentes, problemas = [], []
        for p in sorted(self.pasta.glob("*.json")):
            try:
                cabecalho = json.loads(p.read_text(encoding="utf-8"))
                r = self.ler(cabecalho["chave"])
                if r["estado"] in ("pendente", "conflito"):
                    pendentes.append(r["chave"])
            except Exception as e:
                problemas.append(f"Pendência local ilegível: {type(e).__name__}: {e}")
        self.ss["_rt_rasc_sync_pendente"] = pendentes
        if pendentes or problemas:
            self.ss["_rt_rasc_aviso"] = "; ".join(problemas) or "Há cópias locais pendentes; versões divergentes exigem restauração explícita."
        else:
            self.ss.pop("_rt_rasc_sync_pendente", None)
            self.ss.pop("_rt_rasc_aviso", None)
        return pendentes

    def consultar(self, aba, chave):
        if aba is None:
            raise ConnectionError("Consulta à planilha indisponível; ausência não confirmada")
        rows = aba.get_all_values()  # falha de consulta nunca equivale a ausência
        iguais = [(i, row) for i, row in enumerate(rows[1:], 2) if row and str(row[0]).strip() == chave]
        for _, row in iguais:
            d = self.desserializar(row)
            if d is not None and empresa_dados(d) != self.empresa:
                raise ConflitoRascunho("Chave remota pertence a outra empresa; envio bloqueado")
        if any(len(row) > 3 and str(row[3]).lower().startswith("descartado") for _, row in iguais):
            return "descartado", iguais
        for _, row in iguais:
            d = self.desserializar(row)
            if d is None:
                raise ValueError("Conteúdo remoto ilegível; envio bloqueado")
            if empresa_dados(d) != self.empresa:
                raise ConflitoRascunho("Chave remota pertence a outra empresa; envio bloqueado")
            hash_conteudo(d)
        return ("presente" if iguais else "ausente"), iguais

    def enviar(self, aba, chave, dados, retentativa=True):
        r = self.ler(chave)
        if r is None or r["hash"] != hash_conteudo(dados):
            raise ValueError("Envio sem cópia local íntegra persistida")
        if r["estado"] == "descartado":
            return False
        try:
            cache = json.loads(self.arquivo(chave).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            cache = None  # o registro íntegro separado continua autoritativo
        if cache and empresa_dados(cache) == self.empresa and cache.get("_descartado"):
            self.descartar(chave, cache["_descartado"])
            return False
        situacao, rows = self.consultar(aba, chave)
        if situacao == "descartado":
            self.estado(chave, "descartado")
            return False
        hashes = [hash_conteudo(self.desserializar(row)) for _, row in rows]
        if hashes and all(h == r["hash"] for h in hashes):
            self.estado(chave, "confirmado")
            return True
        # Só o autosave normal pode atualizar sua base confirmada. Todas as
        # retentativas, inclusive em nova sessão, recusam conteúdo diferente.
        if hashes and (retentativa or not r.get("base_hash") or any(h != r["base_hash"] for h in hashes)):
            self.estado(chave, "conflito", "Conteúdo remoto divergente")
            raise ConflitoRascunho("Conteúdo remoto divergente; ambas as cópias foram preservadas. Restaure a versão desejada.")
        if situacao == "ausente":
            celulas = self.serializar(dados)
            # Nunca converte ausência em atualização após outra leitura. Se uma
            # chave concorrente surgir, as duas linhas ficam preservadas.
            aba.append_row([chave, dados.get("salvo_em", ""), celulas[0], ""] + celulas[1:],
                           value_input_option="RAW", insert_data_option="INSERT_ROWS", table_range="A1")
        elif not self.gravar(aba, chave, dados):
            self.estado(chave, "descartado")
            return False
        situacao, depois = self.consultar(aba, chave)
        if situacao == "descartado":
            self.estado(chave, "descartado")
            return False
        if not depois or any(hash_conteudo(self.desserializar(row)) != r["hash"] for _, row in depois):
            raise ConflitoRascunho("Não foi possível confirmar o conteúdo remoto; cópia local preservada")
        self.estado(chave, "confirmado")
        return True

    def listar(self, aba, incluir_descartados=False):
        versoes, descartados, mensagens = {}, set(), []
        ignorados = 0
        def adicionar(chave, dados, origem, linha=None, pendente=False):
            if not self.eh_chave(chave) or empresa_dados(dados) != self.empresa:
                return
            digest = hash_conteudo(dados)
            if dados.get("_descartado"):
                descartados.add(chave)
            v = {"chave": chave, "dados": dados, "linha": linha, "origem": origem,
                 "hash": digest, "pendente": pendente, "descartado": False}
            anterior = versoes.get((chave, digest))
            if anterior is None or origem == "planilha" or (pendente and anterior["origem"] != "planilha"):
                versoes[(chave, digest)] = v
        if aba is not None:
            try:
                rows = aba.get_all_values()
            except Exception as e:
                rows = []
                mensagens.append(f"Falha de consulta à planilha: {e}")
            for i, row in enumerate(rows[1:], 2):
                if not row or not self.eh_chave(str(row[0]).strip()):
                    continue
                chave = str(row[0]).strip()
                d = self.desserializar(row)
                try:
                    if d is not None and empresa_dados(d) != self.empresa:
                        continue
                    if len(row) > 3 and str(row[3]).lower().startswith("descartado"):
                        descartados.add(chave)  # vale mesmo com JSON ilegível
                    if d is None:
                        raise ValueError("Conteúdo remoto ilegível")
                    adicionar(chave, d, "planilha", i)
                except Exception as e:
                    ignorados += 1
                    mensagens.append(f"Rascunho remoto ilegível: {e}")
        elif self.ss.get("_rt_rasc_sync_pendente"):
            mensagens.append("Consulta remota indisponível; cópias locais não comprovam ausência na planilha.")
        envelopes = {}
        for p in sorted(self.pasta.glob("*.json")):
            try:
                chave = json.loads(p.read_text(encoding="utf-8"))["chave"]
                r = self.ler(chave)
                envelopes[chave] = r
                if r["estado"] == "descartado":
                    descartados.add(chave)
                adicionar(chave, r["dados"], "local", pendente=r["estado"] in ("pendente", "conflito"))
            except Exception as e:
                ignorados += 1
                mensagens.append(f"Pendência local ilegível: {e}")
        for p in sorted(self.raiz.glob("_rascunho_relatorio_rt*.json")):
            try:
                d = json.loads(p.read_text(encoding="utf-8"))
                chave = d.get("_chave") or ("thyago" if p.name == "_rascunho_relatorio_rt_em_andamento.json" else p.stem)
                if empresa_dados(d) == self.empresa and d.get("_descartado"):
                    descartados.add(chave)
                if chave in envelopes:
                    continue  # cache não substitui a cópia autoritativa
                adicionar(chave, d, "local" if aba is not None else "desenvolvimento")
            except Exception:
                ignorados += 1
        for chave in descartados:
            try:
                r = self.ler(chave)
                if r is None or r["estado"] != "descartado":
                    self.descartar(chave, "descarte confirmado na listagem")
            except Exception as e:
                mensagens.append(f"Não foi possível persistir o descarte local: {e}")
        lista = list(versoes.values())
        for v in lista:
            v["descartado"] = v["chave"] in descartados
            v["conflito"] = sum(k[0] == v["chave"] for k in versoes) > 1
            if v["conflito"] and v["pendente"] and not v["descartado"]:
                try:
                    if self.ler(v["chave"])["estado"] != "conflito":
                        self.estado(v["chave"], "conflito", "Conteúdo remoto divergente")
                except Exception as e:
                    mensagens.append(f"Não foi possível registrar o conflito local: {e}")
        lista.sort(key=lambda v: (bool(v["pendente"]), self.ordem(v["dados"], v["linha"] or 0)), reverse=True)
        return {"versoes": [v for v in lista if incluir_descartados or not v["descartado"]],
                "ignorados": ignorados, "avisos": mensagens, "empresa": self.empresa}

    def descartar(self, chave, marca):
        r = self.ler(chave)
        if r is None:
            p = self.arquivo(chave)
            dados = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"empresa_ativa": self.empresa, "campos": {}}
            if empresa_dados(dados) != self.empresa:
                raise ValueError("Cópia local pertence a outra empresa; descarte bloqueado")
            r = {"versao": 1, "empresa": self.empresa, "chave": chave, "dados": dados,
                 "hash": hash_conteudo(dados), "base_hash": None}
        r["estado"] = "descartado"
        r["marca_descarte"] = marca
        escrever_atomico(self.caminho(chave), r)
