"""Rascunhos RT no contrato PostgreSQL ``rt_persist`` (opt-in, somente homologação).

Sem Streamlit e sem credenciais próprias: a conexão vem de uma fábrica injetada
(``conectar(principal)``), que em homologação é o executor verificado
(``database/rt/supabase_runner.py``: TLS verify-full, marcador, identidade real,
JIT/DPAPI) e nos testes locais um PostgreSQL de loopback.

Modelo: um documento por empresa + cliente + período (``rel_origem_dados``),
com revisões imutáveis. Cada gravação informa a revisão-base vista pela sessão;
se outra sessão gravou antes, o banco registra conflito e nada é sobrescrito.
Somente ``rt_persist.save`` e ``rt_persist.history`` são usados.
"""
from dataclasses import dataclass
import datetime
import hashlib
import json
import re
import uuid

# Namespace fixo: o mesmo cliente produz sempre o mesmo condominium_id.
NAMESPACE = uuid.UUID("6f2d1c8e-4a7b-5e39-9c1d-acc0a6e57a01")
EMPRESAS = ("aqua_gestao", "bem_star")
# Fábricas de conexão registradas somente por testes locais (loopback). Produção não registra.
FABRICAS = {}
_ORIGEM = re.compile(r"^(aqua_gestao|bem_star)\|((?:id|nome):[^|]+)\|(\d{2})/(\d{4})$")


class RepositorioIndisponivel(RuntimeError):
    """Configuração ausente/inválida ou banco inacessível; nunca equivale a ausência."""


class ConflitoVersao(ValueError):
    """Outra sessão gravou uma revisão mais nova; nada foi sobrescrito."""


@dataclass(frozen=True)
class Identidade:
    empresa: str
    cliente: str
    company_id: uuid.UUID
    condominium_id: uuid.UUID
    period: datetime.date
    principal: str

    @property
    def doc(self):
        return f"{self.company_id}|{self.condominium_id}|{self.period.isoformat()}"

    def envelope(self):
        return {"company_id": str(self.company_id), "condominium_id": str(self.condominium_id),
                "period": self.period.isoformat()}


def validar_config(cfg):
    """{"empresas": {empresa: {"company_id": uuid, "principal": login}}}; sem segredos."""
    if not isinstance(cfg, dict) or not isinstance(cfg.get("empresas"), dict) or not cfg["empresas"]:
        raise RepositorioIndisponivel("Configuração PostgreSQL de rascunhos sem empresas.")
    principais = set()
    for empresa, item in cfg["empresas"].items():
        if empresa not in EMPRESAS or not isinstance(item, dict):
            raise RepositorioIndisponivel("Empresa desconhecida na configuração PostgreSQL.")
        uuid.UUID(str(item.get("company_id")))
        principal = item.get("principal")
        if not isinstance(principal, str) or not re.fullmatch(r"[a-z_][a-z0-9_]{0,62}", principal):
            raise RepositorioIndisponivel("Principal PostgreSQL inválido.")
        if principal in principais:
            raise RepositorioIndisponivel("Cada empresa precisa de um principal próprio.")
        principais.add(principal)
    return cfg


def condominium_id(empresa, cliente):
    return uuid.uuid5(NAMESPACE, f"{empresa}|{cliente}")


def identidade(cfg, empresa, origem):
    """Identidade relacional a partir de rel_origem_dados; recusa o que não for comprovável."""
    m = _ORIGEM.match(str(origem or ""))
    if not m or m.group(1) != empresa:
        raise ValueError("Rascunho sem identidade de empresa, cliente e período; não enviado ao PostgreSQL.")
    mes, ano = int(m.group(3)), int(m.group(4))
    if not 1 <= mes <= 12 or not 2000 <= ano <= 2100:
        raise ValueError("Período do rascunho inválido.")
    item = cfg["empresas"].get(empresa)
    if item is None:
        raise RepositorioIndisponivel("Empresa sem principal PostgreSQL configurado.")
    cliente = m.group(2)
    return Identidade(empresa, cliente, uuid.UUID(str(item["company_id"])),
                      condominium_id(empresa, cliente), datetime.date(ano, mes, 1), item["principal"])


def _canonico(obj):
    return json.dumps(obj, sort_keys=True, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def corpo(dados, ident):
    from utils.rascunho_sync import empresa_dados
    if empresa_dados(dados) != ident.empresa:
        raise ValueError("Rascunho de outra empresa.")
    if not isinstance(dados.get("campos"), dict):
        raise ValueError("Rascunho sem campos.")
    body = {k: v for k, v in dados.items() if not str(k).startswith("_pg")}
    body["identity"] = ident.envelope()  # o banco recusa identidade divergente do envelope
    return body


def operacao(chave, body, base):
    """Mesma chave + mesmo conteúdo + mesma base = mesma operação (reenvio idempotente)."""
    return uuid.uuid5(NAMESPACE, _canonico([chave, hashlib.sha256(_canonico(body).encode()).hexdigest(), base]))


class Repositorio:
    def __init__(self, cfg, conectar):
        self.cfg = validar_config(cfg)
        self.conectar = conectar

    def identidade(self, empresa, origem):
        return identidade(self.cfg, empresa, origem)

    def _executar(self, ident, consulta, parametros):
        from psycopg import errors
        try:
            c = self.conectar(ident.principal)
        except Exception as e:
            raise RepositorioIndisponivel(f"PostgreSQL indisponível ({type(e).__name__}).") from None
        try:
            with c:
                valor = c.execute(consulta, parametros).fetchone()[0]
                c.commit()  # conflito também precisa de COMMIT para ficar registrado
                return valor
        except errors.InsufficientPrivilege:
            raise PermissionError("Sem escopo para este condomínio e período no PostgreSQL.") from None
        except errors.Error as e:
            raise RepositorioIndisponivel(f"Falha PostgreSQL ({type(e).__name__}, SQLSTATE {e.sqlstate}).") from None

    def salvar(self, chave, dados, base):
        from psycopg.types.json import Jsonb
        if isinstance(base, bool) or not isinstance(base, int) or base < 0:
            raise ValueError("Revisão-base inválida.")
        empresa = dados.get("empresa_ativa", "aqua_gestao")
        ident = self.identidade(empresa, (dados.get("campos") or {}).get("rel_origem_dados"))
        body = corpo(dados, ident)
        op = operacao(chave, body, base)
        r = self._executar(ident, "SELECT rt_persist.save(%s,%s,%s,%s,%s,%s,NULL)",
                           (ident.company_id, ident.condominium_id, ident.period, op, base, Jsonb(body)))
        if r.get("status") == "saved":
            return {"status": "saved", "revision": int(r["revision"]), "doc": ident.doc}
        if r.get("status") == "conflict":
            raise ConflitoVersao(f"Versão divergente: o PostgreSQL já está na revisão {r.get('observed_head')}; "
                                 "nada foi sobrescrito.")
        raise RepositorioIndisponivel("Resposta inesperada do contrato PostgreSQL.")

    def ultima(self, empresa, origem):
        """(dados, revisão) da revisão mais recente, ou None se o documento não tem revisões."""
        ident = self.identidade(empresa, origem)
        revisoes = self._executar(ident, "SELECT rt_persist.history(%s,%s,%s)",
                                  (ident.company_id, ident.condominium_id, ident.period))
        if not revisoes:
            return None
        ultima = revisoes[-1]
        dados = dict(ultima["payload"])
        if dados.pop("identity", None) != ident.envelope():
            raise RepositorioIndisponivel("Revisão com identidade divergente; ignorada.")
        from utils.rascunho_sync import empresa_dados
        if empresa_dados(dados) != empresa or (dados.get("campos") or {}).get("rel_origem_dados") != origem:
            raise RepositorioIndisponivel("Revisão de outra empresa ou condomínio; ignorada.")
        dados["_pg"] = {"doc": ident.doc, "revision": int(ultima["revision"])}
        return dados, int(ultima["revision"])
