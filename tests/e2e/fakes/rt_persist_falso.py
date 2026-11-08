"""Duplo em memória do contrato rt_persist (database/rt/001_contract.sql) para testes locais.

Reproduz save/history: escopo por principal (42501), validação do contrato e da
identidade (22023), idempotência por operation, conflito por revisão-base e
efeito somente após COMMIT. Não substitui a validação no PostgreSQL real.
"""
import copy
import hashlib
import json
import uuid

from psycopg import errors


def _jsonb(valor):
    obj = getattr(valor, "obj", valor)
    return json.loads(json.dumps(obj, ensure_ascii=False, default=str))


class ContratoFalso:
    def __init__(self):
        self.escopos = set()        # (principal, company, condo, period)
        self.documentos = {}        # (c, d, p) -> head
        self.revisoes = {}          # (c, d, p) -> [dict]
        self.pedidos = {}           # (c, d, p, op) -> dict
        self.conexoes = []

    def conceder(self, principal, company, condo, period):
        self.escopos.add((principal, uuid.UUID(str(company)), uuid.UUID(str(condo)), period))

    def conectar(self, principal):
        c = ConexaoFalsa(self, principal)
        self.conexoes.append(c)
        return c

    def conflitos(self, c, d, p):
        return [r for (cc, dd, pp, _), r in self.pedidos.items() if (cc, dd, pp) == (c, d, p) and r["outcome"] == "conflict"]


class _Cursor:
    def __init__(self, valor):
        self.valor = valor

    def fetchone(self):
        return (self.valor,)


class ConexaoFalsa:
    def __init__(self, contrato, principal):
        self.contrato, self.principal = contrato, principal
        self._pendente = None
        self.fechada = False

    def __enter__(self):
        return self

    def __exit__(self, tipo, valor, tb):
        self._pendente = None if tipo else self._pendente
        if tipo is None and self._pendente is not None:
            self.commit()
        self.fechada = True
        return False

    def commit(self):
        if self._pendente is not None:
            (self.contrato.documentos, self.contrato.revisoes, self.contrato.pedidos) = self._pendente
            self._pendente = None

    def rollback(self):
        self._pendente = None

    def _estado(self):
        if self._pendente is None:
            k = self.contrato
            self._pendente = (dict(k.documentos), copy.deepcopy(k.revisoes), dict(k.pedidos))
        return self._pendente

    def _escopo(self, c, d, p):
        if (self.principal, c, d, p) not in self.contrato.escopos:
            raise errors.InsufficientPrivilege("scope denied")

    def execute(self, consulta, params):
        if consulta.startswith("SELECT rt_persist.save("):
            if consulta.rstrip().endswith(",NULL)"):
                params = (*params, None)  # restore_revision literal NULL
            return _Cursor(self._save(*params))
        if consulta.startswith("SELECT rt_persist.history("):
            c, d, p = params
            self._escopo(c, d, p)
            docs, revs, _ = self._pendente or (self.contrato.documentos, self.contrato.revisoes, self.contrato.pedidos)
            return _Cursor(copy.deepcopy(revs.get((c, d, p), [])))
        raise AssertionError("consulta fora do contrato: " + consulta)

    def _save(self, c, d, p, op, base, body, restore):
        self._escopo(c, d, p)
        body = _jsonb(body)
        if (op is None or base is None or base < 0 or not isinstance(body, dict)
                or not isinstance(body.get("campos"), dict)
                or len(json.dumps(body).encode()) > 8388608):
            raise errors.InvalidParameterValue("invalid save contract")
        if "identity" in body and body["identity"] != {"company_id": str(c), "condominium_id": str(d),
                                                        "period": p.isoformat()}:
            raise errors.InvalidParameterValue("identity mismatch")
        docs, revs, pedidos = self._estado()
        h = docs.setdefault((c, d, p), 0)
        prior = pedidos.get((c, d, p, op))
        if prior:
            if (prior["base_revision"] != base or prior["payload"] != body
                    or prior["restored_from"] != restore or prior["author"] != self.principal):
                raise errors.InvalidParameterValue("idempotency key reused with different request")
            return copy.deepcopy(prior["result"])
        if base > h:
            raise errors.InvalidParameterValue("base revision does not exist")
        registro = {"base_revision": base, "payload": body, "restored_from": restore,
                    "author": self.principal, "observed_head": h}
        if base != h:
            registro.update(outcome="conflict", result={"status": "conflict", "request_id": str(op), "observed_head": h})
            pedidos[(c, d, p, op)] = registro
            return copy.deepcopy(registro["result"])
        revs.setdefault((c, d, p), []).append({
            "company_id": str(c), "condominium_id": str(d), "period": p.isoformat(), "revision": h + 1,
            "base_revision": h, "restored_from": restore, "schema_version": 1, "payload": body,
            "content_hash": hashlib.sha256(json.dumps(body, sort_keys=True).encode()).hexdigest(),
            "author": self.principal})
        docs[(c, d, p)] = h + 1
        registro.update(outcome="saved", result={"status": "saved", "request_id": str(op), "revision": h + 1,
                                                 "sync_state": "pending"})
        pedidos[(c, d, p, op)] = registro
        return copy.deepcopy(registro["result"])
