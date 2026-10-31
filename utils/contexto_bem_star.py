"""Rascunhos do relatório Bem Star separados por cliente e período."""
from copy import deepcopy


def preparar_contexto_bem_star(estado, cliente, mes, ano, id_cliente=""):
    contexto = (str(id_cliente).strip(), str(cliente).strip(),
                str(mes).strip().zfill(2), str(ano).strip())
    anterior = estado.get("_bs_rel_contexto")
    if anterior == contexto:
        return False

    def pertence(chave):
        return chave.startswith("bs_agenda_") or chave in {
            "csr_operador_rel", "csr_obs_rel", "bs_rel_parecer_tecnico",
            "bs_rel_freq_semana", "bs_rel_dias_semana", "bs_rel_incluir_visitas_vazias",
        }

    rascunhos = dict(estado.get("_bs_rel_rascunhos_por_contexto", {}))
    campos = [k for k in estado if pertence(k)]
    if anterior is not None:
        rascunhos[anterior] = deepcopy({k: estado[k] for k in campos})
    for chave in campos:
        del estado[chave]
    estado.update(deepcopy(rascunhos.get(contexto, {})))
    estado["_bs_rel_rascunhos_por_contexto"] = rascunhos
    estado["_bs_rel_contexto"] = contexto
    return anterior is not None
