"""Rascunhos do relatório Bem Star separados por cliente e período."""
from copy import deepcopy


def preparar_contexto_bem_star(estado, cliente, mes, ano, id_cliente=""):
    contexto = (str(id_cliente).strip(), str(cliente).strip(),
                str(mes).strip().zfill(2), str(ano).strip())
    anterior = estado.get("_bs_rel_contexto")
    if anterior == contexto:
        estado["_bs_rel_troca_contexto"] = False
        return False

    def pertence(chave):
        return chave.startswith(("bs_agenda_", "bs_mensal_")) or chave in {
            "csr_operador_rel", "csr_obs_rel", "bs_rel_parecer_tecnico",
            "bs_rel_freq_semana", "bs_rel_dias_semana", "bs_rel_incluir_visitas_vazias",
        }

    rascunhos = dict(estado.get("_bs_rel_rascunhos_por_contexto", {}))
    campos = [k for k in estado if pertence(k)]
    if anterior is not None:
        rascunhos[anterior] = deepcopy({k: estado[k] for k in campos})
    # Remover a chave não redefine um widget já existente no Streamlit:
    # o navegador pode devolver o valor antigo. Atribua os padrões explicitamente.
    for chave in campos:
        estado[chave] = ""
    estado.update({"csr_operador_rel": "", "csr_obs_rel": "",
                   "bs_rel_parecer_tecnico": "", "bs_rel_freq_semana": 1,
                   "bs_rel_dias_semana": [], "bs_rel_incluir_visitas_vazias": False,
                   "bs_mensal_piscinas": [], "bs_mensal_outras": "", "bs_mensal_fotos": []})
    restaurados = deepcopy(rascunhos.get(contexto, {}))
    estado.update(restaurados)
    estado["_bs_rel_campos_restaurados"] = set(restaurados)
    estado["_bs_rel_troca_contexto"] = True
    estado["_bs_rel_rascunhos_por_contexto"] = rascunhos
    estado["_bs_rel_contexto"] = contexto
    return anterior is not None


def preencher_campo_agenda_bem_star(estado, chave, valor):
    novo = (estado.get("_bs_rel_troca_contexto", False)
            and chave not in estado.get("_bs_rel_campos_restaurados", set()))
    if novo:
        estado[chave] = "" if valor is None else str(valor)
    elif chave not in estado and valor not in (None, ""):
        estado[chave] = str(valor)
