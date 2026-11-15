"""Conferência de preenchimento do mensal; não altera leituras ou conformidade."""
from calendar import monthrange
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from utils.bem_star_mensal import (
    PARAMETROS, canonizar_nomes_piscinas, chave_piscina, data_ordem, expandir,
    filtrar_analises_semanais, numero, texto_cloro_combinado,
)


def conferir_analises(lancamentos, piscinas, mes, ano, hoje=None):
    """Uma coleta por semana por piscina; usa somente a fonte semanal do PDF.

    Segundas do mês são referências de semana, nunca datas impostas à coleta.
    A primeira semana parcial não cria uma obrigação antes do início do mês.
    Semanas ainda abertas/futuras não geram aviso de ausência. Leituras de todo
    o mês, inclusive a primeira semana parcial, aparecem na conferência.
    """
    try:
        mes, ano = int(mes), int(ano)
        inicio = date(ano, mes, 1)
        fim = date(ano, mes, monthrange(ano, mes)[1])
    except (TypeError, ValueError, OverflowError):
        return {"erro": "Informe um mês e um ano válidos para conferir as análises."}
    hoje = hoje or datetime.now(ZoneInfo("America/Sao_Paulo")).date()
    semana = inicio + timedelta(days=(-inicio.weekday()) % 7)
    semanas = []
    while semana <= fim:
        ultimo = min(semana + timedelta(days=6), fim)
        semanas.append({"inicio": semana, "fim": ultimo,
                        "encerrada": ultimo < hoje,
                        "rotulo": f"{semana:%d/%m} a {ultimo:%d/%m}"})
        semana += timedelta(days=7)

    linhas = [l for l in expandir(filtrar_analises_semanais(lancamentos), piscinas)
              if (data_ordem(l.get("data")).year, data_ordem(l.get("data")).month) == (ano, mes)]
    nomes = canonizar_nomes_piscinas(list(piscinas or []) + [l["piscina"] for l in linhas], piscinas)
    resumos, detalhes, leituras = [], [], []
    for nome in nomes:
        registros = [l for l in linhas if chave_piscina(l["piscina"]) == chave_piscina(nome)]
        datas_com_leitura = set()
        for linha in registros:
            ausentes, invalidos = [], []
            tem_leitura = False
            leitura = {"Piscina": nome, "Data": linha["data"]}
            for campo, titulo, unidade in PARAMETROS:
                valor = linha.get(campo)
                num = numero(valor)
                if valor is None or not str(valor).strip():
                    ausentes.append(titulo)
                elif num is None or num < 0:
                    invalidos.append(titulo)
                else:
                    tem_leitura = True
                leitura[titulo + (f" ({unidade})" if unidade else "")] = (
                    "—" if valor is None or not str(valor).strip() else str(valor))
                if campo == "cloro_total":
                    leitura["Cloro combinado (mg/L)"] = texto_cloro_combinado(linha)
            if tem_leitura:
                datas_com_leitura.add(data_ordem(linha["data"]).date())
            leituras.append(leitura)
            if ausentes or invalidos or texto_cloro_combinado(linha) == "Inconsistente":
                detalhes.append({"Piscina": nome, "Data": linha["data"],
                                 "Campos ausentes": ", ".join(ausentes) or "—",
                                 "Valores inválidos": ", ".join(invalidos) or "—",
                                 "Cloro combinado": texto_cloro_combinado(linha)})
        faltantes = [s["rotulo"] for s in semanas if s["encerrada"]
                     and not any(s["inicio"] <= d <= s["fim"] for d in datas_com_leitura)]
        resumos.append({"Piscina": nome, "Datas com análise": len(datas_com_leitura),
                        "Datas disponíveis": ", ".join(f"{d:%d/%m/%Y}" for d in sorted(datas_com_leitura)) or "Nenhuma",
                        "Semanas encerradas sem leitura": "; ".join(faltantes) or "Nenhuma",
                        "semanas_faltantes": len(faltantes)})
    return {"erro": "", "piscinas": resumos, "pendencias": detalhes, "leituras": leituras,
            "semanas_encerradas": sum(s["encerrada"] for s in semanas),
            "semanas_em_aberto": sum(not s["encerrada"] for s in semanas),
            "tem_pendencias": bool(detalhes or any(p["semanas_faltantes"] for p in resumos))}


def renderizar_conferencia(st, lancamentos, piscinas, mes, ano, hoje=None):
    st.markdown("### Conferência das análises semanais")
    st.caption("Estas são as leituras que alimentam as tabelas, os gráficos e o resumo numérico do PDF. "
               "A consulta abaixo não altera as visitas salvas.")
    dados = conferir_analises(lancamentos, piscinas, mes, ano, hoje)
    if dados["erro"]:
        st.info(dados["erro"])
        return
    if dados["tem_pendencias"]:
        st.warning("Há semanas sem leitura ou análises com campos incompletos/inconsistentes. "
                   "Confira os detalhes abaixo antes de gerar o mensal.")
    elif dados["leituras"]:
        st.success("Nenhuma pendência de preenchimento nas análises e semanas encerradas conferidas.")
    st.caption("Referência: uma análise por piscina por semana, de segunda a domingo, a partir da primeira "
               "segunda-feira do mês. Coletas em outros dias atendem à mesma semana e mantêm sua data real. "
               "Semanas em andamento ou futuras não são apontadas como faltantes. "
               "Na última semana, a conferência termina no último dia do mês.")
    if dados["piscinas"]:
        st.dataframe([{k: v for k, v in p.items() if k != "semanas_faltantes"}
                      for p in dados["piscinas"]], hide_index=True, width="stretch")
    st.caption(f"{dados['semanas_encerradas']} semana(s) encerrada(s) e "
               f"{dados['semanas_em_aberto']} em andamento/futura(s). "
               "Este aviso confere o preenchimento, não atesta a qualidade da água nem bloqueia a emissão.")
    if dados["pendencias"]:
        with st.expander("Ver campos ausentes e valores inconsistentes", expanded=True):
            st.dataframe(dados["pendencias"], hide_index=True, width="stretch")
    if dados["leituras"]:
        with st.expander("Consultar as medições usadas no mensal"):
            st.dataframe(dados["leituras"], hide_index=True, width="stretch")
