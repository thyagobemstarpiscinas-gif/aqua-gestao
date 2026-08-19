import streamlit as st
from datetime import datetime


# ============================================================
# MÓDULO ISOLADO — ÁGUA PARA CONSUMO HUMANO
# Não alterar funções Aqua Gestão RT / Bem Star Manutenção.
# Namespace exclusivo: agua_*
# Abas exclusivas: AGUA_*
# ============================================================

AGUA_ABAS_INICIAIS = {
    "AGUA_CLIENTES": [
        "agua_cliente_id",
        "cliente_id",
        "nome",
        "cnpj",
        "endereco",
        "responsavel_local",
        "telefone",
        "email",
        "tipo_abastecimento",
        "concessionaria",
        "possui_fonte_alternativa",
        "numero_torres",
        "numero_reservatorios",
        "status_programa",
        "data_inicio_programa",
        "observacoes",
        "criado_em",
        "atualizado_em",
    ],
    "AGUA_RESERVATORIOS": [
        "reservatorio_id",
        "agua_cliente_id",
        "codigo",
        "torre",
        "localizacao",
        "tipo",
        "capacidade_l",
        "material",
        "fonte_abastecimento",
        "data_ultima_higienizacao",
        "empresa_higienizacao",
        "comprovante_higienizacao",
        "status",
        "observacoes",
        "criado_em",
        "atualizado_em",
    ],
    "AGUA_PONTOS": [
        "ponto_id",
        "agua_cliente_id",
        "reservatorio_id",
        "codigo",
        "tipo_ponto",
        "descricao",
        "localizacao",
        "representa_entrada_rede",
        "ativo",
        "observacoes",
        "criado_em",
        "atualizado_em",
    ],
    "AGUA_VISITAS": [
        "visita_id",
        "agua_cliente_id",
        "data",
        "hora_inicio",
        "hora_fim",
        "profissional",
        "tipo_visita",
        "motivo",
        "situacao",
        "observacoes",
        "criado_em",
        "atualizado_em",
    ],
}


def agua_agora():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def agua_gerar_id(prefixo: str, existentes: list[str]) -> str:
    maior = 0
    for valor in existentes or []:
        texto = str(valor or "").strip()
        if not texto.startswith(prefixo + "-"):
            continue
        try:
            numero = int(texto.split("-")[-1])
            maior = max(maior, numero)
        except Exception:
            pass
    return f"{prefixo}-{maior + 1:06d}"


def agua_obter_ou_criar_aba(conectar_sheets, obter_aba_sheets, nome_aba: str):
    """Obtém ou cria exclusivamente abas AGUA_*.

    Consulta diretamente o objeto da planilha para evitar cache de worksheet
    inexistente. Não altera obter_aba_sheets() nem qualquer função dos módulos
    Aqua Gestão RT / Bem Star Manutenção.
    """
    if nome_aba not in AGUA_ABAS_INICIAIS:
        raise ValueError(f"Aba não autorizada no módulo Água: {nome_aba}")

    sh = conectar_sheets()
    if sh is None:
        return None

    # Busca direta, sem usar o cache da função global.
    try:
        return sh.worksheet(nome_aba)
    except Exception as erro:
        # Só continua para criação quando a worksheet realmente não existe.
        try:
            import gspread
            if not isinstance(erro, gspread.exceptions.WorksheetNotFound):
                raise
        except ImportError:
            # gspread já faz parte da aplicação; fallback conservador.
            if "not found" not in str(erro).lower():
                raise

    cabecalho = AGUA_ABAS_INICIAIS[nome_aba]

    try:
        aba = sh.add_worksheet(
            title=nome_aba,
            rows=1000,
            cols=max(len(cabecalho), 10),
        )
    except Exception:
        # Proteção contra concorrência/rerun:
        # se outra execução acabou de criar a aba, tenta buscá-la novamente.
        aba = sh.worksheet(nome_aba)

    ultima_coluna = ""
    numero = len(cabecalho)
    while numero:
        numero, resto = divmod(numero - 1, 26)
        ultima_coluna = chr(65 + resto) + ultima_coluna

    valores_atuais = aba.row_values(1)
    if not valores_atuais:
        aba.update(
            range_name=f"A1:{ultima_coluna}1",
            values=[cabecalho],
            value_input_option="RAW",
        )

    return aba


def agua_inicializar_banco(conectar_sheets, obter_aba_sheets):
    resultados = {}
    for nome_aba in AGUA_ABAS_INICIAIS:
        try:
            aba = agua_obter_ou_criar_aba(
                conectar_sheets,
                obter_aba_sheets,
                nome_aba,
            )
            resultados[nome_aba] = aba is not None
        except Exception:
            resultados[nome_aba] = False
    return resultados


def agua_listar_clientes(conectar_sheets, obter_aba_sheets) -> list[dict]:
    aba = agua_obter_ou_criar_aba(
        conectar_sheets,
        obter_aba_sheets,
        "AGUA_CLIENTES",
    )
    if aba is None:
        return []

    valores = aba.get_all_records()
    return [dict(item) for item in valores if item.get("agua_cliente_id")]


def agua_render_dashboard(conectar_sheets, obter_aba_sheets):
    st.title("💧 Água para Consumo Humano")
    st.caption("Aqua Gestão — Controle Técnico da Qualidade da Água")

    resultados = agua_inicializar_banco(
        conectar_sheets,
        obter_aba_sheets,
    )

    if all(resultados.values()):
        st.success("Estrutura inicial do módulo disponível.")
    else:
        falhas = [nome for nome, ok in resultados.items() if not ok]
        st.error(
            "Não foi possível acessar/criar: " + ", ".join(falhas)
        )

    clientes = agua_listar_clientes(
        conectar_sheets,
        obter_aba_sheets,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Clientes", len(clientes))
    c2.metric("Reservatórios", "—")
    c3.metric("Visitas", "—")

    st.markdown("### Estrutura inicial")
    st.write("• Clientes")
    st.write("• Reservatórios")
    st.write("• Pontos de amostragem")
    st.write("• Visitas")

    st.info(
        "Próxima etapa: Nova Avaliação — inspeção, medições, "
        "amostras e rastreabilidade."
    )


def agua_render_modulo(conectar_sheets, obter_aba_sheets):
    """Entrada principal isolada do módulo Água para Consumo Humano."""

    if st.button(
        "← Voltar à tela inicial",
        key="agua_btn_voltar_inicio",
    ):
        st.session_state["modo_atual"] = "entrada"
        st.session_state["agua_logado"] = False
        st.rerun()

    agua_render_dashboard(
        conectar_sheets,
        obter_aba_sheets,
    )
