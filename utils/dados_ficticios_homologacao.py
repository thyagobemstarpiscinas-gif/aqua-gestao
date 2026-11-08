"""Dados FICTÍCIOS da planilha de homologação (OAuth drive.file).

Conjunto único usado pelo script de preparação e pelos testes de compatibilidade:
três condomínios (Alfa e Beta — Aqua Gestão; Gama — Bem Star), quatro visitas
(todas em terças-feiras, como exige a importação de visitas do RT) e um operador
INATIVO e SEM PIN (não autentica; serve para conferir a listagem administrativa).

O esquema das abas reproduz o que o app lê e grava:
- 👥 Clientes: colunas A–P de sheets_salvar_cliente / sheets_listar_clientes_completo.
  O campo Empresa (M) usa os valores gravados pelo app ("Aqua Gestão",
  "Bem Star Piscinas"): o filtro por empresa reconhece o texto "aqua".
- 🔬 Visitas: colunas A–Z de sheets_salvar_lancamento_campo; Payload JSON na W.
- 👷 Operadores: Nome | PIN | Condomínios (|) | Ativo | Cadastrado_em | Obs.
Não contém dados reais, PINs ou credenciais.
"""
from __future__ import annotations

import json

MARCADOR = "DADO FICTÍCIO — HOMOLOGAÇÃO"

ABA_CLIENTES = "👥 Clientes"
ABA_VISITAS = "🔬 Visitas"
ABA_OPERADORES = "👷 Operadores"

CABECALHO_CLIENTES = [
    "", "ID", "Nome", "Volume_m3", "Telefone", "Contato_Sindico", "Endereco", "Data_Cadastro",
    "Status", "Vol_Adulto_m3", "Vol_Infantil_m3", "Vol_Family_m3", "Empresa", "CNPJ", "CEP",
    "Piscinas_JSON",
]
CABECALHO_VISITAS = [
    "", "ID Visita", "Data", "ID Cliente", "Condomínio",
    "pH", "CRL", "CT", "Alcalinidade", "Dureza", "CYA",
    "Foto Antes", "Foto Depois", "Foto Casa Máquinas", "Observação",
    "Dosagem Cloro", "Dosagem Bicarb", "Alerta pH", "Alerta Cloro",
    "Status", "Operador", "Problemas", "Payload JSON", "Salvo em",
    "Fonte", "Mês/Ano",
]
CABECALHO_OPERADORES = ["Nome", "PIN", "Condomínios (separados por |)", "Ativo", "Cadastrado_em", "Obs"]

CONDOMINIOS = [
    # id, nome, empresa, contato, endereço, vol_adulto, vol_infantil
    ("C001", "Condomínio Alfa Fictício", "Aqua Gestão", "Síndico Alfa Fictício", "Rua Fictícia A, 1 — Uberlândia/MG", 50, 10),
    ("C002", "Residencial Beta Fictício", "Aqua Gestão", "Síndica Beta Fictícia", "Rua Fictícia B, 2 — Uberlândia/MG", 40, 0),
    ("C003", "Condomínio Gama Fictício", "Bem Star Piscinas", "Síndico Gama Fictício", "Rua Fictícia C, 3 — Uberlândia/MG", 60, 12),
]

VISITAS = [
    # id, data (terça), cliente, operador, piscinas [(nome, ph, crl, ct, alc, dureza, cya)]
    ("V00001", "29/09/2026", "C001", "OPERADOR FICTÍCIO",
     [("Piscina Adulto", "7,4", "2,5", "2,8", "100", "200", "40"), ("Piscina Infantil", "7,5", "2,2", "2,4", "", "", "")]),
    ("V00002", "06/10/2026", "C001", "OPERADOR FICTÍCIO",
     [("Piscina Adulto", "7,1", "0,8", "1,6", "70", "120", "80")]),
    ("V00003", "13/10/2026", "C002", "OPERADOR FICTÍCIO",
     [("Piscina Adulto", "7,6", "3,0", "3,2", "110", "220", "45")]),
    ("V00004", "20/10/2026", "C003", "OPERADOR BEM STAR FICTÍCIO",
     [("Piscina Adulto", "7,3", "2,1", "2,3", "95", "210", "35"), ("Piscina Infantil", "7,4", "1,9", "2,0", "", "", "")]),
]

OPERADOR_INATIVO = ["OPERADOR INATIVO FICTÍCIO", "", "Condomínio Alfa Fictício", "Não", "2026-11-01", MARCADOR]


def _nome(cliente_id: str) -> str:
    return next(c[1] for c in CONDOMINIOS if c[0] == cliente_id)


def linhas_clientes() -> list[list[str]]:
    linhas = [list(CABECALHO_CLIENTES)]
    for cid, nome, empresa, contato, endereco, adulto, infantil in CONDOMINIOS:
        linhas.append([
            "", cid, nome, str(adulto + infantil), "", contato, endereco, "2026-11-01", "Ativo",
            str(adulto) if adulto else "", str(infantil) if infantil else "", "", empresa, "", "", "",
        ])
    return linhas


def linhas_visitas() -> list[list[str]]:
    linhas = [list(CABECALHO_VISITAS)]
    for vid, data, cid, operador, piscinas in VISITAS:
        nome = _nome(cid)
        p0 = piscinas[0]
        payload = {
            "id_visita": vid, "data": data, "condominio": nome, "status": "Concluída",
            "operador": operador, "observacao": MARCADOR, "fotos": [], "dosagens": [],
            "piscinas": [
                {"nome": p[0], "ph": p[1], "cloro_livre": p[2], "cloro_total": p[3],
                 "alcalinidade": p[4], "dureza": p[5], "cianurico": p[6]} for p in piscinas
            ],
        }
        dia, mes, ano = data.split("/")
        linhas.append([
            "", vid, data, cid, nome, p0[1], p0[2], p0[3], p0[4], p0[5], p0[6],
            "", "", "", MARCADOR, "", "", "", "", "Concluída", operador, "",
            json.dumps(payload, ensure_ascii=False), "01/11/2026 09:00:00", "Homologação fictícia", f"{mes}/{ano}",
        ])
    return linhas


def linhas_operadores() -> list[list[str]]:
    return [list(CABECALHO_OPERADORES), list(OPERADOR_INATIVO)]


def abas() -> dict[str, list[list[str]]]:
    return {ABA_CLIENTES: linhas_clientes(), ABA_VISITAS: linhas_visitas(), ABA_OPERADORES: linhas_operadores()}


def faixa(linhas: list[list[str]]) -> str:
    """Intervalo A1 que cobre as linhas (ex.: A1:P4)."""
    largura = max(len(l) for l in linhas)
    col, n = "", largura
    while n:
        n, r = divmod(n - 1, 26)
        col = chr(65 + r) + col
    return f"A1:{col}{len(linhas)}"
