import os
import re
import json
import shutil
from datetime import date, datetime, timedelta
import platform
from pathlib import Path
from urllib.parse import quote

import streamlit as st
from docx import Document
from docx.shared import Inches
from PIL import Image, ImageOps

# =========================================
# INTEGRAÇÃO GOOGLE SHEETS
# =========================================

SHEETS_ID = "1uvZ6qfYCYFl_feGGgvIIXMQlUWvx0MTzTuC8TwfPBlM"
DRIVE_FOTOS_FOLDER_ID = "1KNtPKvLl_NJw-Vm_26ABxc4LG3CiZqDR"


def conectar_drive():
    """Conecta ao Google Drive usando as mesmas credenciais do Sheets."""
    try:
        from googleapiclient.discovery import build
        from google.oauth2.service_account import Credentials

        SCOPES = [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ]

        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            creds_path = Path(__file__).parent / "aqua-gestao-rt-87316ebf5331.json"
            if not creds_path.exists():
                return None
            creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)

        service = build("drive", "v3", credentials=creds)
        return service
    except Exception as e:
        _log_sheets_erro("conectar_drive", e)
        return None


def drive_criar_pasta(nome_pasta: str, pasta_pai_id: str) -> str | None:
    """Cria pasta no Drive se não existir. Retorna o ID da pasta."""
    try:
        service = conectar_drive()
        if not service:
            return None

        # Verifica se já existe
        query = f"name='{nome_pasta}' and '{pasta_pai_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        resultado = service.files().list(q=query, fields="files(id,name)").execute()
        arquivos = resultado.get("files", [])
        if arquivos:
            return arquivos[0]["id"]

        # Cria nova
        meta = {
            "name": nome_pasta,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [pasta_pai_id],
        }
        pasta = service.files().create(body=meta, fields="id").execute()
        return pasta.get("id")
    except Exception as e:
        _log_sheets_erro("drive_criar_pasta", e)
        return None


def drive_upload_foto(arquivo_bytes: bytes, nome_arquivo: str, nome_condominio: str, mes_ano: str = None) -> str | None:
    """Faz upload de foto para o Drive. Retorna o ID do arquivo."""
    try:
        import io
        from googleapiclient.http import MediaIoBaseUpload

        service = conectar_drive()
        if not service:
            return None

        if not mes_ano:
            mes_ano = datetime.now().strftime("%Y-%m")

        # Estrutura: Aqua Gestão – Fotos / Condomínio / Ano-Mês
        pasta_cond = drive_criar_pasta(nome_condominio, DRIVE_FOTOS_FOLDER_ID)
        if not pasta_cond:
            return None
        pasta_mes = drive_criar_pasta(mes_ano, pasta_cond)
        if not pasta_mes:
            return None

        # Detecta tipo
        ext = nome_arquivo.lower().split(".")[-1]
        mime_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp", "heic": "image/heic"}
        mime = mime_map.get(ext, "image/jpeg")

        meta = {"name": nome_arquivo, "parents": [pasta_mes]}
        media = MediaIoBaseUpload(io.BytesIO(arquivo_bytes), mimetype=mime)
        arquivo = service.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
        return arquivo.get("id")
    except Exception as e:
        _log_sheets_erro("drive_upload_foto", e)
        return None


def drive_baixar_foto(file_id: str) -> bytes | None:
    """Baixa foto do Drive pelo ID. Retorna bytes da imagem."""
    try:
        service = conectar_drive()
        if not service:
            return None
        conteudo = service.files().get_media(fileId=file_id).execute()
        return conteudo
    except Exception as e:
        _log_sheets_erro("drive_baixar_foto", e)
        return None



# =========================================
# GESTÃO DE OPERADORES
# =========================================

def _normalizar_chave_acesso(texto: str) -> str:
    """Normaliza nomes para comparação exata de PINs, operadores e condomínios."""
    texto = re.sub(r"\s+", " ", str(texto or "").strip())
    return texto.casefold()


def _condominios_organizar(condominios: list[str] | None) -> list[str]:
    """Limpa, deduplica e preserva a ordem dos condomínios informados."""
    resultado = []
    vistos = set()
    for item in condominios or []:
        valor = re.sub(r"\s+", " ", str(item or "").strip())
        if not valor:
            continue
        chave = _normalizar_chave_acesso(valor)
        if chave in vistos:
            continue
        vistos.add(chave)
        resultado.append(valor)
    return resultado


def _resolver_condominios_permitidos_exatos(condominios_permitidos: list[str], todos_condominios: list[str]) -> list[str]:
    """Resolve permissões por correspondência exata normalizada.

    Mantém o nome oficial disponível no sistema e evita liberações por substring.
    """
    mapa_disponiveis = {}
    for nome in todos_condominios or []:
        chave = _normalizar_chave_acesso(nome)
        if chave and chave not in mapa_disponiveis:
            mapa_disponiveis[chave] = nome

    permitidos_exatos = []
    vistos = set()
    for nome in _condominios_organizar(condominios_permitidos):
        chave = _normalizar_chave_acesso(nome)
        if chave in mapa_disponiveis and chave not in vistos:
            vistos.add(chave)
            permitidos_exatos.append(mapa_disponiveis[chave])
    return permitidos_exatos


def _pin_operador_em_uso(pin: str, nome_ignorar: str = "") -> bool:
    """Verifica se o PIN já está em uso por outro operador."""
    pin_limpo = str(pin or "").strip()
    nome_ignorar_norm = _normalizar_chave_acesso(nome_ignorar)
    if not pin_limpo:
        return False

    for op in sheets_listar_operadores():
        if str(op.get("pin", "")).strip() == pin_limpo:
            if _normalizar_chave_acesso(op.get("nome", "")) != nome_ignorar_norm:
                return True

    for op in carregar_operadores():
        if str(op.get("pin", "")).strip() == pin_limpo:
            if _normalizar_chave_acesso(op.get("nome", "")) != nome_ignorar_norm:
                return True
    return False


def sheets_listar_operadores() -> list[dict]:
    """Lista operadores da aba 👷 Operadores do Sheets."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return []
        try:
            aba = sh.worksheet("👷 Operadores")
        except Exception:
            return []
        todos = aba.get_all_values()
        operadores = []
        for row in todos:
            if len(row) >= 4 and str(row[0]).strip() and str(row[0]).strip() != "Nome":
                nome = re.sub(r"\s+", " ", str(row[0]).strip())
                pin  = str(row[1]).strip()
                conds_raw = str(row[2]).strip()
                ativo = str(row[3]).strip().lower() in ("sim", "ativo", "1", "true", "yes")
                conds = _condominios_organizar(conds_raw.split("|")) if conds_raw else []
                acesso_total = any(_normalizar_chave_acesso(c) == "todos" for c in conds) or not conds
                operadores.append({
                    "nome": nome,
                    "pin": pin,
                    "condomínios": conds,
                    "ativo": ativo,
                    "acesso_total": acesso_total,
                })
        return operadores
    except Exception as e:
        _log_sheets_erro("sheets_listar_operadores", e)
        return []


def sheets_criar_aba_operadores():
    """Cria a aba 👷 Operadores no Sheets se não existir."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return False
        try:
            sh.worksheet("👷 Operadores")
            return True  # já existe
        except Exception:
            pass
        aba = sh.add_worksheet(title="👷 Operadores", rows=100, cols=6)
        # Cabeçalho
        aba.update("A1:F1", [["Nome", "PIN", "Condomínios (separados por |)", "Ativo", "Cadastrado_em", "Obs"]])
        aba.format("A1:F1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.07, "green": 0.16, "blue": 0.46}})
        return True
    except Exception as e:
        _log_sheets_erro("sheets_criar_aba_operadores", e)
        return False


def sheets_salvar_operador(nome: str, pin: str, condomínios: list, ativo: bool = True) -> bool:
    """Salva ou atualiza operador na aba 👷 Operadores."""
    try:
        nome_limpo = re.sub(r"\s+", " ", str(nome or "").strip())
        pin_limpo = str(pin or "").strip()
        conds_limpos = _condominios_organizar(condomínios)

        if not nome_limpo or not pin_limpo:
            st.session_state["_operadores_erro"] = "Nome e PIN são obrigatórios."
            return False

        if _pin_operador_em_uso(pin_limpo, nome_ignorar=nome_limpo):
            st.session_state["_operadores_erro"] = f"O PIN {pin_limpo} já está em uso por outro operador."
            return False

        sh = conectar_sheets()
        if sh is None:
            return False
        sheets_criar_aba_operadores()
        aba = sh.worksheet("👷 Operadores")
        todos = aba.get_all_values()
        conds_str = " | ".join(conds_limpos)
        ativo_str = "Sim" if ativo else "Não"
        nova_linha = [nome_limpo, pin_limpo, conds_str, ativo_str, datetime.now().strftime("%Y-%m-%d"), ""]
        # Verifica se já existe (pelo nome)
        for i, row in enumerate(todos):
            if len(row) > 0 and _normalizar_chave_acesso(row[0]) == _normalizar_chave_acesso(nome_limpo):
                aba.update(f"A{i+1}:F{i+1}", [nova_linha])
                st.session_state.pop("_operadores_erro", None)
                return True
        # Insere novo
        aba.append_row(nova_linha, value_input_option="USER_ENTERED")
        st.session_state.pop("_operadores_erro", None)
        return True
    except Exception as e:
        _log_sheets_erro("sheets_salvar_operador", e)
        return False


def sheets_deletar_operador(nome: str) -> bool:
    """Remove operador da aba 👷 Operadores."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return False
        aba = sh.worksheet("👷 Operadores")
        todos = aba.get_all_values()
        for i, row in enumerate(todos):
            if len(row) > 0 and _normalizar_chave_acesso(row[0]) == _normalizar_chave_acesso(nome):
                aba.delete_rows(i + 1)
                return True
        return False
    except Exception as e:
        _log_sheets_erro("sheets_deletar_operador", e)
        return False


def verificar_pin_operador(pin_digitado: str) -> dict | None:
    """Verifica PIN e retorna dados do operador, ou None se inválido."""
    return validar_pin_operador(pin_digitado)

def _log_sheets_erro(contexto: str, erro: Exception):
    """Armazena o último erro do Google Sheets no session_state para diagnóstico."""
    import traceback
    msg = f"[{contexto}] {type(erro).__name__}: {erro}\n{traceback.format_exc()}"
    st.session_state["_sheets_ultimo_erro"] = msg


def conectar_sheets():
    """Conecta ao Google Sheets usando as credenciais do st.secrets ou arquivo local."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as e:
        _log_sheets_erro("conectar_sheets/import", e)
        st.session_state["_sheets_ultimo_erro"] = (
            "ERRO: gspread ou google-auth não está instalado no ambiente atual.\n"
            "Verifique o requirements.txt e force um redeploy no Streamlit Cloud."
        )
        return None

    try:
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        # Tenta carregar do st.secrets (Streamlit Cloud)
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            # Fallback: arquivo local (uso no computador)
            creds_path = Path(__file__).parent / "aqua-gestao-rt-87316ebf5331.json"
            if not creds_path.exists():
                st.session_state["_sheets_ultimo_erro"] = (
                    "ERRO: Nenhuma credencial encontrada.\n"
                    "No Streamlit Cloud: verifique st.secrets['gcp_service_account'].\n"
                    "Localmente: arquivo aqua-gestao-rt-87316ebf5331.json não encontrado."
                )
                return None
            creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)

        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SHEETS_ID)
        # Limpa erro anterior se conexão ok
        st.session_state.pop("_sheets_ultimo_erro", None)
        return sh
    except Exception as e:
        _log_sheets_erro("conectar_sheets", e)
        return None


def sheets_salvar_lancamento_campo(lancamento: dict, nome_condominio: str):
    """Salva lançamento de campo na aba Visitas do Google Sheets."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return False

        aba = sh.worksheet("🔬 Visitas")
        todos = aba.get_all_values()

        # Encontra próxima linha vazia após o cabeçalho (linha 6 = índice 5)
        proxima_linha = len(todos) + 1

        # Gera ID da visita
        visitas_existentes = [r for r in todos if r and r[1] and r[1].startswith("V")]
        proximo_num = len(visitas_existentes) + 1
        id_visita = f"V{proximo_num:05d}"

        # Busca ID do cliente
        aba_clientes = sh.worksheet("👥 Clientes")
        clientes = aba_clientes.get_all_values()
        id_cliente = ""
        for row in clientes:
            if len(row) > 2 and nome_condominio.lower() in str(row[2]).lower():
                id_cliente = row[1]
                break

        nova_linha = [
            "",                                    # col A  - vazia
            id_visita,                             # col B  - ID visita
            lancamento.get("data", ""),            # col C  - Data
            id_cliente,                            # col D  - ID cliente
            nome_condominio,                       # col E  - Condomínio
            lancamento.get("ph", ""),              # col F  - pH
            lancamento.get("cloro_livre", ""),     # col G  - CRL
            lancamento.get("cloro_total", ""),     # col H  - CT ← adicionado
            lancamento.get("alcalinidade", ""),    # col I  - Alcalinidade
            lancamento.get("dureza", ""),          # col J  - Dureza
            lancamento.get("cianurico", ""),       # col K  - CYA
            "",                                    # col L  - foto antes
            "",                                    # col M  - foto depois
            "",                                    # col N  - foto casa máquinas
            lancamento.get("observacao", ""),      # col O  - Observação
            "",                                    # col P  - dosagem cloro
            "",                                    # col Q  - dosagem bicarb
            "",                                    # col R  - alerta pH
            "",                                    # col S  - alerta cloro
            "Concluída",                           # col T  - Status
            lancamento.get("operador", ""),        # col U  - Operador
            lancamento.get("problemas", ""),       # col V  - Problemas
        ]

        aba.append_row(nova_linha, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        _log_sheets_erro("sheets_salvar_lancamento_campo", e)
        return False


def sheets_salvar_cliente(nome: str, cnpj: str, endereco: str, contato: str, telefone: str,
                           vol_adulto: float = 0, vol_infantil: float = 0, vol_family: float = 0,
                           empresa: str = "Aqua Gestão"):
    """Salva novo cliente na aba Clientes do Google Sheets.
    
    Insere sempre logo após o último cliente real (C001, C002...),
    mantendo a formatação da planilha com cabeçalho e linha de total intactos.
    Colunas J/K/L = Vol_Adulto_m3, Vol_Infantil_m3, Vol_Family_m3
    """
    try:
        import re as _re

        sh = conectar_sheets()
        if sh is None:
            return False

        aba = sh.worksheet("👥 Clientes")
        todos = aba.get_all_values()

        # Identifica linhas reais de clientes: col B começa com C + dígitos
        ultima_linha_cliente = 0  # índice 0-based
        nums = []
        nomes_existentes = []

        for i, row in enumerate(todos):
            if len(row) > 2:
                id_val = str(row[1]).strip()
                nome_val = str(row[2]).strip()
                if _re.match(r"^C[0-9]+$", id_val) and nome_val:
                    ultima_linha_cliente = i
                    nomes_existentes.append(nome_val.lower())
                    m = _re.match(r"^C([0-9]+)$", id_val)
                    if m:
                        nums.append(int(m.group(1)))

        # Verifica duplicata
        if nome.lower().strip() in nomes_existentes:
            return True  # Já existe, considera sucesso

        # Próximo ID baseado no maior existente
        proximo_num = (max(nums) + 1) if nums else 1
        id_cliente = f"C{proximo_num:03d}"

        vol_total = (vol_adulto or 0) + (vol_infantil or 0) + (vol_family or 0)
        nova_linha = [
            "",                                    # A - vazia
            id_cliente,                            # B - ID
            nome,                                  # C - Nome
            str(vol_total) if vol_total else "",   # D - Volume total m3
            formatar_telefone(telefone),           # E - Telefone
            contato,                               # F - Contato síndico
            endereco,                              # G - Endereço
            datetime.now().strftime("%Y-%m-%d"),   # H - Data cadastro
            "Ativo",                               # I - Status
            str(vol_adulto) if vol_adulto else "", # J - Vol Adulto m3
            str(vol_infantil) if vol_infantil else "", # K - Vol Infantil m3
            str(vol_family) if vol_family else "", # L - Vol Family m3
            empresa,                               # M - Empresa
        ]

        # Insere logo após o último cliente real (linha do Sheets = índice + 2)
        # Isso mantém o bloco de clientes agrupado antes do TOTAL
        linha_insercao = ultima_linha_cliente + 2  # +1 índice→sheets, +1 para inserir abaixo
        aba.insert_row(nova_linha, linha_insercao, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        _log_sheets_erro("sheets_salvar_cliente", e)
        return False


def sheets_listar_clientes() -> list[str]:
    """Retorna lista de nomes de clientes da aba Clientes."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return []
        aba = sh.worksheet("👥 Clientes")
        todos = aba.get_all_values()
        nomes = []
        for row in todos:
            if len(row) > 2 and str(row[1]).startswith("C") and row[2].strip():
                nomes.append(row[2].strip())
        return nomes
    except Exception:
        return []


def sheets_listar_clientes_completo() -> list[dict]:
    """Retorna lista de dicts com dados completos de cada cliente do Sheets.
    
    Mapeamento das colunas da planilha:
      B=ID, C=Nome, D=Volume_m3, E=Contato_Sindico/Telefone, 
      F=Email_Sindico, G=Endereco, H=Data_Cadastro, I=Status
    """
    import re as _re
    try:
        sh = conectar_sheets()
        if sh is None:
            return []
        aba = sh.worksheet("👥 Clientes")
        todos = aba.get_all_values()
        clientes = []
        for row in todos:
            if not row or len(row) < 3:
                continue
            id_val = str(row[1]).strip() if len(row) > 1 else ""
            if not _re.match(r"^C[0-9]+$", id_val):
                continue
            nome = str(row[2]).strip() if len(row) > 2 else ""
            if not nome:
                continue
            # Detecta se col E é telefone ou email
            col_e = str(row[4]).strip() if len(row) > 4 else ""
            col_f = str(row[5]).strip() if len(row) > 5 else ""
            # Se col_e tem @ é email do síndico; senão é telefone
            if "@" in col_e:
                telefone = ""
                contato  = ""
                email    = col_e
            else:
                telefone = formatar_telefone(col_e) if col_e else ""
                contato  = col_f if col_f else ""
                email    = ""
            # Volumes das piscinas (colunas J=9, K=10, L=11)
            def _vol(r, idx):
                try: return float(str(r[idx]).replace(",",".").strip() or 0) if len(r) > idx else 0.0
                except: return 0.0
            vol_adulto   = _vol(row, 9)
            vol_infantil = _vol(row, 10)
            vol_family   = _vol(row, 11)
            vol_total    = _vol(row, 3) or (vol_adulto + vol_infantil + vol_family)

            _empresa_cl = str(row[12]).strip() if len(row) > 12 else "Aqua Gestão"
            if not _empresa_cl:
                _empresa_cl = "Aqua Gestão"
            clientes.append({
                "id":           id_val,
                "nome":         nome,
                "cnpj":         "",
                "telefone":     telefone,
                "contato":      contato,
                "email":        email,
                "endereco":     str(row[6]).strip() if len(row) > 6 else "",
                "status":       str(row[8]).strip() if len(row) > 8 else "Ativo",
                "vol_total":    vol_total,
                "vol_adulto":   vol_adulto,
                "vol_infantil": vol_infantil,
                "vol_family":   vol_family,
                "empresa":      _empresa_cl,
                "piscinas_extras": [],  # carregado do JSON local se disponível
            })
        return clientes
    except Exception as e:
        _log_sheets_erro("sheets_listar_clientes_completo", e)
        return []



def sheets_editar_cliente(id_cliente: str, nome: str, cnpj: str, endereco: str,
                           contato: str, telefone: str,
                           vol_adulto: float = 0, vol_infantil: float = 0, vol_family: float = 0,
                           empresa: str = "") -> bool:
    """Edita cliente existente na aba Clientes pelo ID."""
    import re as _re
    try:
        sh = conectar_sheets()
        if sh is None:
            return False
        aba = sh.worksheet("👥 Clientes")
        todos = aba.get_all_values()
        vol_total = (vol_adulto or 0) + (vol_infantil or 0) + (vol_family or 0)
        for i, row in enumerate(todos):
            if len(row) > 1 and str(row[1]).strip() == id_cliente.strip():
                linha_sheets = i + 1
                # Preserva empresa existente se não informada
                _empresa_atual = str(row[12]).strip() if len(row) > 12 else ""
                _empresa_final = empresa if empresa else (_empresa_atual or "Aqua Gestão")
                nova = [
                    "",
                    id_cliente,
                    nome,
                    str(vol_total) if vol_total else "",
                    formatar_telefone(telefone),
                    contato,
                    endereco,
                    str(row[7]).strip() if len(row) > 7 else datetime.now().strftime("%Y-%m-%d"),
                    str(row[8]).strip() if len(row) > 8 else "Ativo",
                    str(vol_adulto) if vol_adulto else "",
                    str(vol_infantil) if vol_infantil else "",
                    str(vol_family) if vol_family else "",
                    _empresa_final,                # M - Empresa
                ]
                aba.update(f"A{linha_sheets}:M{linha_sheets}", [nova], value_input_option="USER_ENTERED")
                return True
        return False
    except Exception as e:
        _log_sheets_erro("sheets_editar_cliente", e)
        return False

def sheets_carregar_cliente_por_nome(nome: str) -> dict:
    """Retorna dict com dados do cliente pelo nome (busca parcial)."""
    clientes = sheets_listar_clientes_completo()
    nome_lower = nome.lower().strip()
    for c in clientes:
        if c["nome"].lower().strip() == nome_lower:
            return c
    # Busca parcial
    for c in clientes:
        if nome_lower in c["nome"].lower():
            return c
    return {}




# =========================================
# GESTÃO DE OPERADORES
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("👷 Gestão de Operadores")
st.caption("Tela administrativa segura para cadastrar, editar, ativar/inativar operadores e definir exatamente quais condomínios cada PIN pode acessar.")
st.info("🔐 PIN geral 2940 mantido como acesso mestre do sistema. Ele continua reservado e não pode ser usado no cadastro de operadores comuns.")

# Cria aba no Sheets se não existir
_col_op_top1, _col_op_top2 = st.columns([1, 1.35])
with _col_op_top1:
    if st.button("🔧 Inicializar aba de Operadores no Sheets", key="btn_init_op_aba"):
        ok_aba = sheets_criar_aba_operadores()
        if ok_aba:
            st.success("✅ Aba '👷 Operadores' criada/confirmada no Google Sheets!")
        else:
            st.error("❌ Erro ao criar aba. Verifique a conexão com o Sheets.")
with _col_op_top2:
    st.caption("Use este painel para centralizar o controle de PINs, status e permissões por condomínio, sem expor os PINs completos na tela.")

if st.session_state.get("_operadores_erro"):
    st.warning(st.session_state.get("_operadores_erro"))

# Lista operadores cadastrados
@st.cache_data(ttl=30)
def _listar_ops():
    return sheets_listar_operadores()

ops_cadastrados = _listar_ops()

# Cache de clientes para o painel de operadores
@st.cache_data(ttl=60)
def _clientes_para_painel_op():
    return sheets_listar_clientes_completo()

_todos_clientes_painel = _clientes_para_painel_op()
_nomes_todos_clientes = _condominios_organizar([c.get("nome", "") for c in _todos_clientes_painel if c.get("nome")])


def _mascarar_pin_admin(pin: str) -> str:
    _pin = str(pin or "").strip()
    if not _pin:
        return "—"
    if len(_pin) <= 2:
        return "*" * len(_pin)
    return f"{_pin[:2]}{'*' * (len(_pin) - 2)}"


def _op_tem_acesso_total(op: dict) -> bool:
    _conds = _condominios_organizar(op.get("condomínios", []))
    return op.get("acesso_total", False) or any(_normalizar_chave_acesso(c) == "todos" for c in _conds) or not _conds


def _resumo_acesso_admin(op: dict) -> str:
    _conds = _condominios_organizar(op.get("condomínios", []))
    if _op_tem_acesso_total(op):
        return f"Todos os {len(_nomes_todos_clientes)} condomínio(s) cadastrados"
    _exatos = _resolver_condominios_permitidos_exatos(_conds, _nomes_todos_clientes)
    _orfaos = [c for c in _conds if _normalizar_chave_acesso(c) not in {_normalizar_chave_acesso(x) for x in _exatos}]
    _extra = f" | {len(_orfaos)} não localizado(s) exatamente" if _orfaos else ""
    return f"{len(_exatos)} condomínio(s) vinculado(s){_extra}"


def _filtrar_ops_admin(lista: list[dict], busca: str, status: str) -> list[dict]:
    busca_norm = _normalizar_chave_acesso(busca)
    resultado = []
    for op in lista:
        if status == "Ativos" and not op.get("ativo"):
            continue
        if status == "Inativos" and op.get("ativo"):
            continue
        if busca_norm:
            alvo = " | ".join([
                str(op.get("nome", "")),
                str(op.get("pin", "")),
                " | ".join(_condominios_organizar(op.get("condomínios", []))),
            ])
            if busca_norm not in _normalizar_chave_acesso(alvo):
                continue
        resultado.append(op)
    return resultado


_total_ops = len(ops_cadastrados)
_total_ativos = sum(1 for op in ops_cadastrados if op.get("ativo"))
_total_inativos = _total_ops - _total_ativos
_total_restritos = sum(1 for op in ops_cadastrados if not _op_tem_acesso_total(op))

_mop1, _mop2, _mop3, _mop4 = st.columns(4)
with _mop1:
    st.metric("Operadores", _total_ops)
with _mop2:
    st.metric("Ativos", _total_ativos)
with _mop3:
    st.metric("Inativos", _total_inativos)
with _mop4:
    st.metric("Acesso restrito", _total_restritos)

if not ops_cadastrados:
    st.info("Nenhum operador cadastrado ainda. Use a aba 'Cadastrar novo operador'. O PIN 2940 continua funcionando como acesso geral.")

_tab_ops1, _tab_ops2 = st.tabs(["🛡️ Administrar operadores", "➕ Cadastrar novo operador"])

with _tab_ops1:
    if not ops_cadastrados:
        st.caption("Assim que houver operadores cadastrados, esta tela permitirá editar PIN, status e permissões com mais segurança.")
    else:
        _flt1, _flt2 = st.columns([1.1, 1.9])
        with _flt1:
            _status_ops = st.selectbox(
                "Filtrar status",
                ["Todos", "Ativos", "Inativos"],
                key="ops_admin_status",
            )
        with _flt2:
            _busca_ops = st.text_input(
                "Buscar operador, PIN mascarado ou condomínio",
                key="ops_admin_busca",
                placeholder="Ex.: João, Terra Nova, 12****",
            )

        _ops_visiveis = _filtrar_ops_admin(ops_cadastrados, _busca_ops, _status_ops)

        if not _ops_visiveis:
            st.warning("Nenhum operador encontrado com os filtros atuais.")
        else:
            _rotulos_ops = [
                f"{'🟢' if op.get('ativo') else '🔴'} {op.get('nome', 'Sem nome')} | PIN {_mascarar_pin_admin(op.get('pin', ''))} | {_resumo_acesso_admin(op)}"
                for op in _ops_visiveis
            ]
            _idx_op = st.selectbox(
                "Selecione o operador para administrar",
                options=list(range(len(_ops_visiveis))),
                format_func=lambda i: _rotulos_ops[i],
                key="ops_admin_selector",
            )
            _op_sel = _ops_visiveis[_idx_op]
            _op_nome_sel = _op_sel.get("nome", "")
            _op_pin_sel = str(_op_sel.get("pin", "")).strip()
            _op_conds_sel = _condominios_organizar(_op_sel.get("condomínios", []))
            _op_total_sel = _op_tem_acesso_total(_op_sel)
            _op_exatos_sel = _resolver_condominios_permitidos_exatos(_op_conds_sel, _nomes_todos_clientes)
            _op_exatos_set = {_normalizar_chave_acesso(c) for c in _op_exatos_sel}
            _op_orfaos_sel = [c for c in _op_conds_sel if _normalizar_chave_acesso(c) not in _op_exatos_set and _normalizar_chave_acesso(c) != "todos"]

            st.markdown("### Painel do operador selecionado")
            _sum1, _sum2, _sum3 = st.columns([1.1, 1, 1.4])
            with _sum1:
                st.caption("Operador")
                st.markdown(f"**{_op_nome_sel}**")
                st.caption(f"PIN mascarado: {_mascarar_pin_admin(_op_pin_sel)}")
            with _sum2:
                st.caption("Status")
                st.markdown("**🟢 Ativo**" if _op_sel.get("ativo") else "**🔴 Inativo**")
                st.caption("PIN geral 2940 não aparece aqui por segurança.")
            with _sum3:
                st.caption("Escopo de acesso")
                if _op_total_sel:
                    st.markdown(f"**Acesso total** aos {len(_nomes_todos_clientes)} condomínio(s) cadastrados")
                else:
                    st.markdown(f"**{len(_op_exatos_sel)} condomínio(s)** com correspondência exata")
                    if _op_orfaos_sel:
                        st.caption(f"{len(_op_orfaos_sel)} permissão(ões) salva(s) sem cliente exato no cadastro atual")

            with st.expander("🔎 Ver permissões atuais deste operador", expanded=False):
                if _op_total_sel:
                    st.success("Este operador está com acesso total liberado.")
                else:
                    if _op_exatos_sel:
                        st.markdown("**Condomínios localizados exatamente no cadastro atual:**")
                        for _c in _op_exatos_sel:
                            st.caption(f"✅ {_c}")
                    else:
                        st.caption("Nenhum condomínio localizado exatamente no cadastro atual.")
                    if _op_orfaos_sel:
                        st.markdown("**Permissões salvas que não batem exatamente com o cadastro atual:**")
                        for _c in _op_orfaos_sel:
                            st.caption(f"⚠️ {_c}")

            with st.form(f"form_admin_seguro_{_normalizar_chave_acesso(_op_nome_sel)}"):
                st.markdown("#### Editar operador com segurança")
                _ed1, _ed2 = st.columns(2)
                with _ed1:
                    st.text_input("Nome do operador", value=_op_nome_sel, disabled=True, help="Para preservar o comportamento atual do sistema, o nome é tratado como chave de atualização.")
                    _editar_pin = st.checkbox(
                        "Redefinir PIN deste operador",
                        value=False,
                        key=f"edit_pin_toggle_{_normalizar_chave_acesso(_op_nome_sel)}",
                        help="O PIN atual não é exibido em tela. Marque apenas se quiser trocar o PIN deste operador.",
                    )
                    _novo_pin_edit = st.text_input(
                        "Novo PIN",
                        value="",
                        type="password",
                        disabled=not _editar_pin,
                        key=f"novo_pin_{_normalizar_chave_acesso(_op_nome_sel)}",
                        placeholder="Digite um novo PIN exclusivo",
                    )
                with _ed2:
                    _ativo_edit = st.checkbox(
                        "Operador ativo",
                        value=bool(_op_sel.get("ativo", True)),
                        key=f"ativo_edit_{_normalizar_chave_acesso(_op_nome_sel)}",
                    )
                    _acesso_total_edit = st.checkbox(
                        "Acesso total a todos os condomínios",
                        value=_op_total_sel,
                        key=f"total_edit_{_normalizar_chave_acesso(_op_nome_sel)}",
                    )

                if not _acesso_total_edit:
                    _conds_edit = st.multiselect(
                        "Condomínios permitidos para este PIN",
                        options=_nomes_todos_clientes,
                        default=_op_exatos_sel,
                        key=f"conds_edit_{_normalizar_chave_acesso(_op_nome_sel)}",
                        help="Seleção exata dos condomínios liberados para este operador.",
                    )
                else:
                    st.caption("Com acesso total marcado, o operador continuará vendo todos os condomínios disponíveis.")
                    _conds_edit = ["TODOS"]

                _salvar_edit = st.form_submit_button("💾 Salvar alterações do operador", type="primary", use_container_width=True)

            if _salvar_edit:
                _pin_final_edit = _novo_pin_edit.strip() if _editar_pin else _op_pin_sel
                _conds_final_edit = ["TODOS"] if _acesso_total_edit else _condominios_organizar(_conds_edit)

                if not _pin_final_edit or len(_pin_final_edit) < 4:
                    st.error("PIN deve ter pelo menos 4 caracteres.")
                elif _pin_final_edit == "2940":
                    st.error("O PIN 2940 é reservado para acesso geral. Escolha outro para este operador.")
                elif _pin_operador_em_uso(_pin_final_edit, nome_ignorar=_op_nome_sel):
                    st.error(f"O PIN {_pin_final_edit} já está em uso por outro operador.")
                elif not _acesso_total_edit and not _conds_final_edit:
                    st.error("Selecione ao menos um condomínio para este operador ou marque acesso total.")
                else:
                    with st.spinner("Salvando alterações do operador..."):
                        ok_edit = sheets_salvar_operador(
                            nome=_op_nome_sel,
                            pin=_pin_final_edit,
                            condomínios=_conds_final_edit,
                            ativo=_ativo_edit,
                        )
                    if ok_edit:
                        st.session_state.pop("_operadores_erro", None)
                        st.success(f"✅ Operador '{_op_nome_sel}' atualizado com sucesso.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(st.session_state.get("_operadores_erro") or "Erro ao salvar alterações do operador. Verifique a conexão com o Sheets.")

            st.markdown("---")
            st.markdown("**Ações administrativas adicionais**")
            _rm_key = f"confirmar_remocao_{_normalizar_chave_acesso(_op_nome_sel)}"
            _rm1, _rm2, _rm3 = st.columns([1.2, 1.1, 1.1])
            with _rm1:
                if st.button("🗑 Solicitar remoção deste operador", key=f"btn_solicitar_rm_{_normalizar_chave_acesso(_op_nome_sel)}", use_container_width=True):
                    st.session_state[_rm_key] = True
                    st.rerun()
            if st.session_state.get(_rm_key):
                st.warning(f"Confirme a remoção do operador '{_op_nome_sel}'. Esta ação apaga o cadastro da aba 👷 Operadores.")
                with _rm2:
                    if st.button("✅ Confirmar remoção", key=f"btn_confirma_rm_{_normalizar_chave_acesso(_op_nome_sel)}", use_container_width=True):
                        if sheets_deletar_operador(_op_nome_sel):
                            st.session_state.pop(_rm_key, None)
                            st.session_state.pop("_operadores_erro", None)
                            st.success(f"Operador '{_op_nome_sel}' removido.")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Erro ao remover operador. Verifique a conexão com o Sheets.")
                with _rm3:
                    if st.button("Cancelar remoção", key=f"btn_cancela_rm_{_normalizar_chave_acesso(_op_nome_sel)}", use_container_width=True):
                        st.session_state.pop(_rm_key, None)
                        st.rerun()

with _tab_ops2:
    with st.form("form_novo_operador_seguro"):
        st.markdown("#### Cadastro seguro de novo operador")
        _novo1, _novo2 = st.columns(2)
        with _novo1:
            op_nome_novo = st.text_input("Nome do operador *", key="op_novo_nome", placeholder="Ex.: João Silva")
            op_pin_novo = st.text_input(
                "PIN exclusivo *",
                key="op_novo_pin",
                placeholder="Ex.: 1234",
                max_chars=10,
                type="password",
                help="Mínimo 4 caracteres. Não use 2940, pois este PIN é reservado para o acesso geral do sistema.",
            )
        with _novo2:
            op_ativo_novo = st.checkbox("Operador ativo", value=True, key="op_novo_ativo")
            op_acesso_total_novo = st.checkbox("Acesso a todos os condomínios", value=False, key="op_novo_acesso_total")

        if not op_acesso_total_novo:
            op_conds_novo = st.multiselect(
                "Condomínios permitidos para este novo PIN",
                options=_nomes_todos_clientes,
                key="op_novo_conds",
                help="Seleção exata dos condomínios liberados para o novo operador.",
            )
        else:
            st.caption("Com acesso total marcado, o novo operador verá todos os condomínios disponíveis no sistema.")
            op_conds_novo = ["TODOS"]

        _salvar_novo = st.form_submit_button("💾 Cadastrar operador", type="primary", use_container_width=True)

    if _salvar_novo:
        _nome_op_limpo = re.sub(r"\s+", " ", op_nome_novo.strip())
        _pin_op_limpo = op_pin_novo.strip()
        _conds_op_final = ["TODOS"] if op_acesso_total_novo else _condominios_organizar(op_conds_novo)

        if not _nome_op_limpo:
            st.error("Informe o nome do operador.")
        elif not _pin_op_limpo or len(_pin_op_limpo) < 4:
            st.error("PIN deve ter pelo menos 4 caracteres.")
        elif _pin_op_limpo == "2940":
            st.error("O PIN 2940 é reservado para acesso geral. Escolha outro.")
        elif _pin_operador_em_uso(_pin_op_limpo, nome_ignorar=_nome_op_limpo):
            st.error(f"O PIN {_pin_op_limpo} já está em uso por outro operador.")
        elif not op_acesso_total_novo and not _conds_op_final:
            st.error("Selecione ao menos um condomínio para este operador ou marque acesso total.")
        else:
            with st.spinner("Salvando operador..."):
                ok_op = sheets_salvar_operador(
                    nome=_nome_op_limpo,
                    pin=_pin_op_limpo,
                    condomínios=_conds_op_final,
                    ativo=op_ativo_novo,
                )
            if ok_op:
                st.session_state.pop("_operadores_erro", None)
                st.success(f"✅ Operador '{_nome_op_limpo}' cadastrado com sucesso.")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(st.session_state.get("_operadores_erro") or "❌ Erro ao salvar operador. Verifique a conexão com o Sheets.")

st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# CADASTRO DE CLIENTES — GOOGLE SHEETS
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("👥 Cadastro de Clientes")
st.caption("Clientes cadastrados aqui ficam disponíveis para o operador selecionar no celular.")

# ── Diagnóstico de conexão (visível para admin) ──────────────────────────────
with st.expander("🔌 Testar conexão com Google Sheets", expanded=False):
    if st.button("▶ Executar teste de conexão agora", key="btn_teste_sheets"):
        with st.spinner("Testando conexão..."):
            sh_teste = conectar_sheets()
        if sh_teste is not None:
            try:
                abas_disponiveis = [w.title for w in sh_teste.worksheets()]
                st.success("✅ Conexão estabelecida com sucesso!")
                st.write(f"**Abas encontradas na planilha:** {abas_disponiveis}")
                abas_esperadas = ["👥 Clientes", "🔬 Visitas"]
                faltando = [a for a in abas_esperadas if a not in abas_disponiveis]
                if faltando:
                    st.warning(
                        f"⚠️ Aba(s) esperada(s) não encontrada(s): {faltando}\n\n"
                        "Verifique se os nomes das abas na planilha estão escritos **exatamente** assim, "
                        "incluindo os emojis."
                    )
                else:
                    st.success("✅ Todas as abas necessárias foram encontradas.")
            except Exception as ex:
                st.warning(f"Conexão ok, mas erro ao listar abas: {ex}")
        else:
            st.error("❌ Falha na conexão com o Google Sheets.")
            erro_detalhado = st.session_state.get("_sheets_ultimo_erro", "Sem detalhes.")
            st.code(erro_detalhado, language="text")
            st.markdown(
                "**Próximos passos para resolver:**\n"
                "1. Confirme que `gspread` e `google-auth` estão no `requirements.txt`\n"
                "2. No Streamlit Cloud → Settings → Secrets → verifique se `[gcp_service_account]` está presente\n"
                "3. Force redeploy: no painel do Streamlit Cloud clique em **Reboot app**\n"
                "4. Verifique se a conta de serviço `aqua-gestao-sheets@aqua-gestao-rt.iam.gserviceaccount.com` tem acesso **Editor** à planilha"
            )
    st.caption("Use este botão sempre que a conexão com o Sheets falhar para identificar a causa exata.")

# Mostra clientes já cadastrados no Sheets
@st.cache_data(ttl=30)
def _clientes_cadastrados():
    return sheets_listar_clientes()

clientes_cadastrados = _clientes_cadastrados()

if clientes_cadastrados:
    st.success(f"✅ {len(clientes_cadastrados)} cliente(s) cadastrado(s) no Google Sheets:")
    for c in clientes_cadastrados:
        st.caption(f"• {c}")
else:
    st.info("Nenhum cliente cadastrado ainda. Use o formulário abaixo para adicionar.")

# Processa flag de limpeza ANTES de renderizar os widgets
if st.session_state.pop("_cc_limpar", False):
    for k in ["cc_nome","cc_cnpj","cc_cep","cc_endereco","cc_contato","cc_telefone",
              "cc_vol_adulto","cc_vol_infantil","cc_vol_family",
              "cc_pisc_extra1_nome","cc_pisc_extra1_vol",
              "cc_pisc_extra2_nome","cc_pisc_extra2_vol"]:
        st.session_state[k] = ""

# ── Seletor de edição ────────────────────────────────────────────────────────
_cc_modo = st.radio("Ação", ["➕ Novo cliente", "✏️ Editar cliente existente"],
    key="cc_modo_acao", horizontal=True, label_visibility="collapsed")

_cc_cliente_editar = {}
if _cc_modo == "✏️ Editar cliente existente":
    @st.cache_data(ttl=30)
    def _clientes_completos_edit():
        return sheets_listar_clientes_completo()
    _clientes_edit = _clientes_completos_edit()
    if _clientes_edit:
        _nomes_edit = [c["nome"] for c in _clientes_edit]
        _sel_edit = st.selectbox("Selecione o cliente para editar", _nomes_edit, key="cc_sel_editar")
        _cc_cliente_editar = next((c for c in _clientes_edit if c["nome"] == _sel_edit), {})
        if _cc_cliente_editar and st.button("📂 Carregar dados", key="btn_carregar_editar"):
            st.session_state["cc_nome"]         = _cc_cliente_editar.get("nome","")
            st.session_state["cc_cnpj"]         = _cc_cliente_editar.get("cnpj","")
            st.session_state["cc_cep"]          = _cc_cliente_editar.get("cep","")
            st.session_state["cc_endereco"]     = _cc_cliente_editar.get("endereco","")
            # Pré-seleciona empresa no radio
            _emp_carregada = _cc_cliente_editar.get("empresa","Aqua Gestão")
            _emp_map_inv = {"Aqua Gestão": "🔵 Aqua Gestão",
                            "Bem Star Piscinas": "⭐ Bem Star Piscinas",
                            "Ambas": "🔵⭐ Ambas"}
            st.session_state["cc_empresa"] = _emp_map_inv.get(_emp_carregada, "🔵 Aqua Gestão")
            st.session_state["cc_contato"]      = _cc_cliente_editar.get("contato","")
            st.session_state["cc_telefone"]     = _cc_cliente_editar.get("telefone","")
            st.session_state["cc_vol_adulto"]   = str(_cc_cliente_editar.get("vol_adulto","") or "")
            st.session_state["cc_vol_infantil"] = str(_cc_cliente_editar.get("vol_infantil","") or "")
            st.session_state["cc_vol_family"]   = str(_cc_cliente_editar.get("vol_family","") or "")
            _piscs_extras = _cc_cliente_editar.get("piscinas_extras", [])
            st.session_state["cc_pisc_extra1_nome"] = _piscs_extras[0]["nome"] if len(_piscs_extras) > 0 else ""
            st.session_state["cc_pisc_extra1_vol"]  = str(_piscs_extras[0].get("vol","") or "") if len(_piscs_extras) > 0 else ""
            st.session_state["cc_pisc_extra2_nome"] = _piscs_extras[1]["nome"] if len(_piscs_extras) > 1 else ""
            st.session_state["cc_pisc_extra2_vol"]  = str(_piscs_extras[1].get("vol","") or "") if len(_piscs_extras) > 1 else ""
            st.rerun()
    else:
        st.info("Nenhum cliente para editar.")

# ── Formulário ───────────────────────────────────────────────────────────────
def _mask_cc_cnpj():
    st.session_state["cc_cnpj"] = formatar_cnpj(st.session_state.get("cc_cnpj",""))

def _mask_cc_telefone():
    st.session_state["cc_telefone"] = formatar_telefone(st.session_state.get("cc_telefone",""))

# Seletor de empresa no cadastro
cc_empresa = st.radio(
    "Empresa vinculada",
    ["🔵 Aqua Gestão", "⭐ Bem Star Piscinas", "🔵⭐ Ambas"],
    key="cc_empresa",
    horizontal=True,
    help="Define para qual empresa este cliente pertence."
)
_cc_empresa_val = {"🔵 Aqua Gestão": "Aqua Gestão",
                   "⭐ Bem Star Piscinas": "Bem Star Piscinas",
                   "🔵⭐ Ambas": "Ambas"}.get(cc_empresa, "Aqua Gestão")

cc1, cc2 = st.columns(2)
with cc1:
    cc_nome     = st.text_input("Nome do condomínio / local *", key="cc_nome", placeholder="Ex.: Residencial Bella Vista")
    # CEP com busca automática ViaCEP
    # Aplica CEP formatado se acabou de buscar
    if st.session_state.get("_cc_cep_fmt"):
        st.session_state["cc_cep"] = st.session_state.pop("_cc_cep_fmt")
    _cep_col1, _cep_col2 = st.columns([3, 1])
    with _cep_col1:
        cc_cep = st.text_input("CEP", key="cc_cep", placeholder="00000-000",
            help="Digite o CEP e clique em 🔍 para preencher o endereço automaticamente")
    with _cep_col2:
        st.markdown("<br>", unsafe_allow_html=True)
        _btn_cep = st.button("🔍", key="btn_buscar_cep", help="Buscar CEP")
    if _btn_cep:
        _cep_valor = re.sub(r"\D", "", st.session_state.get("cc_cep", ""))
        if len(_cep_valor) == 8:
            with st.spinner("Buscando CEP..."):
                _dados_cep = buscar_cep(_cep_valor)
            if _dados_cep:
                _logradouro = _dados_cep.get("logradouro", "")
                _bairro     = _dados_cep.get("bairro", "")
                _cidade     = _dados_cep.get("localidade", "")
                _uf         = _dados_cep.get("uf", "")
                _cep_fmt    = f"{_cep_valor[:5]}-{_cep_valor[5:]}"
                _end_auto   = ", ".join(p for p in [_logradouro, _bairro, f"{_cidade}/{_uf}", _cep_fmt] if p)
                st.session_state["cc_endereco"] = _end_auto
                st.session_state["_cc_cep_fmt"] = _cep_fmt
                st.rerun()
            else:
                st.error("CEP não encontrado. Verifique e tente novamente.")
        else:
            st.warning("Digite um CEP válido com 8 dígitos.")
    cc_endereco = st.text_area("Endereço completo", key="cc_endereco", height=70, placeholder="Rua, número, bairro, cidade")
with cc2:
    cc_cnpj     = st.text_input("CNPJ (opcional)", key="cc_cnpj", placeholder="00.000.000/0000-00", on_change=_mask_cc_cnpj)
    cc_contato  = st.text_input("Síndico / responsável", key="cc_contato", placeholder="Nome do responsável")
    cc_telefone = st.text_input("Telefone (opcional)", key="cc_telefone", placeholder="(34) 99999-9999", on_change=_mask_cc_telefone)

# ── Volumes das piscinas ─────────────────────────────────────────────────────
st.markdown("**🏊 Volumes das piscinas (m³)**")
st.caption("Preencha apenas as piscinas que este local possui. O volume é usado para calcular dosagens automaticamente.")

cv1, cv2, cv3 = st.columns(3)
with cv1:
    cc_vol_adulto   = st.text_input("🏊 Adulto (m³)", key="cc_vol_adulto",
        placeholder="ex: 150", help="Volume da piscina adulto em metros cúbicos")
with cv2:
    cc_vol_infantil = st.text_input("🐣 Infantil (m³)", key="cc_vol_infantil",
        placeholder="ex: 30", help="Volume da piscina infantil em metros cúbicos")
with cv3:
    cc_vol_family   = st.text_input("👨‍👩‍👧 Family (m³)", key="cc_vol_family",
        placeholder="ex: 50", help="Volume da piscina family em metros cúbicos")

# Piscinas extras (outra, SPA, coberta, etc.)
st.markdown("**Outras piscinas** (SPA, coberta, olímpica, etc.)")
_cv_extra1, _cv_extra2 = st.columns(2)
with _cv_extra1:
    cc_pisc_extra1_nome = st.text_input("Nome da piscina extra 1", key="cc_pisc_extra1_nome",
        placeholder="ex: SPA, Coberta, Olímpica")
with _cv_extra2:
    cc_pisc_extra1_vol  = st.text_input("Volume (m³)", key="cc_pisc_extra1_vol",
        placeholder="ex: 80")
_cv_extra3, _cv_extra4 = st.columns(2)
with _cv_extra3:
    cc_pisc_extra2_nome = st.text_input("Nome da piscina extra 2", key="cc_pisc_extra2_nome",
        placeholder="ex: Aquecida, Semiolímpica")
with _cv_extra4:
    cc_pisc_extra2_vol  = st.text_input("Volume (m³) ", key="cc_pisc_extra2_vol",
        placeholder="ex: 120")

def _parse_vol(v):
    try: return float(str(v).replace(",",".").strip() or 0)
    except: return 0.0

# ── Botão salvar ─────────────────────────────────────────────────────────────
_btn_label = "💾 Salvar cliente no Google Sheets" if _cc_modo == "➕ Novo cliente" else "💾 Salvar alterações"
if st.button(_btn_label, type="primary", use_container_width=True):
    if not cc_nome.strip():
        st.error("Informe o nome do condomínio.")
    else:
        _vol_a = _parse_vol(cc_vol_adulto)
        _vol_i = _parse_vol(cc_vol_infantil)
        _vol_f = _parse_vol(cc_vol_family)
        with st.spinner("Salvando no Google Sheets..."):
            if _cc_modo == "✏️ Editar cliente existente" and _cc_cliente_editar.get("id"):
                # Coleta piscinas extras
                _piscs_extras_editar = []
                for _en, _ev in [
                    (st.session_state.get("cc_pisc_extra1_nome","").strip(),
                     st.session_state.get("cc_pisc_extra1_vol","").strip()),
                    (st.session_state.get("cc_pisc_extra2_nome","").strip(),
                     st.session_state.get("cc_pisc_extra2_vol","").strip()),
                ]:
                    if _en:
                        try: _ev_f = float(_ev.replace(",",".")) if _ev else 0
                        except: _ev_f = 0
                        _piscs_extras_editar.append({"nome": _en, "vol": _ev_f})
                ok = sheets_editar_cliente(
                    id_cliente=_cc_cliente_editar["id"],
                    nome=cc_nome.strip(), cnpj=cc_cnpj.strip(),
                    endereco=cc_endereco.strip(), contato=cc_contato.strip(),
                    telefone=cc_telefone.strip(),
                    vol_adulto=_vol_a, vol_infantil=_vol_i, vol_family=_vol_f,
                    empresa=_cc_empresa_val,
                )
                # Salva piscinas extras no JSON local
                if _piscs_extras_editar:
                    _pasta_extras2 = GENERATED_DIR / slugify_nome(cc_nome.strip())
                    _pasta_extras2.mkdir(parents=True, exist_ok=True)
                    _dados_extras2 = carregar_dados_condominio(_pasta_extras2) or {}
                    _dados_extras2["piscinas_extras"] = _piscs_extras_editar
                    _dados_extras2["nome_condominio"] = cc_nome.strip()
                    salvar_dados_condominio(_pasta_extras2, _dados_extras2)
                msg_ok = f"✅ Cliente '{cc_nome}' atualizado!"
            else:
                # Coleta piscinas extras
                _piscs_extras_salvar = []
                for _en, _ev in [
                    (st.session_state.get("cc_pisc_extra1_nome","").strip(),
                     st.session_state.get("cc_pisc_extra1_vol","").strip()),
                    (st.session_state.get("cc_pisc_extra2_nome","").strip(),
                     st.session_state.get("cc_pisc_extra2_vol","").strip()),
                ]:
                    if _en:
                        try: _ev_f = float(_ev.replace(",",".")) if _ev else 0
                        except: _ev_f = 0
                        _piscs_extras_salvar.append({"nome": _en, "vol": _ev_f})
                ok = sheets_salvar_cliente(
                    nome=cc_nome.strip(), cnpj=cc_cnpj.strip(),
                    endereco=cc_endereco.strip(), contato=cc_contato.strip(),
                    telefone=cc_telefone.strip(),
                    vol_adulto=_vol_a, vol_infantil=_vol_i, vol_family=_vol_f,
                    empresa=_cc_empresa_val,
                )
                # Salva piscinas extras no JSON local
                if _piscs_extras_salvar:
                    _pasta_extras = GENERATED_DIR / slugify_nome(cc_nome.strip())
                    _pasta_extras.mkdir(parents=True, exist_ok=True)
                    _dados_extras = carregar_dados_condominio(_pasta_extras) or {}
                    _dados_extras["piscinas_extras"] = _piscs_extras_salvar
                    _dados_extras["nome_condominio"] = cc_nome.strip()
                    salvar_dados_condominio(_pasta_extras, _dados_extras)
                msg_ok = f"✅ Cliente '{cc_nome}' salvo! O operador já pode selecioná-lo no celular."
        if ok:
            st.success(msg_ok)
            st.cache_data.clear()
            st.session_state["_cc_limpar"] = True
            st.rerun()
        else:
            st.error("❌ Não foi possível salvar no Google Sheets.")
            erro_detalhado = st.session_state.get("_sheets_ultimo_erro","")
            if erro_detalhado:
                with st.expander("🔍 Ver diagnóstico do erro", expanded=True):
                    st.code(erro_detalhado, language="text")

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# PAINEL DE VENCIMENTOS
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Painel de vencimentos")

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Total monitorado", len(painel_filtrado))
with m2:
    st.metric("Vencidos", len([i for i in painel_filtrado if i["status"]["codigo"] == "vencido"]))
with m3:
    st.metric("Vencem em breve", len([i for i in painel_filtrado if i["status"]["codigo"] == "vencendo"]))
with m4:
    st.metric("Vigentes", len([i for i in painel_filtrado if i["status"]["codigo"] == "vigente"]))
with m5:
    st.metric("Sem vigência", len([i for i in painel_filtrado if i["status"]["codigo"] == "indefinido"]))

st.caption(
    "Painel operacional separado do histórico comum, com filtro central, exportação de cadastro e visualização dos últimos documentos."
)

aba1, aba2, aba3 = st.tabs(
    ["Condomínios vencidos", "Condomínios que vencem em breve", "Sem vigência válida"]
)

def render_exportacao_e_docs(item: dict, item_key: str):
    dados = item["dados"]
    ultimo_contrato = item["ultimo_contrato"]
    ultimo_aditivo = item["ultimo_aditivo"]
    ultimo_relatorio = item.get("ultimo_relatorio")

    st.markdown("<div class='docs-note'><strong>Documentos mais recentes</strong></div>", unsafe_allow_html=True)

    dc1, dc2, dc3, dc4, dc5 = st.columns(5)
    with dc1:
        if ultimo_contrato:
            if st.button("Abrir último contrato", key=f"abrir_contrato_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_contrato["path"])
        else:
            st.button("Sem contrato", key=f"sem_contrato_{item_key}", disabled=True, use_container_width=True)

    with dc2:
        if ultimo_aditivo:
            if st.button("Abrir último aditivo", key=f"abrir_aditivo_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_aditivo["path"])
        else:
            st.button("Sem aditivo", key=f"sem_aditivo_{item_key}", disabled=True, use_container_width=True)

    with dc3:
        if ultimo_relatorio:
            if st.button("Abrir último relatório", key=f"abrir_relatorio_{item_key}", use_container_width=True):
                abrir_arquivo_windows(ultimo_relatorio["path"])
        else:
            st.button("Sem relatório", key=f"sem_relatorio_{item_key}", disabled=True, use_container_width=True)

    with dc4:
        if dados:
            json_bytes = json.dumps(dados, ensure_ascii=False, indent=2).encode("utf-8")
            st.download_button(
                "Exportar JSON backup",
                data=json_bytes,
                file_name=f"{item['slug']}_cadastro.json",
                mime="application/json",
                key=f"download_json_{item_key}",
                use_container_width=True,
            )
        else:
            st.button("Sem JSON", key=f"sem_json_{item_key}", disabled=True, use_container_width=True)

    with dc5:
        if dados:
            html = gerar_html_resumo_cadastro(item).encode("utf-8")
            st.download_button(
                "Exportar resumo HTML",
                data=html,
                file_name=f"{item['slug']}_resumo_cadastro.html",
                mime="text/html",
                key=f"download_html_{item_key}",
                use_container_width=True,
            )
        else:
            st.button("Sem resumo", key=f"sem_resumo_{item_key}", disabled=True, use_container_width=True)


with aba1:
    if not itens_vencidos:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio vencido no momento.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_vencidos:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            data_fim = item["data_fim"] or "Não informada"
            item_key = chave_segura(f"painel_vencido_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Data final:</strong> {data_fim}</div>
                    <div class="venc-meta"><strong>Situação:</strong> {texto_dias_restantes(status)}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("Editar cadastro", key=f"editar_vencido_{item_key}", use_container_width=True):
                    aplicar_dados_no_formulario(dados_salvos)
                    st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                    st.rerun()

            with c2:
                if st.button("Abrir pasta", key=f"abrir_vencido_{item_key}", use_container_width=True):
                    abrir_pasta_windows(pasta)

            with c3:
                if st.button("Renovar no formulário", key=f"renovar_vencido_{item_key}", use_container_width=True):
                    ok, msg = preparar_renovacao_no_formulario(dados_salvos)
                    if ok:
                        st.session_state.painel_acao_msg = f"{msg} Condomínio: {nome_exibicao}."
                        st.rerun()
                    else:
                        st.error(msg)

            with c4:
                if st.button("Gerar aditivo renovação", key=f"aditivo_vencido_{item_key}", use_container_width=True):
                    ok, msg = gerar_aditivo_renovacao_por_painel(pasta, st.session_state.alerta_vencimento_dias)
                    if ok:
                        st.session_state.painel_acao_msg = msg
                        st.rerun()
                    else:
                        st.error(msg)

            render_exportacao_e_docs(item, item_key)

with aba2:
    if not itens_vencendo:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio dentro da faixa de alerta no momento.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_vencendo:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            data_fim = item["data_fim"] or "Não informada"
            item_key = chave_segura(f"painel_vencendo_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Data final:</strong> {data_fim}</div>
                    <div class="venc-meta"><strong>Situação:</strong> {texto_dias_restantes(status)}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("Editar cadastro", key=f"editar_vencendo_{item_key}", use_container_width=True):
                    aplicar_dados_no_formulario(dados_salvos)
                    st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                    st.rerun()

            with c2:
                if st.button("Abrir pasta", key=f"abrir_vencendo_{item_key}", use_container_width=True):
                    abrir_pasta_windows(pasta)

            with c3:
                if st.button("Renovar no formulário", key=f"renovar_vencendo_{item_key}", use_container_width=True):
                    ok, msg = preparar_renovacao_no_formulario(dados_salvos)
                    if ok:
                        st.session_state.painel_acao_msg = f"{msg} Condomínio: {nome_exibicao}."
                        st.rerun()
                    else:
                        st.error(msg)

            with c4:
                if st.button("Gerar aditivo renovação", key=f"aditivo_vencendo_{item_key}", use_container_width=True):
                    ok, msg = gerar_aditivo_renovacao_por_painel(pasta, st.session_state.alerta_vencimento_dias)
                    if ok:
                        st.session_state.painel_acao_msg = msg
                        st.rerun()
                    else:
                        st.error(msg)

            render_exportacao_e_docs(item, item_key)

with aba3:
    if not itens_indefinidos:
        st.markdown(
            "<div class='venc-empty'>Nenhum condomínio sem vigência válida encontrado.</div>",
            unsafe_allow_html=True,
        )
    else:
        for item in itens_indefinidos:
            nome_exibicao = item["nome_exibicao"]
            pasta = item["pasta"]
            dados_salvos = item["dados"]
            status = item["status"]
            origem = item["origem"]
            total_arquivos = len(item["arquivos"])
            item_key = chave_segura(f"painel_indefinido_{pasta}")

            st.markdown(
                f"""
                <div class="venc-row">
                    <div class="venc-nome">{nome_exibicao}</div>
                    <div class="venc-meta"><strong>Status:</strong> {status['rotulo']}</div>
                    <div class="venc-meta"><strong>Arquivos encontrados:</strong> {total_arquivos}</div>
                    <span class="status-badge {status['css']}">{status['rotulo']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if origem == "legado_sem_json":
                st.markdown(
                    "<div class='legacy-note'>Histórico antigo sem <strong>dados_condominio.json</strong>. "
                    "Agora você pode criar um cadastro inicial diretamente a partir desta pasta.</div>",
                    unsafe_allow_html=True,
                )

                c1, c2, c3 = st.columns(3)
                with c1:
                    if st.button("Criar cadastro desta pasta", key=f"legado_{item_key}", use_container_width=True):
                        preparar_cadastro_legado(pasta.name)
                        st.session_state.painel_acao_msg = f"Cadastro inicial preparado a partir da pasta '{pasta.name}'."
                        st.rerun()

                with c2:
                    if st.button("Abrir pasta", key=f"abrir_legado_{item_key}", use_container_width=True):
                        abrir_pasta_windows(pasta)

                with c3:
                    st.button("Gerar aditivo renovação", key=f"aditivo_legado_{item_key}", disabled=True, use_container_width=True)
            else:
                st.markdown(
                    "<div class='legacy-note'>Existe cadastro salvo, porém sem data final válida. "
                    "Carregue os dados no formulário e corrija a vigência.</div>",
                    unsafe_allow_html=True,
                )

                c1, c2, c3 = st.columns(3)
                with c1:
                    if st.button("Editar cadastro", key=f"editar_indefinido_{item_key}", use_container_width=True):
                        aplicar_dados_no_formulario(dados_salvos)
                        st.session_state.painel_acao_msg = f"Cadastro de '{nome_exibicao}' carregado para edição."
                        st.rerun()
                with c2:
                    if st.button("Abrir pasta", key=f"abrir_indefinido_{item_key}", use_container_width=True):
                        abrir_pasta_windows(pasta)
                with c3:
                    st.button("Gerar aditivo renovação", key=f"aditivo_indefinido_{item_key}", disabled=True, use_container_width=True)

            render_exportacao_e_docs(item, item_key)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# CLIENTES SEM RT — CADASTRO E RELATÓRIO TÉCNICO
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Clientes sem RT — Relatório Técnico Simples")
st.caption("Para condomínios que não possuem contrato de RT mas recebem visita técnica com análise e dosagem.")

with st.expander("📋 Cadastrar / selecionar cliente sem RT", expanded=False):

    # Lista clientes sem RT já cadastrados
    CLIENTES_SEM_RT_JSON = GENERATED_DIR / "_clientes_sem_rt.json"

    def carregar_clientes_sem_rt() -> list:
        if CLIENTES_SEM_RT_JSON.exists():
            try:
                return json.loads(CLIENTES_SEM_RT_JSON.read_text(encoding="utf-8"))
            except Exception:
                return []
        return []

    def salvar_clientes_sem_rt(lista: list):
        GENERATED_DIR.mkdir(exist_ok=True)
        CLIENTES_SEM_RT_JSON.write_text(json.dumps(lista, ensure_ascii=False, indent=2), encoding="utf-8")

    clientes_sem_rt = carregar_clientes_sem_rt()

    # ── Importar do Sheets ────────────────────────────────────────────────────
    @st.cache_data(ttl=60)
    def _clientes_sheets_csr():
        return sheets_listar_clientes_completo()

    _cls_sheets = _clientes_sheets_csr()
    if _cls_sheets:
        _opcoes_csr_imp = ["— Importar do cadastro principal (Sheets) —"] + [c["nome"] for c in _cls_sheets]
        _sel_csr_imp = st.selectbox(
            "🔗 Importar cliente do Google Sheets",
            _opcoes_csr_imp,
            key="sel_importar_csr",
            help="Importa os dados do cadastro principal para o formulário abaixo."
        )
        if _sel_csr_imp and _sel_csr_imp != "— Importar do cadastro principal (Sheets) —":
            if st.button("⬇️ Importar dados", key="btn_imp_csr", use_container_width=False):
                _d = next((c for c in _cls_sheets if c["nome"] == _sel_csr_imp), {})
                if _d:
                    st.session_state["csr_nome"]     = _d.get("nome", "")
                    st.session_state["csr_cnpj"]     = formatar_cnpj(_d.get("cnpj", ""))
                    st.session_state["csr_endereco"] = _d.get("endereco", "")
                    st.session_state["csr_contato"]  = _d.get("contato", "")
                    st.session_state["csr_telefone"] = formatar_telefone(_d.get("telefone", ""))
                    st.success(f"✅ Dados de '{_sel_csr_imp}' importados!")
                    st.rerun()
        st.markdown("---")

    st.markdown("**Novo cliente sem RT:**")

    def _mask_csr_cnpj():
        st.session_state["csr_cnpj"] = formatar_cnpj(st.session_state.get("csr_cnpj", ""))

    def _mask_csr_telefone():
        st.session_state["csr_telefone"] = formatar_telefone(st.session_state.get("csr_telefone", ""))

    csr1, csr2 = st.columns(2)
    with csr1:
        csr_nome = st.text_input("Nome do local / condomínio", key="csr_nome", placeholder="Ex.: Residencial Sol Nascente")
        if st.session_state.get("_csr_cep_fmt"):
            st.session_state["csr_cep"] = st.session_state.pop("_csr_cep_fmt")
        _csr_cep_c1, _csr_cep_c2 = st.columns([3, 1])
        with _csr_cep_c1:
            st.text_input("CEP", key="csr_cep", placeholder="00000-000",
                help="Digite o CEP e clique em 🔍 para preencher o endereço automaticamente")
        with _csr_cep_c2:
            st.markdown("<br>", unsafe_allow_html=True)
            _btn_csr_cep = st.button("🔍", key="btn_buscar_cep_csr", help="Buscar CEP")
        if _btn_csr_cep:
            _cep_v = re.sub(r"\D", "", st.session_state.get("csr_cep", ""))
            if len(_cep_v) == 8:
                with st.spinner("Buscando CEP..."):
                    _dc = buscar_cep(_cep_v)
                if _dc:
                    _end = ", ".join(p for p in [_dc.get("logradouro",""), _dc.get("bairro",""), f"{_dc.get('localidade','')}/{_dc.get('uf','')}", f"{_cep_v[:5]}-{_cep_v[5:]}"] if p)
                    st.session_state["csr_endereco"] = _end
                    st.session_state["_csr_cep_fmt"] = f"{_cep_v[:5]}-{_cep_v[5:]}"
                    st.rerun()
                else:
                    st.error("CEP não encontrado.")
            else:
                st.warning("Digite um CEP válido com 8 dígitos.")
        csr_endereco = st.text_area("Endereço", key="csr_endereco", height=70, placeholder="Rua, número, bairro, cidade")
    with csr2:
        csr_cnpj = st.text_input("CNPJ (opcional)", key="csr_cnpj", placeholder="00.000.000/0000-00", on_change=_mask_csr_cnpj)
        csr_contato = st.text_input("Responsável / contato", key="csr_contato", placeholder="Nome do responsável")
        csr_telefone = st.text_input("Telefone (opcional)", key="csr_telefone", placeholder="(34) 99999-9999", on_change=_mask_csr_telefone)

    if st.button("➕ Salvar cliente (Bem Star)", use_container_width=True):
        if not csr_nome.strip():
            st.error("Informe o nome do local.")
        else:
            novo = {
                "nome": csr_nome.strip(),
                "cnpj": formatar_cnpj(csr_cnpj.strip()),
                "endereco": csr_endereco.strip(),
                "contato": csr_contato.strip(),
                "telefone": formatar_telefone(csr_telefone.strip()),
                "cadastrado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }
            # Atualiza se já existe, senão adiciona
            nomes_existentes = [c["nome"].lower() for c in clientes_sem_rt]
            if csr_nome.strip().lower() in nomes_existentes:
                idx_ex = nomes_existentes.index(csr_nome.strip().lower())
                clientes_sem_rt[idx_ex] = novo
                st.success(f"Cliente '{csr_nome}' atualizado.")
            else:
                clientes_sem_rt.append(novo)
                st.success(f"Cliente '{csr_nome}' cadastrado com sucesso.")
            salvar_clientes_sem_rt(clientes_sem_rt)
            # Salva também no Google Sheets (col M = Bem Star Piscinas)
            with st.spinner("Sincronizando com Google Sheets..."):
                _cl_sheets = sheets_listar_clientes_completo()
                _existe_sheets = next((c for c in _cl_sheets
                    if c["nome"].lower().strip() == csr_nome.strip().lower()), None)
                if _existe_sheets:
                    sheets_editar_cliente(
                        id_cliente=_existe_sheets["id"],
                        nome=csr_nome.strip(),
                        cnpj=formatar_cnpj(csr_cnpj.strip()),
                        endereco=csr_endereco.strip(),
                        contato=csr_contato.strip(),
                        telefone=formatar_telefone(csr_telefone.strip()),
                        empresa="Bem Star Piscinas",
                    )
                else:
                    sheets_salvar_cliente(
                        nome=csr_nome.strip(),
                        cnpj=formatar_cnpj(csr_cnpj.strip()),
                        endereco=csr_endereco.strip(),
                        contato=csr_contato.strip(),
                        telefone=formatar_telefone(csr_telefone.strip()),
                        empresa="Bem Star Piscinas",
                    )
            st.cache_data.clear()
            # Cria pasta do cliente no generated
            pasta_csr = GENERATED_DIR / slugify_nome(csr_nome.strip())
            pasta_csr.mkdir(parents=True, exist_ok=True)
            salvar_dados_condominio(pasta_csr, {
                "nome_condominio": csr_nome.strip(),
                "cnpj_condominio": csr_cnpj.strip(),
                "endereco_condominio": csr_endereco.strip(),
                "nome_sindico": csr_contato.strip(),
                "tipo": "sem_rt",
                "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            })
            st.rerun()

    if clientes_sem_rt:
        st.markdown(f"**{len(clientes_sem_rt)} cliente(s) cadastrado(s) sem RT:**")
        for c in clientes_sem_rt:
            st.caption(f"📍 {c['nome']} | {c.get('contato','–')} | {c.get('endereco','–')[:50]}")

# ---- GERAÇÃO DO RELATÓRIO TÉCNICO SIMPLES ----
st.markdown("---")
st.markdown("**📊 Relatório técnico Bem Star Piscinas (sem RT)**")

# Carrega clientes Bem Star do Sheets (fonte principal) + fallback JSON local
@st.cache_data(ttl=30)
def _clientes_bem_star_relatorio():
    _todos = sheets_listar_clientes_completo()
    _bs = filtrar_clientes_por_empresa(_todos, "bem_star")
    # Fallback: também inclui clientes do JSON local
    _json_local = carregar_clientes_sem_rt() if CLIENTES_SEM_RT_JSON.exists() else []
    _nomes_sheets = {c["nome"].lower() for c in _bs}
    for _cl in _json_local:
        if _cl["nome"].lower() not in _nomes_sheets:
            _bs.append(_cl)
    return _bs

clientes_sem_rt_reload = _clientes_bem_star_relatorio()
opcoes_csr = [c["nome"] for c in clientes_sem_rt_reload]

if not opcoes_csr:
    st.info("Cadastre um cliente Bem Star acima para gerar o relatório técnico.")
else:
    rts1, rts2, rts3 = st.columns([2, 1, 1])
    with rts1:
        csr_sel = st.selectbox("Selecione o cliente", opcoes_csr, key="csr_sel_relatorio")
    with rts2:
        csr_mes = st.text_input("Mês", key="csr_mes_rel", placeholder=datetime.now().strftime("%m"))
    with rts3:
        csr_ano = st.text_input("Ano", key="csr_ano_rel", placeholder=str(datetime.now().year))

    csr_dados_sel = next((c for c in clientes_sem_rt_reload if c["nome"] == csr_sel), {})

    # ── Busca lançamentos: JSON local + Google Sheets ──────────────────────
    pasta_csr_sel = GENERATED_DIR / slugify_nome(csr_sel) if csr_sel else None
    _lanc_local_csr = []
    if pasta_csr_sel and pasta_csr_sel.exists():
        _dados_json_csr = carregar_dados_condominio(pasta_csr_sel)
        _lanc_local_csr = (_dados_json_csr or {}).get("lancamentos_campo", [])

    _lanc_sheets_csr = []
    if csr_sel:
        try:
            _lanc_sheets_csr = sheets_listar_lancamentos(csr_sel)
        except Exception:
            _lanc_sheets_csr = []

    # Deduplica local + Sheets
    _vistos_csr = set()
    _lanc_todos_csr = []
    for _lc in _lanc_local_csr + _lanc_sheets_csr:
        _ch = f"{_lc.get('data','')}-{_lc.get('operador','')}-{_lc.get('ph','')}"
        if _ch not in _vistos_csr:
            _vistos_csr.add(_ch)
            _lanc_todos_csr.append(_lc)

    # Filtra por mês/ano
    def _filtrar_mes_csr(lancamentos, mes, ano):
        if not mes or not ano:
            return lancamentos
        filtrados = []
        for lc in lancamentos:
            data = lc.get("data", "")
            try:
                if "/" in data:
                    p = data.split("/")
                    if len(p) == 3 and p[1] == mes.zfill(2) and p[2] == ano:
                        filtrados.append(lc)
                elif "-" in data:
                    p = data.split("-")
                    if len(p) == 3 and p[1] == mes.zfill(2) and p[0] == ano:
                        filtrados.append(lc)
            except Exception:
                filtrados.append(lc)
        return filtrados

    _mes_csr  = (csr_mes or "").strip()
    _ano_csr  = (csr_ano or str(datetime.now().year)).strip()
    lancamentos_csr = _filtrar_mes_csr(_lanc_todos_csr, _mes_csr, _ano_csr)

    # ── Painel de lançamentos disponíveis ────────────────────────────────
    if lancamentos_csr:
        _fonte_csr = "📱 local + Sheets" if _lanc_sheets_csr else "📱 local"
        _periodo_csr = f"{lancamentos_csr[0].get('data','?')} → {lancamentos_csr[-1].get('data','?')}"
        st.markdown(
            f"<div style='border:1px solid rgba(20,120,60,0.3);border-radius:12px;padding:12px 16px;"
            f"background:rgba(20,120,60,0.07);margin-bottom:12px;'>"
            f"<strong>📱 {len(lancamentos_csr)} lançamento(s) encontrado(s) — {_fonte_csr}</strong><br>"
            f"<span style='font-size:0.85rem;color:#3a6a3a;'>Período: {_periodo_csr}</span></div>",
            unsafe_allow_html=True,
        )
        with st.expander("👁 Ver lançamentos"):
            for _lc in lancamentos_csr:
                st.caption(f"📅 {_lc.get('data','?')} | pH {_lc.get('ph','–')} | CRL {_lc.get('cloro_livre','–')} | op: {_lc.get('operador','–')}")
    else:
        st.info("Nenhum lançamento encontrado para este cliente/período. O operador precisa registrar visitas no modo campo.")

    csr_operador_nome = st.text_input("Operador responsável", key="csr_operador_rel", placeholder="Nome do operador")
    csr_obs_geral     = st.text_area("Observações gerais", key="csr_obs_rel", height=80,
        placeholder="Condições gerais da piscina, ocorrências, recomendações...")

    # ── Coleta fotos das visitas ─────────────────────────────────────────
    pasta_fotos_csr = (GENERATED_DIR / slugify_nome(csr_sel) / "fotos_campo") if csr_sel else None
    fotos_csr = []
    if pasta_fotos_csr and pasta_fotos_csr.exists():
        for _lc in lancamentos_csr:
            for _nf in _lc.get("fotos", []):
                _pf = pasta_fotos_csr / _nf
                if _pf.exists():
                    fotos_csr.append((_lc.get("data",""), _pf))

    if fotos_csr:
        st.caption(f"📷 {len(fotos_csr)} foto(s) serão incluídas no relatório.")

    if st.button("📄 Gerar relatório Bem Star (PDF)", type="primary", use_container_width=True):
        if not csr_sel or not _mes_csr or not _ano_csr:
            st.error("Selecione o cliente, mês e ano.")
        else:
            try:
                # Monta lista de lançamentos — deduplica por data+operador+ph
                _lanc_vistos = set()
                _lanc_para_relatorio = []
                for _lc in lancamentos_csr:
                    _chave_lc = f"{_lc.get('data','')}-{_lc.get('operador','')}-{_lc.get('ph','') or (_lc.get('piscinas',[{}])[0].get('ph','') if _lc.get('piscinas') else '')}"
                    if _chave_lc in _lanc_vistos:
                        continue
                    _lanc_vistos.add(_chave_lc)
                    piscinas = _lc.get("piscinas", [])
                    if piscinas:
                        _dados = piscinas[0]
                    else:
                        _dados = _lc
                    _lanc_para_relatorio.append({
                        "data":         _lc.get("data", ""),
                        "ph":           _dados.get("ph", _lc.get("ph", "")),
                        "cloro_livre":  _dados.get("cloro_livre", _lc.get("cloro_livre", "")),
                        "cloro_total":  _dados.get("cloro_total", _lc.get("cloro_total", "")),
                        "alcalinidade": _dados.get("alcalinidade", _lc.get("alcalinidade", "")),
                        "dureza":       _dados.get("dureza", _lc.get("dureza", "")),
                        "cianurico":    _dados.get("cianurico", _lc.get("cianurico", "")),
                        "operador":     _lc.get("operador", csr_operador_nome),
                        "observacao":   _lc.get("observacao", ""),
                        "problemas":    _lc.get("problemas", ""),
                        "dosagens":     _dados.get("dosagens", _lc.get("dosagens", [])),
                    })

                # obs_geral passa apenas obs gerais do campo de texto.
                # Problemas e observações dos lançamentos são extraídos
                # diretamente pela função gerar_relatorio_visita_docx via
                # os campos "problemas" e "observacao" de cada lançamento.
                _obs_final = csr_obs_geral.strip()

                pasta_csr_out = GENERATED_DIR / slugify_nome(csr_sel)
                pasta_csr_out.mkdir(parents=True, exist_ok=True)
                _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                docx_csr = pasta_csr_out / f"{_ts}_{slugify_nome(csr_sel)}_RELATORIO_BS.docx"
                pdf_csr  = pasta_csr_out / f"{_ts}_{slugify_nome(csr_sel)}_RELATORIO_BS.pdf"

                with st.spinner("Gerando relatório Bem Star..."):
                    # fotos_csr é list[(data_str, Path)] — extrai só os Paths
                    # Deduplica por nome base de arquivo (remove duplicatas entre lançamentos)
                    _fotos_vistas = set()
                    _fotos_paths = []
                    for _, _fp in (fotos_csr or []):
                        # Usa nome sem timestamp inicial para detectar duplicatas
                        _nome_base = "_".join(_fp.name.split("_")[2:]) if _fp.name.count("_") >= 2 else _fp.name
                        if _nome_base not in _fotos_vistas:
                            _fotos_vistas.add(_nome_base)
                            _fotos_paths.append(_fp)
                    _ok_docx, _err_docx = gerar_relatorio_visita_docx(
                        output_path   = docx_csr,
                        nome_local    = csr_dados_sel.get("nome", csr_sel),
                        cnpj          = csr_dados_sel.get("cnpj", ""),
                        endereco      = csr_dados_sel.get("endereco", ""),
                        responsavel   = csr_dados_sel.get("contato", ""),
                        operador      = csr_operador_nome,
                        mes           = _mes_csr,
                        ano           = _ano_csr,
                        lancamentos   = _lanc_para_relatorio,
                        obs_geral     = _obs_final,
                        incluir_rt    = False,
                        fotos         = _fotos_paths,
                    )

                if not _ok_docx:
                    st.error(f"Erro ao gerar DOCX: {_err_docx}")
                else:
                    ok_pdf_csr, err_pdf_csr = converter_docx_para_pdf(docx_csr, pdf_csr)
                    registrar_documento_manifest(pasta_csr_out, csr_sel, "Relatório", docx_csr, pdf_csr, ok_pdf_csr, err_pdf_csr)
                    st.success(f"✅ Relatório Bem Star gerado! {len(fotos_csr)} foto(s) incluída(s).")
                    dl1, dl2 = st.columns(2)
                    with dl1:
                        with open(docx_csr, "rb") as _f:
                            st.download_button("⬇️ Baixar DOCX", data=_f, file_name=docx_csr.name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True)
                    with dl2:
                        if ok_pdf_csr and pdf_csr.exists():
                            with open(pdf_csr, "rb") as _f:
                                st.download_button("⬇️ Baixar PDF", data=_f, file_name=pdf_csr.name,
                                    mime="application/pdf", use_container_width=True)
                        else:
                            st.warning(f"PDF não gerado: {err_pdf_csr}")

                    # ── Bloco de envio ────────────────────────────────────
                    _msg_rel = montar_mensagem_bem_star(
                        nome_local  = csr_dados_sel.get("nome", csr_sel),
                        responsavel = csr_dados_sel.get("contato", ""),
                        tipo        = "relatorio",
                        mes         = _mes_csr,
                        ano         = _ano_csr,
                    )
                    exibir_bloco_envio_bem_star(
                        nome_local  = csr_dados_sel.get("nome", csr_sel),
                        pasta       = pasta_csr_out,
                        telefone    = csr_dados_sel.get("telefone", ""),
                        email       = csr_dados_sel.get("email", ""),
                        mensagem    = _msg_rel,
                        key_suffix  = "relatorio",
                    )

            except Exception as e:
                st.error(f"Erro ao gerar relatório Bem Star: {e}")
                import traceback
                st.code(traceback.format_exc(), language="text")


st.markdown("</div>", unsafe_allow_html=True)


# =========================================
# CONTRATO BEM STAR PISCINAS
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("📝 Contrato Bem Star Piscinas")
st.caption("Gera o contrato de prestação de serviços de limpeza e manutenção de piscinas em PDF.")

with st.expander("📋 Preencher e gerar contrato", expanded=False):

    # ── Seletor de cliente ────────────────────────────────────────────────────
    @st.cache_data(ttl=30)
    def _clientes_bs_contrato():
        _todos = sheets_listar_clientes_completo()
        _locais = carregar_clientes_sem_rt() if CLIENTES_SEM_RT_JSON.exists() else []
        _nomes_sheets = [c["nome"] for c in _todos]
        for _cl in _locais:
            if _cl["nome"] not in _nomes_sheets:
                _todos.append(_cl)
        return _todos

    _bs_clientes = _clientes_bs_contrato()
    _bs_nomes = ["— selecione ou preencha manualmente —"] + [c["nome"] for c in _bs_clientes]

    _bs_sel = st.selectbox("Carregar dados de cliente cadastrado", _bs_nomes,
        key="bs_cont_cliente_sel")

    if st.button("📂 Carregar dados do cliente", key="btn_bs_cont_carregar"):
        _bs_dado = next((c for c in _bs_clientes if c["nome"] == _bs_sel), {})
        if _bs_dado:
            st.session_state["bs_cont_nome"]     = _bs_dado.get("nome", "")
            st.session_state["bs_cont_cnpj"]     = _bs_dado.get("cnpj", "")
            st.session_state["bs_cont_endereco"] = _bs_dado.get("endereco", "")
            st.session_state["bs_cont_contato"]  = _bs_dado.get("contato", "")
            st.session_state["bs_cont_telefone"] = _bs_dado.get("telefone", "")
            st.success(f"✅ Dados de '{_bs_dado['nome']}' carregados.")
            st.rerun()

    st.markdown("---")
    st.markdown("**Dados do Contratante**")

    _bc1, _bc2 = st.columns(2)
    with _bc1:
        bs_nome     = st.text_input("Nome / Razão social *", key="bs_cont_nome",
            placeholder="Ex.: Condomínio Residencial Bella Vista")
        bs_endereco = st.text_area("Endereço completo", key="bs_cont_endereco",
            height=70, placeholder="Rua, número, bairro, cidade/UF, CEP")
    with _bc2:
        bs_cnpj     = st.text_input("CPF / CNPJ", key="bs_cont_cnpj",
            placeholder="00.000.000/0000-00")
        bs_contato  = st.text_input("Representante / síndico", key="bs_cont_contato",
            placeholder="Nome completo do responsável")
        bs_telefone = st.text_input("Telefone / WhatsApp", key="bs_cont_telefone",
            placeholder="(34) 99999-9999")

    st.markdown("**Descrição da(s) piscina(s) atendida(s)**")
    bs_piscinas = st.text_area("Piscinas atendidas", key="bs_cont_piscinas",
        height=60,
        placeholder="Ex.: Piscina adulto (150 m³), piscina infantil (30 m³), descobertas")

    st.markdown("**Condições do serviço**")
    _bs_c1, _bs_c2, _bs_c3 = st.columns(3)
    with _bs_c1:
        bs_frequencia = st.selectbox("Frequência de visitas", 
            ["1 visita semanal", "2 visitas semanais", "3 visitas semanais", "Outra"],
            key="bs_cont_frequencia")
        if bs_frequencia == "Outra":
            bs_frequencia = st.text_input("Especificar frequência", key="bs_cont_freq_outro",
                placeholder="Ex.: quinzenal")
    with _bs_c2:
        bs_produtos = st.radio("Produtos químicos", 
            ["Incluídos no valor", "Não incluídos (por conta do contratante)"],
            key="bs_cont_produtos")
    with _bs_c3:
        bs_prazo = st.text_input("Prazo de vigência (meses)", key="bs_cont_prazo",
            placeholder="Ex.: 12")

    st.markdown("**Valores e pagamento**")
    _bs_v1, _bs_v2, _bs_v3, _bs_v4 = st.columns(4)
    with _bs_v1:
        bs_valor = st.text_input("Valor mensal (R$) *", key="bs_cont_valor",
            placeholder="Ex.: 350,00")
    with _bs_v2:
        bs_valor_extenso = st.text_input("Valor por extenso", key="bs_cont_valor_extenso",
            placeholder="Ex.: trezentos e cinquenta reais")
    with _bs_v3:
        bs_vencimento = st.text_input("Dia de vencimento", key="bs_cont_vencimento",
            placeholder="Ex.: 10")
    with _bs_v4:
        bs_pagamento = st.selectbox("Forma de pagamento",
            ["Pix", "Boleto", "Transferência bancária", "Dinheiro", "Outro"],
            key="bs_cont_pagamento")

    st.markdown("**Datas**")
    _bs_d1, _bs_d2, _bs_d3 = st.columns(3)
    with _bs_d1:
        bs_data_inicio = st.text_input("Data de início", key="bs_cont_data_inicio",
            placeholder="dd/mm/aaaa", value=hoje_br())
    with _bs_d2:
        bs_data_fim = st.text_input("Data de término", key="bs_cont_data_fim",
            placeholder="dd/mm/aaaa")
    with _bs_d3:
        bs_local_ass = st.text_input("Local de assinatura", key="bs_cont_local",
            placeholder="Ex.: Uberlândia/MG", value="Uberlândia/MG")
    bs_data_ass = st.text_input("Data de assinatura", key="bs_cont_data_ass",
        placeholder="dd/mm/aaaa", value=hoje_br())

    st.markdown("---")

    if st.button("📄 Gerar Contrato Bem Star (PDF)", type="primary", use_container_width=True,
            key="btn_gerar_contrato_bs"):
        if not (st.session_state.get("bs_cont_nome","")).strip():
            st.error("Informe o nome do contratante.")
        elif not (st.session_state.get("bs_cont_valor","")).strip():
            st.error("Informe o valor mensal.")
        else:
            try:
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import cm
                from reportlab.lib import colors
                from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                    Table, TableStyle, HRFlowable)
                from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
                import io as _io

                # ── Coleta valores ─────────────────────────────────────────
                _nome     = (st.session_state.get("bs_cont_nome","")).strip()
                _cnpj     = (st.session_state.get("bs_cont_cnpj","")).strip()
                _end      = (st.session_state.get("bs_cont_endereco","")).strip()
                _contato  = (st.session_state.get("bs_cont_contato","")).strip()
                _tel      = (st.session_state.get("bs_cont_telefone","")).strip()
                _piscinas = (st.session_state.get("bs_cont_piscinas","")).strip()
                _freq     = st.session_state.get("bs_cont_freq_outro","").strip() or                             st.session_state.get("bs_cont_frequencia","")
                _prods_inc = "incluídos" in st.session_state.get("bs_cont_produtos","").lower()
                _prazo    = (st.session_state.get("bs_cont_prazo","")).strip() or "12"
                _valor    = (st.session_state.get("bs_cont_valor","")).strip()
                _ext      = (st.session_state.get("bs_cont_valor_extenso","")).strip() or _valor
                _venc     = (st.session_state.get("bs_cont_vencimento","")).strip() or "10"
                _pgto     = st.session_state.get("bs_cont_pagamento","Pix")
                _inicio   = (st.session_state.get("bs_cont_data_inicio","")).strip() or hoje_br()
                _fim      = (st.session_state.get("bs_cont_data_fim","")).strip() or "—"
                _local    = (st.session_state.get("bs_cont_local","")).strip() or "Uberlândia/MG"
                _data_ass = (st.session_state.get("bs_cont_data_ass","")).strip() or hoje_br()
                _qualif   = f"inscrito(a) no CPF/CNPJ sob nº {_cnpj}," if _cnpj else ""
                _piscinas_txt = _piscinas or "conforme descrição operacional acordada entre as partes"
                _prod_txt = "estão incluídos no valor mensal contratado" if _prods_inc                             else "não estão incluídos no valor mensal contratado"

                # ── Estilos ReportLab ──────────────────────────────────────
                styles = getSampleStyleSheet()
                s_titulo = ParagraphStyle("titulo", parent=styles["Heading1"],
                    fontSize=14, alignment=TA_CENTER, spaceAfter=4, textColor=colors.HexColor("#0d3d75"))
                s_sub = ParagraphStyle("sub", parent=styles["Normal"],
                    fontSize=11, alignment=TA_CENTER, spaceAfter=2, textColor=colors.HexColor("#0d3d75"))
                s_clausula = ParagraphStyle("clausula", parent=styles["Normal"],
                    fontSize=10, spaceBefore=10, spaceAfter=3, fontName="Helvetica-Bold",
                    textColor=colors.HexColor("#0d3d75"),
                    borderPad=4, borderColor=colors.HexColor("#0d3d75"),
                    leftIndent=0)
                s_body = ParagraphStyle("body", parent=styles["Normal"],
                    fontSize=9.5, alignment=TA_JUSTIFY, spaceBefore=2, spaceAfter=4,
                    leading=14, leftIndent=8)
                s_center = ParagraphStyle("center", parent=styles["Normal"],
                    fontSize=10, alignment=TA_CENTER, spaceBefore=4)
                s_small = ParagraphStyle("small", parent=styles["Normal"],
                    fontSize=8, textColor=colors.grey, alignment=TA_CENTER)

                # ── Monta story ────────────────────────────────────────────
                story = []

                # Logo Bem Star se disponível
                _logo_bs = encontrar_logo_bem_star()
                if _logo_bs and _logo_bs.exists():
                    from reportlab.platypus import Image as RLImage
                    _img = RLImage(str(_logo_bs), width=7*cm, height=2.5*cm,
                        kind="proportional")
                    _img.hAlign = "CENTER"
                    story.append(_img)
                    story.append(Spacer(1, 0.4*cm))

                story.append(Paragraph("CONTRATO DE PRESTAÇÃO DE SERVIÇOS", s_sub))
                story.append(Paragraph("Limpeza e Manutenção de Piscinas", ParagraphStyle(
                    "sub2", parent=styles["Normal"], fontSize=10,
                    alignment=TA_CENTER, textColor=colors.HexColor("#5d7288"), spaceAfter=4)))
                story.append(Spacer(1, 0.3*cm))
                story.append(HRFlowable(width="100%", thickness=2,
                    color=colors.HexColor("#0d3d75")))
                story.append(Spacer(1, 0.3*cm))

                # Tabela identificação
                id_data = [
                    ["CONTRATADA", "BEM STAR PISCINAS LTDA., CNPJ 26.799.958/0001-88\nAv. Getúlio Vargas, 4411, Jardim das Palmeiras, Uberlândia/MG, CEP 38.412-316"],
                    ["CONTRATANTE", f"{_nome}{', ' + _qualif if _qualif else ''} com endereço em {_end or '—'}."],
                ]
                t_id = Table(id_data, colWidths=[3.5*cm, 14*cm])
                t_id.setStyle(TableStyle([
                    ("FONTNAME",  (0,0), (0,-1), "Helvetica-Bold"),
                    ("FONTSIZE",  (0,0), (-1,-1), 9),
                    ("VALIGN",    (0,0), (-1,-1), "MIDDLE"),
                    ("BOX",       (0,0), (-1,-1), 1, colors.HexColor("#0d3d75")),
                    ("INNERGRID", (0,0), (-1,-1), 0.5, colors.HexColor("#c0c8d8")),
                    ("BACKGROUND",(0,0), (0,-1), colors.HexColor("#0d3d75")),
                    ("TEXTCOLOR", (0,0), (0,-1), colors.white),
                    ("TOPPADDING",(0,0),(-1,-1), 7),
                    ("BOTTOMPADDING",(0,0),(-1,-1), 7),
                    ("LEFTPADDING",(0,0),(-1,-1), 8),
                    ("RIGHTPADDING",(0,0),(-1,-1), 8),
                ]))
                story.append(t_id)
                story.append(Spacer(1, 0.3*cm))
                story.append(Paragraph(
                    "As partes acima identificadas têm entre si justo e contratado o presente "
                    "instrumento, regido pelas cláusulas e condições seguintes.", s_body))

                # Cláusulas
                clausulas = [
                    ("CLÁUSULA 1 — DO OBJETO",
                     "O presente contrato tem por objeto a prestação, pela CONTRATADA, de serviços regulares "
                     "de limpeza, conservação e manutenção operacional de piscina(s) localizada(s) no "
                     f"endereço do CONTRATANTE. Piscina(s) atendida(s): {_piscinas_txt}. "
                     "Os serviços abrangem: aspiração, escovação de paredes e bordas, peneiração e retirada "
                     "de resíduos, limpeza de cestos de skimmer e pré-filtro, acompanhamento visual das "
                     "condições da água e operações rotineiras de circulação e filtração. Este contrato "
                     "não inclui obras civis, substituição estrutural de equipamentos, reformas hidráulicas, "
                     "reparos elétricos, laudos, perícias ou outros serviços extraordinários não previstos."),

                    ("CLÁUSULA 2 — DA FREQUÊNCIA E EXECUÇÃO",
                     f"Os serviços serão executados com a seguinte frequência: {_freq}. "
                     "As visitas ocorrerão em dias e horários definidos conforme programação operacional "
                     "da CONTRATADA, podendo haver ajustes por necessidade climática, operacional, "
                     "feriados, caso fortuito ou força maior. Serviços extraordinários, emergenciais "
                     "ou fora da rotina poderão ser cobrados à parte, mediante comunicação prévia."),

                    ("CLÁUSULA 3 — DOS PRODUTOS E MATERIAIS",
                     f"Os produtos químicos, acessórios, insumos e materiais consumíveis {_prod_txt}. "
                     + ("Quando não incluídos, caberá ao CONTRATANTE providenciar todos os produtos e "
                        "materiais necessários em quantidade e qualidade suficientes para a execução dos serviços. "
                        "A falta de produtos ou condições inadequadas poderá impactar a qualidade do resultado "
                        "operacional, sem que isso caracterize inadimplemento da CONTRATADA."
                        if not _prods_inc else "")),

                    ("CLÁUSULA 4 — DO PREÇO E DO PAGAMENTO",
                     f"Pela prestação dos serviços, o CONTRATANTE pagará à CONTRATADA o valor mensal de "
                     f"R$ {_valor} ({_ext}). O vencimento ocorrerá todo dia {_venc} de cada mês, "
                     f"mediante {_pgto}. O atraso sujeitará o CONTRATANTE a multa de 2%, juros de 1% "
                     "ao mês pro rata die e correção monetária. Persistindo a inadimplência, a CONTRATADA "
                     "poderá suspender os serviços após comunicação prévia."),

                    ("CLÁUSULA 5 — DO PRAZO DE VIGÊNCIA",
                     f"O presente contrato vigorará pelo prazo de {_prazo} meses, com início em {_inicio} "
                     f"e término em {_fim}. Findo o prazo, poderá ser renovado por acordo entre as partes, "
                     "inclusive de forma tácita, caso a prestação prossiga sem oposição expressa. "
                     "Em contratos superiores a 12 meses, o valor poderá ser reajustado anualmente pelo IPCA/IBGE."),

                    ("CLÁUSULA 6 — DAS OBRIGAÇÕES DAS PARTES",
                     "A CONTRATADA executará os serviços com zelo, técnica e boa-fé; informará o "
                     "CONTRATANTE sobre irregularidades que interfiram na conservação da piscina; e "
                     "manterá sigilo sobre informações não públicas. "
                     "O CONTRATANTE garantirá livre acesso ao local; manterá os sistemas básicos em "
                     "funcionamento; comunicará previamente alterações relevantes de uso, eventos ou reformas; "
                     "e efetuará o pagamento na forma e prazo pactuados."),

                    ("CLÁUSULA 7 — DAS LIMITAÇÕES DE RESPONSABILIDADE",
                     "A CONTRATADA responde pela execução dos serviços dentro do escopo previsto, "
                     "não se responsabilizando por falhas estruturais preexistentes ou supervenientes; "
                     "defeitos elétricos, hidráulicos ou mecânicos fora do escopo; danos decorrentes de "
                     "mau uso, acesso indevido de terceiros, vandalismo, intempéries ou ausência de insumos."),

                    ("CLÁUSULA 8 — DA RESCISÃO",
                     "Este contrato poderá ser rescindido por mútuo acordo; por qualquer das partes, "
                     "mediante aviso prévio por escrito de 30 dias; imediatamente, em caso de "
                     "descumprimento contratual relevante após notificação sem saneamento; ou por "
                     "inadimplência do CONTRATANTE. Permanecerão exigíveis os valores já vencidos e "
                     "serviços efetivamente prestados."),

                    ("CLÁUSULA 9 — DAS DISPOSIÇÕES GERAIS",
                     "Os dados fornecidos serão utilizados exclusivamente para execução do contrato, "
                     "comunicações operacionais e rotinas administrativas. Qualquer alteração de escopo, "
                     "frequência, preço ou condição relevante deverá ser formalizada por escrito. "
                     "Fica eleito o foro da Comarca de Uberlândia/MG para dirimir quaisquer controvérsias "
                     "oriundas deste contrato, com renúncia de qualquer outro, por mais privilegiado que seja."),
                ]

                for titulo_cl, texto_cl in clausulas:
                    story.append(Paragraph(titulo_cl, s_clausula))
                    story.append(Paragraph(texto_cl, s_body))

                story.append(Spacer(1, 0.5*cm))
                story.append(HRFlowable(width="100%", thickness=0.5,
                    color=colors.HexColor("#c0c8d8")))
                story.append(Spacer(1, 0.3*cm))
                story.append(Paragraph(
                    f"E, por estarem justas e contratadas, firmam o presente instrumento em 2 (duas) "
                    f"vias de igual teor e forma.", s_body))
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph(f"{_local}, {_data_ass}.", s_center))
                story.append(Spacer(1, 1*cm))

                # Tabela de assinaturas
                ass_data = [
                    ["___________________________________",
                     "___________________________________"],
                    ["BEM STAR PISCINAS LTDA.\nCONTRATADA",
                     f"{_nome}\nCONTRATANTE"],
                    ["", ""],
                    ["___________________________________",
                     "___________________________________"],
                    ["TESTEMUNHA 1\nNome:\nCPF:",
                     "TESTEMUNHA 2\nNome:\nCPF:"],
                ]
                t_ass = Table(ass_data, colWidths=[9*cm, 9*cm])
                t_ass.setStyle(TableStyle([
                    ("ALIGN",    (0,0), (-1,-1), "CENTER"),
                    ("FONTSIZE", (0,0), (-1,-1), 9),
                    ("VALIGN",   (0,0), (-1,-1), "TOP"),
                    ("TOPPADDING", (0,0),(-1,-1), 4),
                ]))
                story.append(t_ass)
                story.append(Spacer(1, 0.3*cm))
                story.append(HRFlowable(width="100%", thickness=0.5,
                    color=colors.HexColor("#0d3d75")))
                story.append(Spacer(1, 0.15*cm))
                story.append(Paragraph(
                    "Bem Star Piscinas LTDA. | CNPJ 26.799.958/0001-88 | "
                    "Av. Getúlio Vargas, 4411, Uberlândia/MG | Documento de uso operacional",
                    s_small))

                # ── Gera PDF ───────────────────────────────────────────────
                pasta_bs_cont = GENERATED_DIR / slugify_nome(_nome)
                pasta_bs_cont.mkdir(parents=True, exist_ok=True)
                _ts_bs = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_bs_path = pasta_bs_cont / f"{_ts_bs}_{slugify_nome(_nome)}_CONTRATO_BS.pdf"

                _buf = _io.BytesIO()
                doc_rl = SimpleDocTemplate(
                    _buf,
                    pagesize=A4,
                    topMargin=2*cm, bottomMargin=2*cm,
                    leftMargin=2.5*cm, rightMargin=2.5*cm,
                    title=f"Contrato Bem Star — {_nome}",
                    author="Bem Star Piscinas",
                )
                doc_rl.build(story)
                _pdf_bytes = _buf.getvalue()

                with open(pdf_bs_path, "wb") as _pf:
                    _pf.write(_pdf_bytes)

                st.success(f"✅ Contrato gerado para {_nome}!")
                st.download_button(
                    "⬇️ Baixar Contrato PDF",
                    data=_pdf_bytes,
                    file_name=pdf_bs_path.name,
                    mime="application/pdf",
                    use_container_width=True,
                    key="btn_dl_contrato_bs",
                )

                # ── Bloco de envio ────────────────────────────────────────
                _msg_cont = montar_mensagem_bem_star(
                    nome_local   = _nome,
                    responsavel  = _contato,
                    tipo         = "contrato",
                )
                exibir_bloco_envio_bem_star(
                    nome_local   = _nome,
                    pasta        = pasta_bs_cont,
                    telefone     = _tel,
                    email        = "",
                    mensagem     = _msg_cont,
                    key_suffix   = "contrato",
                )

            except Exception as _e:
                st.error(f"Erro ao gerar contrato: {_e}")
                import traceback as _tb
                st.code(_tb.format_exc(), language="text")

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# FORMULÁRIO
# =========================================

st.markdown('<div class="section-card aq-only" id="sec-formulario">', unsafe_allow_html=True)
st.subheader("Dados do contrato e aditivo")

# ── Seletor de cliente do Sheets ──────────────────────────────────────────────
@st.cache_data(ttl=60)
def _clientes_completos_cache():
    return sheets_listar_clientes_completo()

_clientes_rt = _clientes_completos_cache()
if _clientes_rt:
    _opcoes_rt = ["— Selecionar cliente cadastrado —"] + [c["nome"] for c in _clientes_rt]
    _sel_rt = st.selectbox(
        "🔗 Carregar dados de cliente cadastrado",
        _opcoes_rt,
        key="sel_cliente_rt",
        help="Selecione um cliente para preencher automaticamente os campos abaixo."
    )
    if _sel_rt and _sel_rt != "— Selecionar cliente cadastrado —":
        if st.button("⬇️ Preencher formulário com dados deste cliente", key="btn_carregar_cliente_rt", use_container_width=True):
            _dados_rt = next((c for c in _clientes_rt if c["nome"] == _sel_rt), {})
            if _dados_rt:
                st.session_state["nome_condominio"]   = _dados_rt.get("nome", "")
                st.session_state["cnpj_condominio"]   = formatar_cnpj(_dados_rt.get("cnpj", ""))
                st.session_state["endereco_condominio"] = _dados_rt.get("endereco", "")
                st.session_state["nome_sindico"]      = _dados_rt.get("contato", "")
                st.session_state["whatsapp_cliente"]  = formatar_telefone(_dados_rt.get("telefone", ""))
                st.session_state["email_cliente"]     = _dados_rt.get("email", "")
                st.success(f"✅ Dados de '{_sel_rt}' carregados! Ajuste o que precisar antes de gerar.")
                st.rerun()
else:
    st.info("💡 Cadastre clientes na seção acima para carregar os dados automaticamente aqui.")

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.text_input("Nome do condomínio", key="nome_condominio", on_change=on_change_nome_condominio)
    st.text_input(
        "CNPJ do condomínio",
        key="cnpj_condominio",
        on_change=on_change_cnpj,
        placeholder="00.000.000/0000-00",
    )
    if st.session_state.get("_rt_cep_fmt"):
        st.session_state["rt_cep"] = st.session_state.pop("_rt_cep_fmt")
    _rt_cep_c1, _rt_cep_c2 = st.columns([3, 1])
    with _rt_cep_c1:
        st.text_input("CEP", key="rt_cep", placeholder="00000-000",
            help="Digite o CEP e clique em 🔍 para preencher o endereço automaticamente")
    with _rt_cep_c2:
        st.markdown("<br>", unsafe_allow_html=True)
        _btn_rt_cep = st.button("🔍", key="btn_buscar_cep_rt", help="Buscar CEP")
    if _btn_rt_cep:
        _cep_v = re.sub(r"\D", "", st.session_state.get("rt_cep", ""))
        if len(_cep_v) == 8:
            with st.spinner("Buscando CEP..."):
                _dc = buscar_cep(_cep_v)
            if _dc:
                _end = ", ".join(p for p in [_dc.get("logradouro",""), _dc.get("bairro",""), f"{_dc.get('localidade','')}/{_dc.get('uf','')}", f"{_cep_v[:5]}-{_cep_v[5:]}"] if p)
                st.session_state["endereco_condominio"] = _end
                st.session_state["_rt_cep_fmt"] = f"{_cep_v[:5]}-{_cep_v[5:]}"
                st.rerun()
            else:
                st.error("CEP não encontrado.")
        else:
            st.warning("Digite um CEP válido com 8 dígitos.")
    st.text_area("Endereço do condomínio", key="endereco_condominio", height=100)
    st.text_input("Nome do síndico / representante", key="nome_sindico")
    st.text_input(
        "CPF do síndico / representante",
        key="cpf_sindico",
        on_change=on_change_cpf,
        placeholder="000.000.000-00",
    )

with col2:
    st.text_input(
        "Valor mensal",
        key="valor_mensal",
        on_change=on_change_valor_mensal,
        placeholder="R$ 1.621,00",
    )
    st.text_input(
        "Valor com desconto/aditivo",
        key="valor_aditivo",
        on_change=on_change_valor_aditivo,
        placeholder="R$ 810,50",
    )
    st.text_input(
        "Data de início",
        key="data_inicio",
        placeholder="01/04/2026",
        on_change=on_change_data_inicio,
    )
    st.text_input(
        "Data de fim",
        key="data_fim",
        placeholder="31/03/2027",
        on_change=on_change_data_fim,
    )
    st.text_input(
        "Data de assinatura",
        key="data_assinatura",
        placeholder="dd/mm/aaaa",
        on_change=on_change_data_assinatura,
    )

if modo == "Modo Campo":
    st.markdown("---")
    col_campo1, col_campo2 = st.columns(2)
    with col_campo1:
        st.text_input(
            "WhatsApp do cliente (opcional)",
            key="whatsapp_cliente",
            on_change=on_change_whatsapp,
            placeholder="(34) 99999-9999",
        )
    with col_campo2:
        st.text_input("E-mail do cliente (opcional)", key="email_cliente")
else:
    st.markdown("---")
    col_cont1, col_cont2 = st.columns(2)
    with col_cont1:
        st.text_input(
            "WhatsApp do cliente",
            key="whatsapp_cliente",
            on_change=on_change_whatsapp,
            placeholder="(34) 99999-9999",
        )
    with col_cont2:
        st.text_input("E-mail do cliente", key="email_cliente")

st.text_area(
    "Observações internas (não vai para contrato/aditivo)",
    key="observacoes_internas",
    height=100,
    placeholder="Ex.: condição comercial específica, observação operacional, histórico jurídico...",
)

st.markdown('</div>', unsafe_allow_html=True)

# =========================================
# AÇÕES DE CADASTRO / RENOVAÇÃO
# =========================================

st.markdown('<div class="section-card aq-only" id="sec-cadastro-renovacao">', unsafe_allow_html=True)
st.subheader("Cadastro e renovação")

r1, r2, r3 = st.columns([1.15, 1.15, 2])

with r1:
    if st.button("Salvar cadastro / atualizar JSON", use_container_width=True):
        ok, msg = salvar_cadastro_sem_gerar_documentos()
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

with r2:
    if st.button("Renovação anual rápida", use_container_width=True):
        novo_inicio, novo_fim = calcular_renovacao_anual(st.session_state.data_fim)
        if not novo_inicio or not novo_fim:
            st.error("Não foi possível renovar. Verifique se a data final atual está válida.")
        else:
            st.session_state.data_inicio = formatar_data_br(novo_inicio)
            st.session_state.data_fim = formatar_data_br(novo_fim)
            st.session_state.data_assinatura = hoje_br()
            st.success("Nova vigência anual preenchida automaticamente.")
            st.rerun()

with r3:
    st.caption(
        "Salvar cadastro atualiza apenas o JSON do condomínio. "
        "Renovação anual rápida ajusta datas no formulário sem gerar documento automaticamente."
    )

st.markdown("---")
st.markdown("**Geração de documentos contratuais**")

col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1.5, 1.5, 1, 1])

with col_btn1:
    gerar = st.button(
        "✅ Gerar contrato + aditivo",
        type="primary",
        use_container_width=True,
    )

with col_btn2:
    gerar_aditivo_rapido = st.button(
        "📄 Gerar somente aditivo rápido",
        use_container_width=True,
    )

with col_btn3:
    if st.button("🗑️ Limpar formulário", use_container_width=True):
        limpar_formulario()
        st.rerun()

with col_btn4:
    if st.button("📁 Abrir pasta gerada", use_container_width=True):
        abrir_pasta_windows(GENERATED_DIR)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# DOWNLOADS DOS ÚLTIMOS DOCUMENTOS GERADOS
# =========================================
# Exibe os downloads logo aqui, antes do relatório, independente de onde o botão foi clicado.

_ultimos = st.session_state.get("ultimos_docs_gerados")
if _ultimos:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("⬇️ Últimos documentos gerados")
    _dc1, _dc2 = st.columns(2)
    with _dc1:
        st.markdown("**Contrato**")
        _p = _ultimos.get("contrato_docx")
        if _p and Path(_p).exists():
            with open(_p, "rb") as _f:
                st.download_button("Baixar DOCX do contrato", data=_f, file_name=Path(_p).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_contrato_docx_top")
        _p = _ultimos.get("contrato_pdf")
        if _p and Path(_p).exists():
            with open(_p, "rb") as _f:
                st.download_button("Baixar PDF do contrato", data=_f, file_name=Path(_p).name,
                    mime="application/pdf", use_container_width=True, key="dl_contrato_pdf_top")
    with _dc2:
        st.markdown("**Aditivo**")
        _p = _ultimos.get("aditivo_docx")
        if _p and Path(_p).exists():
            with open(_p, "rb") as _f:
                st.download_button("Baixar DOCX do aditivo", data=_f, file_name=Path(_p).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_aditivo_docx_top")
        _p = _ultimos.get("aditivo_pdf")
        if _p and Path(_p).exists():
            with open(_p, "rb") as _f:
                st.download_button("Baixar PDF do aditivo", data=_f, file_name=Path(_p).name,
                    mime="application/pdf", use_container_width=True, key="dl_aditivo_pdf_top")
    _p_rel = _ultimos.get("relatorio_docx")
    _p_rel_pdf = _ultimos.get("relatorio_pdf")
    if (_p_rel and Path(_p_rel).exists()) or (_p_rel_pdf and Path(_p_rel_pdf).exists()):
        st.markdown("**Relatório mensal**")
        _dc3, _dc4 = st.columns(2)
        with _dc3:
            if _p_rel and Path(_p_rel).exists():
                with open(_p_rel, "rb") as _f:
                    st.download_button("Baixar DOCX do relatório", data=_f, file_name=Path(_p_rel).name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_relatorio_docx_top")
        with _dc4:
            if _p_rel_pdf and Path(_p_rel_pdf).exists():
                with open(_p_rel_pdf, "rb") as _f:
                    st.download_button("Baixar PDF do relatório", data=_f, file_name=Path(_p_rel_pdf).name,
                        mime="application/pdf", use_container_width=True, key="dl_relatorio_pdf_top")
    st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# RELATÓRIO MENSAL DE RT
# =========================================

st.markdown('<div class="section-card aq-only" id="sec-relatorio-rt">', unsafe_allow_html=True)
st.subheader("Relatório mensal de responsabilidade técnica")


st.caption(f"Dados fixos automáticos do RT: {RESPONSAVEL_TECNICO_ASSINATURA} | Certificações {CERTIFICACOES_RT}")

# ── Seletor de cliente do Sheets ──────────────────────────────────────────────
@st.cache_data(ttl=60)
def _clientes_completos_rel_cache():
    return sheets_listar_clientes_completo()

_clientes_rel = _clientes_completos_rel_cache()
if _clientes_rel:
    _opcoes_rel = ["— Selecionar cliente cadastrado —"] + [c["nome"] for c in _clientes_rel]
    _rel_col1, _rel_col2 = st.columns([3, 1])
    with _rel_col1:
        _sel_rel = st.selectbox(
            "🔗 Carregar dados de cliente cadastrado",
            _opcoes_rel,
            key="sel_cliente_rel",
            help="Selecione um cliente para preencher automaticamente os campos do relatório."
        )
    with _rel_col2:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        if st.button("⬇️ Carregar", key="btn_carregar_cliente_rel", use_container_width=True):
            if _sel_rel and _sel_rel != "— Selecionar cliente cadastrado —":
                _dados_rel = next((c for c in _clientes_rel if c["nome"] == _sel_rel), {})
                if _dados_rel:
                    st.session_state["rel_nome_condominio"]   = _dados_rel.get("nome", "")
                    st.session_state["rel_cnpj_condominio"]   = formatar_cnpj(_dados_rel.get("cnpj", ""))
                    st.session_state["rel_endereco_condominio"] = _dados_rel.get("endereco", "")
                    st.session_state["rel_representante"]     = _dados_rel.get("contato", "")
                    st.session_state["rel_cpf_cnpj_representante"] = ""
                    st.success(f"✅ Dados de '{_sel_rel}' carregados no relatório!")
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

rr0a, rr0b, rr0c = st.columns([1.1, 1.2, 1.1])
with rr0a:
    st.selectbox("Tipo de atendimento", ["Contrato ativo", "Visita técnica avulsa", "Inspeção técnica", "Acompanhamento sem contrato"], key="rel_tipo_atendimento")
with rr0b:
    if st.button("Carregar dados do cadastro no relatório", use_container_width=True):
        carregar_dados_cadastro_no_relatorio()
        st.rerun()
with rr0c:
    st.checkbox("Salvar alterações deste relatório no cadastro principal", key="rel_salvar_alteracoes_cadastro")

# ---- Importar lançamentos de campo (local + Google Sheets) ----
nome_rel_atual = (st.session_state.get("rel_nome_condominio") or st.session_state.get("nome_condominio") or "").strip()
pasta_rel_atual = GENERATED_DIR / slugify_nome(nome_rel_atual) if nome_rel_atual else None

# Busca lançamentos locais
lancamentos_local = []
if pasta_rel_atual and pasta_rel_atual.exists():
    dados_rel_json = carregar_dados_condominio(pasta_rel_atual)
    lancamentos_local = (dados_rel_json or {}).get("lancamentos_campo", [])

# Busca lançamentos do Google Sheets (aba Visitas)
lancamentos_sheets = []
if nome_rel_atual:
    lancamentos_sheets = sheets_listar_lancamentos(nome_rel_atual)

# Filtra pelo mês de referência
mes_ref = (st.session_state.get("rel_mes_referencia") or "").strip()
ano_ref = (st.session_state.get("rel_ano_referencia") or str(datetime.now().year)).strip()

def _filtrar_mes(lancamentos, mes, ano):
    if not mes or not ano:
        return lancamentos
    filtrados = []
    for lc in lancamentos:
        data = lc.get("data","")
        # Formato dd/mm/aaaa ou aaaa-mm-dd
        try:
            if "/" in data:
                partes = data.split("/")
                if len(partes) == 3 and partes[1] == mes.zfill(2) and partes[2] == ano:
                    filtrados.append(lc)
            elif "-" in data:
                partes = data.split("-")
                if len(partes) == 3 and partes[1] == mes.zfill(2) and partes[0] == ano:
                    filtrados.append(lc)
        except Exception:
            filtrados.append(lc)
    return filtrados

# Une local + Sheets sem duplicar (por data+operador)
_vistos = set()
lancamentos_disponiveis = []
for lc in lancamentos_local + lancamentos_sheets:
    _chave = f"{lc.get('data','')}-{lc.get('operador','')}-{lc.get('ph','')}"
    if _chave not in _vistos:
        _vistos.add(_chave)
        lancamentos_disponiveis.append(lc)

# Filtra por mês se informado
lancamentos_disponiveis = _filtrar_mes(lancamentos_disponiveis, mes_ref, ano_ref)

def _importar_lancamentos(lancamentos):
    """Preenche o relatório com os lançamentos de campo."""
    garantir_campos_analises(max(len(lancamentos), ANALISES_PADRAO))
    for i, lc in enumerate(lancamentos[:ANALISES_MAX_SUGERIDO]):
        # Suporte a múltiplas piscinas — usa dados da primeira piscina ou direto
        piscinas = lc.get("piscinas", [])
        if piscinas:
            lc_dados = piscinas[0]  # primeira piscina para o relatório principal
        else:
            lc_dados = lc
        st.session_state[f"rel_analise_data_{i}"]     = lc.get("data", "")
        st.session_state[f"rel_analise_ph_{i}"]        = lc_dados.get("ph", lc.get("ph",""))
        st.session_state[f"rel_analise_cl_{i}"]        = lc_dados.get("cloro_livre", lc.get("cloro_livre",""))
        st.session_state[f"rel_analise_ct_{i}"]        = lc_dados.get("cloro_total", lc.get("cloro_total",""))
        st.session_state[f"rel_analise_alc_{i}"]       = lc_dados.get("alcalinidade", lc.get("alcalinidade",""))
        st.session_state[f"rel_analise_dc_{i}"]        = lc_dados.get("dureza", lc.get("dureza",""))
        st.session_state[f"rel_analise_cya_{i}"]       = lc_dados.get("cianurico", lc.get("cianurico",""))
        st.session_state[f"rel_analise_operador_{i}"]  = lc.get("operador", "")
    # Importa dosagens da última visita com dosagem registrada
    for lc in reversed(lancamentos):
        if lc.get("dosagens"):
            aplicar_dosagens_ultimas_no_relatorio({"dosagens_ultimas": lc["dosagens"]})
            break
    # Concatena observações e problemas
    obs_lista = []
    for lc in lancamentos:
        if lc.get("problemas","").strip():
            obs_lista.append(f"[{lc.get('data','')}] ⚠️ {lc['problemas']}")
        if lc.get("observacao","").strip():
            obs_lista.append(f"[{lc.get('data','')}] {lc['observacao']}")
    obs_txt = "\n".join(obs_lista[:10])
    if obs_txt:
        st.session_state["rel_observacoes_gerais"] = obs_txt

if lancamentos_disponiveis:
    _total = len(lancamentos_disponiveis)
    _fonte = "📱 local + Sheets" if lancamentos_sheets else "📱 local"
    _periodo = f"{lancamentos_disponiveis[0].get('data','?')} → {lancamentos_disponiveis[-1].get('data','?')}"

    st.markdown(f"""
    <div style="border:1px solid rgba(20,120,60,0.3);border-radius:12px;padding:12px 16px;
    background:rgba(20,120,60,0.07);margin-bottom:12px;">
    <strong>📱 {_total} lançamento(s) de campo disponível(is) — {_fonte}</strong><br>
    <span style="font-size:0.85rem;color:#3a6a3a;">Período: {_periodo}</span>
    </div>
    """, unsafe_allow_html=True)

    imp1, imp2, imp3 = st.columns([1.5, 1.5, 1])
    with imp1:
        if st.button("📥 Importar lançamentos para o relatório", use_container_width=True, type="primary"):
            _importar_lancamentos(lancamentos_disponiveis)
            st.success(f"✅ {_total} lançamento(s) importado(s) com sucesso!")
            st.rerun()

    with imp2:
        if st.button("🗑️ Limpar lançamentos locais após gerar", use_container_width=True):
            if pasta_rel_atual and pasta_rel_atual.exists():
                dados_limpar = carregar_dados_condominio(pasta_rel_atual) or {}
                dados_limpar["lancamentos_campo"] = []
                salvar_dados_condominio(pasta_rel_atual, dados_limpar)
                st.success("Lançamentos locais limpos.")
                st.rerun()

    with imp3:
        with st.expander("👁 Ver lançamentos"):
            for lc in lancamentos_disponiveis:
                piscinas = lc.get("piscinas",[])
                if piscinas:
                    for p in piscinas:
                        st.caption(f"📅 {lc.get('data','?')} | {p.get('nome','Piscina')} | pH:{p.get('ph','–')} CRL:{p.get('cloro_livre','–')}")
                else:
                    st.caption(f"📅 {lc.get('data','?')} | Op:{lc.get('operador','–')} | pH:{lc.get('ph','–')} CRL:{lc.get('cloro_livre','–')}")
                if lc.get("problemas","").strip():
                    st.caption(f"   ⚠️ {lc['problemas']}")
                if lc.get("observacao","").strip():
                    st.caption(f"   📝 {lc['observacao']}")

else:
    if nome_rel_atual:
        st.info(f"Nenhum lançamento de campo encontrado para **{nome_rel_atual}**{f' no mês {mes_ref}/{ano_ref}' if mes_ref else ''}. O operador precisa registrar as visitas pelo modo campo.")

st.markdown("**Dados do condomínio / local atendido**")
rd1, rd2 = st.columns(2)
with rd1:
    st.text_input("Condomínio / estabelecimento", key="rel_nome_condominio")
    if st.session_state.get("_rel_cep_fmt"):
        st.session_state["rel_cep"] = st.session_state.pop("_rel_cep_fmt")
    _rel_cep_c1, _rel_cep_c2 = st.columns([3, 1])
    with _rel_cep_c1:
        st.text_input("CEP", key="rel_cep", placeholder="00000-000",
            help="Digite o CEP e clique em 🔍 para preencher o endereço automaticamente")
    with _rel_cep_c2:
        st.markdown("<br>", unsafe_allow_html=True)
        _btn_rel_cep = st.button("🔍", key="btn_buscar_cep_rel", help="Buscar CEP")
    if _btn_rel_cep:
        _cep_v = re.sub(r"\D", "", st.session_state.get("rel_cep", ""))
        if len(_cep_v) == 8:
            with st.spinner("Buscando CEP..."):
                _dc = buscar_cep(_cep_v)
            if _dc:
                _end = ", ".join(p for p in [_dc.get("logradouro",""), _dc.get("bairro",""), f"{_dc.get('localidade','')}/{_dc.get('uf','')}", f"{_cep_v[:5]}-{_cep_v[5:]}"] if p)
                st.session_state["rel_endereco_condominio"] = _end
                st.session_state["_rel_cep_fmt"] = f"{_cep_v[:5]}-{_cep_v[5:]}"
                st.rerun()
            else:
                st.error("CEP não encontrado.")
        else:
            st.warning("Digite um CEP válido com 8 dígitos.")
    st.text_area("Endereço do local", key="rel_endereco_condominio", height=90)
    st.text_input("Representante / síndico / contato local", key="rel_representante")
with rd2:
    st.text_input(
        "CNPJ do condomínio / estabelecimento",
        key="rel_cnpj_condominio",
        on_change=lambda: st.session_state.__setitem__("rel_cnpj_condominio", formatar_cnpj(st.session_state.get("rel_cnpj_condominio", "")))
    )
    st.text_input("CPF/CNPJ do representante", key="rel_cpf_cnpj_representante", on_change=lambda: on_change_rel_documento_representante())
    st.file_uploader(
        "Upload de fotos do relatório",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="rel_fotos_upload",
        help="As fotos serão salvas na pasta do condomínio e inseridas no relatório."
    )
    st.caption("Esses dados podem ser preenchidos diretamente aqui, mesmo quando o relatório for avulso e sem contrato.")

r1, r2, r3, r4 = st.columns(4)
with r1:
    st.text_input("Mês de referência", key="rel_mes_referencia", placeholder="04")
with r2:
    st.text_input("Ano de referência", key="rel_ano_referencia", placeholder="2026")
with r3:
    st.selectbox("Status da ART", ["Emitida", "Não emitida", "Em tramitação"], key="rel_art_status", on_change=on_change_rel_art_status)
with r4:
    st.text_input("Data de emissão", key="rel_data_emissao", placeholder="dd/mm/aaaa", on_change=lambda: formatar_data_relatorio_chave("rel_data_emissao"))

r5, r6, r7, r8 = st.columns(4)
with r5:
    st.text_input("ART nº", key="rel_art_numero", on_change=on_change_rel_art_numero, disabled=st.session_state.get("rel_art_status") != "Emitida")
with r6:
    st.text_input("Vigência da ART - início", key="rel_art_inicio", placeholder="dd/mm/aaaa", on_change=lambda: formatar_data_relatorio_chave("rel_art_inicio"), disabled=st.session_state.get("rel_art_status") != "Emitida")
with r7:
    st.text_input("Vigência da ART - fim", key="rel_art_fim", placeholder="dd/mm/aaaa", on_change=lambda: formatar_data_relatorio_chave("rel_art_fim"), disabled=st.session_state.get("rel_art_status") != "Emitida")
with r8:
    opcoes_status_agua = ["CONFORME", "NÃO CONFORME", "EM CORREÇÃO"]
    status_atual = st.session_state.get("rel_status_agua", "CONFORME")
    if status_atual not in opcoes_status_agua:
        status_atual = "CONFORME"
    opcao_status = st.selectbox(
        "Status geral da água",
        opcoes_status_agua,
        index=opcoes_status_agua.index(status_atual),
        key="rel_status_agua_select",
    )
    st.session_state["rel_status_agua"] = opcao_status

if st.session_state.get("rel_art_status") != "Emitida":
    st.caption("Como a ART não está emitida, os campos ART nº e vigência ficam desabilitados e o relatório preencherá automaticamente como N/A, com observação institucional conforme o status selecionado.")

c_auto1, c_auto2 = st.columns([1,2])
with c_auto1:
    if st.button("Preencher parecer automático", use_container_width=True):
        aplicar_textos_automaticos_relatorio()
with c_auto2:
    st.caption("O sistema calcula não conformidades e cloro combinado (cloraminas = CT - CL), preenche diagnóstico, observações e recomendações, e você ainda pode editar antes de gerar o relatório.")

st.text_area("Diagnóstico técnico", key="rel_diagnostico", height=120, placeholder="Será preenchido automaticamente conforme os parâmetros, mas permanece editável.")

st.markdown("**Análises físico-químicas**")
garantir_campos_analises(st.session_state.get("rel_analises_total", ANALISES_PADRAO))
ctrl_a1, ctrl_a2 = st.columns([1, 3])
with ctrl_a1:
    if st.button("Adicionar análise extra", use_container_width=True):
        adicionar_analise_extra()
        st.rerun()
with ctrl_a2:
    st.caption(f"{st.session_state.get("rel_analises_total", ANALISES_PADRAO)} linha(s) disponíveis neste relatório. Base padrão: 9 análises mensais.")
for i in range(int(st.session_state.get("rel_analises_total", ANALISES_PADRAO) or ANALISES_PADRAO)):
    cols = st.columns([1.05,0.7,0.8,0.8,0.9,0.9,0.9,1.1])
    cols[0].text_input(f"Data {i+1}", key=f"rel_analise_data_{i}", label_visibility="collapsed", placeholder="dd/mm/aaaa", on_change=lambda chave=f"rel_analise_data_{i}": formatar_data_relatorio_chave(chave))
    cols[1].text_input(f"pH {i+1}", key=f"rel_analise_ph_{i}", label_visibility="collapsed", placeholder="pH")
    cols[2].text_input(f"CRL {i+1}", key=f"rel_analise_cl_{i}", label_visibility="collapsed", placeholder="CRL")
    cols[3].text_input(f"CT {i+1}", key=f"rel_analise_ct_{i}", label_visibility="collapsed", placeholder="CT")
    cols[4].text_input(f"ALC {i+1}", key=f"rel_analise_alc_{i}", label_visibility="collapsed", placeholder="Alc")
    cols[5].text_input(f"DC {i+1}", key=f"rel_analise_dc_{i}", label_visibility="collapsed", placeholder="DC")
    cols[6].text_input(f"CYA {i+1}", key=f"rel_analise_cya_{i}", label_visibility="collapsed", placeholder="CYA")
    cols[7].text_input(f"Operador {i+1}", key=f"rel_analise_operador_{i}", label_visibility="collapsed", placeholder="Operador")

st.markdown("**Dosagens de produtos químicos**")
ctrl_d1, ctrl_d2 = st.columns([1.1, 2.4])
with ctrl_d1:
    if st.button("Carregar usados pela última vez", use_container_width=True):
        nome_rel = (st.session_state.get("rel_nome_condominio") or st.session_state.get("nome_condominio") or "").strip()
        if nome_rel:
            pasta_rel = GENERATED_DIR / slugify_nome(nome_rel)
            dados_rel_salvos = carregar_dados_condominio(pasta_rel) if pasta_rel.exists() else None
            aplicar_dosagens_ultimas_no_relatorio(dados_rel_salvos)
        else:
            aplicar_dosagens_ultimas_no_relatorio(obter_snapshot_relatorio_independente())
        st.rerun()
with ctrl_d2:
    st.caption("Ao gerar o relatório, as dosagens deste condomínio passam a ficar salvas como usados pela última vez.")
for i in range(7):
    cols = st.columns([1.4,1.1,0.7,0.7,1.3])
    cols[0].text_input(f"Produto {i+1}", key=f"rel_dos_produto_{i}", label_visibility="collapsed", placeholder="Produto químico")
    cols[1].text_input(f"Lote {i+1}", key=f"rel_dos_lote_{i}", label_visibility="collapsed", placeholder="Fabricante/Lote")
    cols[2].text_input(f"Qtd {i+1}", key=f"rel_dos_qtd_{i}", label_visibility="collapsed", placeholder="Qtd")
    cols[3].text_input(f"Un {i+1}", key=f"rel_dos_un_{i}", label_visibility="collapsed", placeholder="Unid.")
    cols[4].text_input(f"Finalidade {i+1}", key=f"rel_dos_finalidade_{i}", label_visibility="collapsed", placeholder="Finalidade técnica")

st.markdown("**Observações automáticas / editáveis**")
for i in range(5):
    st.text_area(f"Observação {i+1}", key=f"rel_obs_{i}", height=70)

st.markdown("**Recomendações técnicas**")
for i in range(5):
    cols = st.columns([2.0,0.8,1.0])
    cols[0].text_input(f"Recomendação {i+1}", key=f"rel_rec_texto_{i}", label_visibility="collapsed", placeholder="Recomendação técnica")
    cols[1].text_input(f"Prazo {i+1}", key=f"rel_rec_prazo_{i}", label_visibility="collapsed", placeholder="Prazo")
    cols[2].text_input(f"Responsável {i+1}", key=f"rel_rec_resp_{i}", label_visibility="collapsed", placeholder="Responsável")

cx1, cx2, cx3 = st.columns(3)
with cx1:
    st.text_area("ABNT NBR 10339 / segurança operacional – Evidência / observação", key="rel_nbr_11238", height=90, placeholder="Profundidade, circulação, higienização, retrolavagem, condição geral operacional...")
with cx2:
    st.text_area("NR-26 / GHS – Checklist", key="rel_nr_26", height=90, placeholder="FISPQs/FDS, rótulos GHS, sinalização, incompatibilidade, emergência...")
with cx3:
    st.text_area("NR-06 / EPI – Checklist", key="rel_nr_06", height=90, placeholder="Fornecimento, fiscalização e uso dos EPIs...")

st.markdown("**EPIs — preenchimento rápido**")
status_epi_opcoes = ["Conforme", "Pendente", "Não informado", "N/A"]
for rotulo, chave_base in [
    ("Luvas", "luvas"),
    ("Óculos", "oculos"),
    ("Respirador", "respirador"),
    ("Botas", "botas"),
]:
    c_status, c_ca = st.columns([1.2, 1])
    with c_status:
        st.selectbox(
            f"{rotulo} — status",
            options=status_epi_opcoes,
            key=f"rel_epi_{chave_base}_status",
        )
    with c_ca:
        st.text_input(f"{rotulo} — CA nº (opcional)", key=f"rel_epi_{chave_base}_ca")

st.caption("O relatório mensal pode ser gerado independentemente de contrato. Basta preencher os dados do local neste próprio módulo.")

# ---- Barra de rascunho ----
_nome_rasc = (st.session_state.get("rel_nome_condominio") or "").strip()
_pasta_rasc = GENERATED_DIR / slugify_nome(_nome_rasc) if _nome_rasc else None
_rasc_existente = carregar_rascunho_relatorio(_pasta_rasc) if _pasta_rasc and _pasta_rasc.exists() else None

col_rasc1, col_rasc2, col_rasc3 = st.columns([1.2, 1.2, 1.6])
with col_rasc1:
    if st.button("💾 Salvar rascunho", use_container_width=True):
        if _nome_rasc:
            _pasta_rasc.mkdir(parents=True, exist_ok=True)
            salvar_rascunho_relatorio(_pasta_rasc)
            st.success("Rascunho salvo! Os dados serão restaurados mesmo após reiniciar o sistema.")
        else:
            st.warning("Informe o nome do condomínio antes de salvar o rascunho.")

with col_rasc2:
    if _rasc_existente:
        if st.button(f"↩️ Restaurar rascunho ({_rasc_existente.get('salvo_em','?')})", use_container_width=True):
            aplicar_rascunho_no_formulario(_rasc_existente)
            st.success("Rascunho restaurado!")
            st.rerun()
    else:
        st.button("↩️ Sem rascunho salvo", disabled=True, use_container_width=True)

with col_rasc3:
    if _rasc_existente:
        st.markdown(
            f"<div style='font-size:0.82rem;color:#3a7a3a;padding:8px 0;'>"
            f"✅ Rascunho disponível — salvo em {_rasc_existente.get('salvo_em','?')}</div>",
            unsafe_allow_html=True
        )
    else:
        st.caption("Salve o rascunho para não perder dados ao reiniciar o sistema.")

st.markdown("---")
btn_col1, btn_col2 = st.columns([2, 1])
with btn_col1:
    rel_gerar = st.button("🚀 Gerar relatório mensal", use_container_width=True, type="primary")
with btn_col2:
    if st.button("🗑️ Limpar rascunho", use_container_width=True):
        if _pasta_rasc and _pasta_rasc.exists():
            excluir_rascunho_relatorio(_pasta_rasc)
            st.success("Rascunho excluído.")
            st.rerun()

st.markdown('</div>', unsafe_allow_html=True)


# =========================================
# AÇÕES PRINCIPAIS
# =========================================

dados = {
    "DATA_ASSINATURA": st.session_state.data_assinatura.strip(),
    "NOME_CONDOMINIO": st.session_state.nome_condominio.strip(),
    "CNPJ_CONDOMINIO": st.session_state.cnpj_condominio.strip(),
    "ENDERECO_CONDOMINIO": st.session_state.endereco_condominio.strip(),
    "NOME_SINDICO": st.session_state.nome_sindico.strip(),
    "CPF_SINDICO": st.session_state.cpf_sindico.strip(),
    "VALOR_MENSAL": valor_para_template(st.session_state.valor_mensal.strip()),
    "VALOR_ADITIVO": valor_para_template(st.session_state.valor_aditivo.strip()),
    "DATA_INICIO": st.session_state.data_inicio.strip(),
    "DATA_FIM": st.session_state.data_fim.strip(),
}

placeholders = {
    "{{DATA_ASSINATURA}}": dados["DATA_ASSINATURA"],
    "{{NOME_CONDOMINIO}}": dados["NOME_CONDOMINIO"],
    "{{CNPJ_CONDOMINIO}}": dados["CNPJ_CONDOMINIO"],
    "{{ENDERECO_CONDOMINIO}}": dados["ENDERECO_CONDOMINIO"],
    "{{NOME_SINDICO}}": dados["NOME_SINDICO"],
    "{{CPF_SINDICO}}": dados["CPF_SINDICO"],
    "{{VALOR_MENSAL}}": dados["VALOR_MENSAL"],
    "{{VALOR_ADITIVO}}": dados["VALOR_ADITIVO"],
    "{{DATA_INICIO}}": dados["DATA_INICIO"],
    "{{DATA_FIM}}": dados["DATA_FIM"],
}

# =========================================
# FUNÇÕES DE PROCESSAMENTO DE DOCUMENTOS
# =========================================

def exibir_bloco_envio(
    nome_condominio: str,
    pasta_condominio: Path,
    whatsapp_cliente: str,
    email_cliente: str,
    mensagem: str,
):
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Mensagem pronta para envio")
    st.text_area(
        "Texto",
        value=mensagem,
        height=220,
        label_visibility="collapsed",
    )

    componente_copiar(mensagem)

    col_env1, col_env2, col_env3 = st.columns(3)

    with col_env1:
        if whatsapp_cliente.strip():
            url_wa = link_whatsapp(whatsapp_cliente, mensagem)
            st.link_button("Abrir WhatsApp", url_wa, use_container_width=True)
        else:
            st.button("Abrir WhatsApp", disabled=True, use_container_width=True)

    with col_env2:
        if email_cliente.strip():
            assunto = f"Documentação contratual – {nome_condominio}"
            url_mail = link_email(email_cliente, assunto, mensagem)
            st.link_button("Abrir e-mail", url_mail, use_container_width=True)
        else:
            st.button("Abrir e-mail", disabled=True, use_container_width=True)

    with col_env3:
        if st.button("Abrir pasta deste condomínio", use_container_width=True):
            abrir_pasta_windows(pasta_condominio)

    st.markdown("</div>", unsafe_allow_html=True)


def gerar_contrato_e_aditivo():
    email_cliente = st.session_state.email_cliente.strip()
    erros = validar_para_geracao(dados, email_cliente)

    if erros:
        st.error("Corrija os campos antes de gerar os documentos:")
        for erro in erros:
            st.write(f"- {erro}")
        return

    try:
        nome_condominio = st.session_state.nome_condominio.strip()
        nome_sindico = st.session_state.nome_sindico.strip()
        whatsapp_cliente = st.session_state.whatsapp_cliente.strip()

        nome_pasta = slugify_nome(nome_condominio)
        pasta_condominio = GENERATED_DIR / nome_pasta
        pasta_condominio.mkdir(parents=True, exist_ok=True)

        salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_nome_contrato = limpar_nome_arquivo(f"Contrato_RT_{nome_condominio}_{timestamp}")
        base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")

        contrato_docx = pasta_condominio / f"{base_nome_contrato}.docx"
        contrato_pdf = pasta_condominio / f"{base_nome_contrato}.pdf"

        aditivo_docx = pasta_condominio / f"{base_nome_aditivo}.docx"
        aditivo_pdf = pasta_condominio / f"{base_nome_aditivo}.pdf"

        with st.spinner("Gerando contrato..."):
            gerar_documento(
                template_path=TEMPLATE_CONTRATO,
                output_docx=contrato_docx,
                placeholders=placeholders,
                incluir_assinaturas=True,
                nome_sindico=nome_sindico,
                nome_condominio=nome_condominio,
                cnpj_condominio=st.session_state.cnpj_condominio.strip(),
            )

        with st.spinner("Gerando aditivo..."):
            gerar_documento(
                template_path=TEMPLATE_ADITIVO,
                output_docx=aditivo_docx,
                placeholders=placeholders,
                incluir_assinaturas=True,
                nome_sindico=nome_sindico,
                nome_condominio=nome_condominio,
                cnpj_condominio=st.session_state.cnpj_condominio.strip(),
            )

        ok_contrato, erro_contrato = converter_docx_para_pdf(contrato_docx, contrato_pdf)
        ok_aditivo, erro_aditivo = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)

        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Contrato",
            arquivo_docx=contrato_docx,
            arquivo_pdf=contrato_pdf,
            pdf_gerado=ok_contrato,
            erro_pdf=erro_contrato,
            dados_utilizados=dados,
        )
        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Aditivo",
            arquivo_docx=aditivo_docx,
            arquivo_pdf=aditivo_pdf,
            pdf_gerado=ok_aditivo,
            erro_pdf=erro_aditivo,
            dados_utilizados=dados,
        )

        st.session_state.ultima_pasta_gerada = str(pasta_condominio)
        st.session_state.ultimos_docs_gerados = {
            "contrato_docx": str(contrato_docx) if contrato_docx.exists() else None,
            "contrato_pdf": str(contrato_pdf) if ok_contrato and contrato_pdf.exists() else None,
            "aditivo_docx": str(aditivo_docx) if aditivo_docx.exists() else None,
            "aditivo_pdf": str(aditivo_pdf) if ok_aditivo and aditivo_pdf.exists() else None,
        }

        st.success("Contrato e aditivo gerados com sucesso.")

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Arquivos gerados")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("**Contrato**")
            if contrato_docx.exists():
                with open(contrato_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do contrato",
                        data=f,
                        file_name=contrato_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            if ok_contrato and contrato_pdf.exists():
                with open(contrato_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do contrato",
                        data=f,
                        file_name=contrato_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do contrato não gerado. Erro: {erro_contrato}")

        with c2:
            st.markdown("**Aditivo**")
            if aditivo_docx.exists():
                with open(aditivo_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do aditivo",
                        data=f,
                        file_name=aditivo_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            if ok_aditivo and aditivo_pdf.exists():
                with open(aditivo_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do aditivo",
                        data=f,
                        file_name=aditivo_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do aditivo não gerado. Erro: {erro_aditivo}")

        st.markdown("</div>", unsafe_allow_html=True)

        mensagem = montar_mensagem_envio(
            nome_condominio=nome_condominio,
            nome_sindico=nome_sindico,
            caminho_contrato_pdf=contrato_pdf if contrato_pdf.exists() else None,
            caminho_aditivo_pdf=aditivo_pdf if aditivo_pdf.exists() else None,
        )

        exibir_bloco_envio(
            nome_condominio=nome_condominio,
            pasta_condominio=pasta_condominio,
            whatsapp_cliente=whatsapp_cliente,
            email_cliente=email_cliente,
            mensagem=mensagem,
        )

    except Exception as e:
        st.error(f"Erro na geração dos documentos: {e}")


def gerar_somente_aditivo_rapido():
    email_cliente = st.session_state.email_cliente.strip()
    erros = validar_para_geracao(dados, email_cliente)

    if erros:
        st.error("Corrija os campos antes de gerar o aditivo rápido:")
        for erro in erros:
            st.write(f"- {erro}")
        return

    try:
        nome_condominio = st.session_state.nome_condominio.strip()
        nome_sindico = st.session_state.nome_sindico.strip()
        whatsapp_cliente = st.session_state.whatsapp_cliente.strip()

        nome_pasta = slugify_nome(nome_condominio)
        pasta_condominio = GENERATED_DIR / nome_pasta
        pasta_condominio.mkdir(parents=True, exist_ok=True)

        salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")

        aditivo_docx = pasta_condominio / f"{base_nome_aditivo}.docx"
        aditivo_pdf = pasta_condominio / f"{base_nome_aditivo}.pdf"

        with st.spinner("Gerando aditivo rápido..."):
            gerar_documento(
                template_path=TEMPLATE_ADITIVO,
                output_docx=aditivo_docx,
                placeholders=placeholders,
                incluir_assinaturas=True,
                nome_sindico=nome_sindico,
                nome_condominio=nome_condominio,
                cnpj_condominio=st.session_state.cnpj_condominio.strip(),
            )

        ok_aditivo, erro_aditivo = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)

        registrar_documento_manifest(
            pasta_condominio=pasta_condominio,
            nome_condominio=nome_condominio,
            tipo="Aditivo",
            arquivo_docx=aditivo_docx,
            arquivo_pdf=aditivo_pdf,
            pdf_gerado=ok_aditivo,
            erro_pdf=erro_aditivo,
            dados_utilizados=dados,
        )

        st.session_state.ultima_pasta_gerada = str(pasta_condominio)
        st.session_state.ultimos_docs_gerados = st.session_state.get("ultimos_docs_gerados") or {}
        st.session_state.ultimos_docs_gerados.update({
            "aditivo_docx": str(aditivo_docx) if aditivo_docx.exists() else None,
            "aditivo_pdf": str(aditivo_pdf) if ok_aditivo and aditivo_pdf.exists() else None,
        })

        st.success("Aditivo rápido gerado com sucesso.")

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Arquivo gerado")

        c1, c2 = st.columns(2)

        with c1:
            if aditivo_docx.exists():
                with open(aditivo_docx, "rb") as f:
                    st.download_button(
                        "Baixar DOCX do aditivo",
                        data=f,
                        file_name=aditivo_docx.name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

        with c2:
            if ok_aditivo and aditivo_pdf.exists():
                with open(aditivo_pdf, "rb") as f:
                    st.download_button(
                        "Baixar PDF do aditivo",
                        data=f,
                        file_name=aditivo_pdf.name,
                        mime="application/pdf",
                        use_container_width=True,
                    )
            else:
                st.warning(f"PDF do aditivo não gerado. Erro: {erro_aditivo}")

        st.markdown("</div>", unsafe_allow_html=True)

        mensagem = montar_mensagem_envio(
            nome_condominio=nome_condominio,
            nome_sindico=nome_sindico,
            caminho_contrato_pdf=None,
            caminho_aditivo_pdf=aditivo_pdf if aditivo_pdf.exists() else None,
        )

        exibir_bloco_envio(
            nome_condominio=nome_condominio,
            pasta_condominio=pasta_condominio,
            whatsapp_cliente=whatsapp_cliente,
            email_cliente=email_cliente,
            mensagem=mensagem,
        )

    except Exception as e:
        st.error(f"Erro na geração do aditivo rápido: {e}")


# =========================================
# PROCESSAMENTO
# =========================================

if gerar:
    gerar_contrato_e_aditivo()

if gerar_aditivo_rapido:
    gerar_somente_aditivo_rapido()

if rel_gerar:
    ok_rel, msg_rel = gerar_relatorio_mensal()
    if ok_rel:
        st.success(msg_rel)
    else:
        st.error(msg_rel)

# =========================================
# RODAPÉ
# =========================================

st.markdown("---")
st.caption(
    f"{APP_TITLE} • {RESPONSAVEL_TÉCNICO} • {CRQ} • Ambiente prioritário: Windows"
)