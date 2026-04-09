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
                nome = str(row[0]).strip()
                pin  = str(row[1]).strip()
                conds_raw = str(row[2]).strip()
                ativo = str(row[3]).strip().lower() in ("sim", "ativo", "1", "true", "yes")
                conds = [c.strip() for c in conds_raw.split("|") if c.strip()] if conds_raw else []
                operadores.append({"nome": nome, "pin": pin, "condomínios": conds, "ativo": ativo})
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
    import re as _re
    try:
        sh = conectar_sheets()
        if sh is None:
            return False
        sheets_criar_aba_operadores()
        aba = sh.worksheet("👷 Operadores")
        todos = aba.get_all_values()
        conds_str = " | ".join(condomínios)
        ativo_str = "Sim" if ativo else "Não"
        nova_linha = [nome, pin, conds_str, ativo_str, datetime.now().strftime("%Y-%m-%d"), ""]
        # Verifica se já existe (pelo nome)
        for i, row in enumerate(todos):
            if len(row) > 0 and str(row[0]).strip().lower() == nome.lower().strip():
                aba.update(f"A{i+1}:F{i+1}", [nova_linha])
                return True
        # Insere novo
        aba.append_row(nova_linha, value_input_option="USER_ENTERED")
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
            if len(row) > 0 and str(row[0]).strip().lower() == nome.lower().strip():
                aba.delete_rows(i + 1)
                return True
        return False
    except Exception as e:
        _log_sheets_erro("sheets_deletar_operador", e)
        return False


def verificar_pin_operador(pin_digitado: str) -> dict | None:
    """Verifica PIN e retorna dados do operador, ou None se inválido."""
    operadores = sheets_listar_operadores()
    for op in operadores:
        if op["pin"] == pin_digitado.strip() and op["ativo"]:
            return op
    return None

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
            "",  # col A vazia
            id_visita,
            lancamento.get("data", ""),
            id_cliente,
            nome_condominio,
            lancamento.get("ph", ""),
            lancamento.get("cloro_livre", ""),
            lancamento.get("alcalinidade", ""),
            lancamento.get("dureza", ""),
            lancamento.get("cianurico", ""),
            "",  # foto antes
            "",  # foto depois
            "",  # foto casa máquinas
            lancamento.get("observacao", ""),
            "",  # dosagem cloro (calculado)
            "",  # dosagem bicarbonato (calculado)
            "",  # alerta pH
            "",  # alerta cloro
            "Concluída",
            lancamento.get("operador", ""),
        ]

        aba.append_row(nova_linha, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        _log_sheets_erro("sheets_salvar_lancamento_campo", e)
        return False


def sheets_salvar_cliente(nome: str, cnpj: str, endereco: str, contato: str, telefone: str,
                           vol_adulto: float = 0, vol_infantil: float = 0, vol_family: float = 0):
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
            })
        return clientes
    except Exception as e:
        _log_sheets_erro("sheets_listar_clientes_completo", e)
        return []



def sheets_editar_cliente(id_cliente: str, nome: str, cnpj: str, endereco: str,
                           contato: str, telefone: str,
                           vol_adulto: float = 0, vol_infantil: float = 0, vol_family: float = 0) -> bool:
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
                ]
                aba.update(f"A{linha_sheets}:L{linha_sheets}", [nova], value_input_option="USER_ENTERED")
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

OPERADORES_JSON = None  # será inicializado após BASE_DIR

def _get_operadores_path():
    return BASE_DIR / "_operadores.json"

def carregar_operadores() -> list[dict]:
    """Carrega lista de operadores do arquivo JSON local."""
    try:
        p = _get_operadores_path()
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
        return []
    except Exception:
        return []

def salvar_operadores(lista: list):
    """Salva lista de operadores no arquivo JSON local."""
    try:
        p = _get_operadores_path()
        p.write_text(json.dumps(lista, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def validar_pin_operador(pin: str) -> dict | None:
    """Valida PIN do operador. Retorna dict do operador ou None se inválido.
    Também aceita PIN global 5010 (acesso total)."""
    # PIN global continua funcionando — acesso total
    if pin == PIN_OPERADOR:
        return {"nome": "Operador", "pin": pin, "condomínios": ["TODOS"], "acesso_total": True}
    # Busca nos operadores do Sheets
    try:
        operadores = sheets_listar_operadores()
        for op in operadores:
            if op.get("pin","").strip() == pin.strip() and op.get("pin","").strip():
                return op
    except Exception:
        pass
    # Fallback: JSON local
    operadores_local = carregar_operadores()
    for op in operadores_local:
        if op.get("pin","") == pin:
            return op
    return None

def sheets_listar_lancamentos(nome_condominio: str) -> list[dict]:
    """Retorna lançamentos de visitas de um condomínio."""
    try:
        sh = conectar_sheets()
        if sh is None:
            return []
        aba = sh.worksheet("🔬 Visitas")
        todos = aba.get_all_values()
        lancamentos = []
        for row in todos:
            if len(row) > 4 and nome_condominio.lower() in str(row[4]).lower():
                lancamentos.append({
                    "data": row[2] if len(row) > 2 else "",
                    "ph": row[5] if len(row) > 5 else "",
                    "cloro_livre": row[6] if len(row) > 6 else "",
                    "alcalinidade": row[7] if len(row) > 7 else "",
                    "dureza": row[8] if len(row) > 8 else "",
                    "cianurico": row[9] if len(row) > 9 else "",
                    "observacao": row[13] if len(row) > 13 else "",
                    "operador": row[19] if len(row) > 19 else "",
                })
        return lancamentos
    except Exception as e:
        _log_sheets_erro("sheets_listar_lancamentos", e)
        return []

# =========================================
# OPERADORES — CONTROLE DE ACESSO
# =========================================










def filtrar_condomínios_por_operador(nome_operador: str, todos_condomínios: list[str]) -> list[str]:
    """Retorna lista de condomínios que o operador tem permissão de ver."""
    if not nome_operador.strip():
        return todos_condomínios  # sem nome → mostra todos (modo antigo)
    operadores = sheets_listar_operadores()
    for op in operadores:
        if op["nome"].lower().strip() == nome_operador.lower().strip():
            if "TODOS" in op["condomínios"]:
                return todos_condomínios
            # Filtra os condomínios permitidos
            permitidos = []
            for cond in todos_condomínios:
                for perm in op["condomínios"]:
                    if perm.lower() in cond.lower() or cond.lower() in perm.lower():
                        permitidos.append(cond)
                        break
            return permitidos if permitidos else todos_condomínios
    # Operador não cadastrado → mostra todos (retrocompatibilidade)
    return todos_condomínios



# =========================================
# MOTOR DE SUGESTÕES DE DOSAGEM
# =========================================

# Doses padrão APSP/WHO
DOSE_HIPOCLORITO_65  = 13.0   # g/m³ por ppm de CRL
DOSE_DICLORO_56      = 15.0   # g/m³ por ppm de CRL
DOSE_ACIDO_MURIATICO = 18.0   # mL/m³ por 0,1 de pH
DOSE_BICARBONATO     = 15.0   # g/m³ por 10 ppm de Alc
DOSE_CLORETO_CALCIO  = 17.0   # g/m³ por 10 ppm de Dureza
FATOR_DEMANDA        = 1.20   # +20% por demanda orgânica/UV

# Metas ideais (centro da faixa)
META_CRL        = 3.0    # ppm
META_PH         = 7.35   # centro entre 7,2-7,5
META_ALC        = 100.0  # ppm
META_DUREZA     = 225.0  # ppm
META_CYA        = 40.0   # ppm (só monitoramento)

# Faixas ideais
FAIXA_PH_MIN    = 7.2;  FAIXA_PH_MAX    = 7.8
FAIXA_CRL_MIN   = 0.5;  FAIXA_CRL_MAX   = 5.0
FAIXA_ALC_MIN   = 80;   FAIXA_ALC_MAX   = 120
FAIXA_DC_MIN    = 150;  FAIXA_DC_MAX    = 300
FAIXA_CYA_MIN   = 30;   FAIXA_CYA_MAX   = 50


def calcular_sugestoes_dosagem(ph: float | None, crl: float | None,
                                alc: float | None, dc: float | None,
                                cya: float | None, volume_m3: float) -> list[dict]:
    """
    Calcula sugestões de produtos e doses baseado nos parâmetros medidos.
    Retorna lista de dicts com: produto, quantidade, unidade, prioridade, justificativa.
    """
    sugestoes = []

    if volume_m3 <= 0:
        return sugestoes

    # ── 1. pH — SEMPRE CORRIGIR ANTES DO CLORO ───────────────────────────────
    if ph is not None:
        if ph > 7.8:
            # pH muito alto — reduzir antes de clorar
            deficit_ph = round(ph - 7.5, 1)
            dose_ml = round(deficit_ph / 0.1 * DOSE_ACIDO_MURIATICO * volume_m3 * FATOR_DEMANDA)
            sugestoes.append({
                "prioridade": 1,
                "produto": "Ácido muriático 31%",
                "quantidade": dose_ml,
                "unidade": "mL",
                "acao": "Reduzir pH",
                "justificativa": f"pH {ph:.1f} acima de 7,8 — reduzir para 7,4–7,5 antes de clorar. "
                                 f"Cloro perde >50% eficiência com pH alto.",
                "norma": "APSP / ABNT NBR 10339",
            })
        elif ph < 7.2:
            # pH baixo — hipoclorito vai elevar naturalmente
            sugestoes.append({
                "prioridade": 1,
                "produto": "Observação",
                "quantidade": 0,
                "unidade": "",
                "acao": "pH baixo — Hipoclorito de cálcio vai elevar",
                "justificativa": f"pH {ph:.1f} abaixo de 7,2 — aplicar Hipoclorito de cálcio 65% "
                                 f"(pH ~11,5) vai elevar o pH naturalmente enquanto desinfeta.",
                "norma": "APSP / WHO",
            })

    # ── 2. CLORO — produto baseado no pH ─────────────────────────────────────
    if crl is not None and crl < META_CRL:
        deficit_crl = round(META_CRL - crl, 2)

        # Seleciona produto pelo pH
        if ph is None or ph <= 7.5:
            produto_cloro = "Hipoclorito de cálcio 65%"
            fator_cloro   = DOSE_HIPOCLORITO_65
            motivo_cloro  = f"pH {ph:.1f} ≤ 7,5 — produto ideal nesta faixa" if ph else "pH não medido — usando padrão"
        else:
            produto_cloro = "Dicloro 56%"
            fator_cloro   = DOSE_DICLORO_56
            motivo_cloro  = f"pH {ph:.1f} entre 7,5–7,8 — Dicloro é mais indicado (mais ácido)"

        dose_g = round(deficit_crl * volume_m3 * fator_cloro * FATOR_DEMANDA)

        if dose_g >= 1000:
            qtd_fmt = round(dose_g / 1000, 2)
            unid = "kg"
        else:
            qtd_fmt = dose_g
            unid = "g"

        sugestoes.append({
            "prioridade": 2,
            "produto": produto_cloro,
            "quantidade": qtd_fmt,
            "unidade": unid,
            "acao": f"Elevar CRL de {crl:.1f} → {META_CRL:.1f} ppm",
            "justificativa": f"CRL {crl:.1f} ppm abaixo da meta ({META_CRL} ppm). "
                             f"Déficit: {deficit_crl} ppm × {volume_m3}m³ × {fator_cloro}g × 1,2 dem. "
                             f"| {motivo_cloro}.",
            "norma": "APSP / WHO",
        })

    elif crl is not None and crl > FAIXA_CRL_MAX:
        sugestoes.append({
            "prioridade": 2,
            "produto": "Aeração + aguardar",
            "quantidade": 0,
            "unidade": "",
            "acao": f"CRL {crl:.1f} ppm — acima de {FAIXA_CRL_MAX} ppm",
            "justificativa": "Cloro excessivo — não adicionar mais. "
                             "Aeração (agitação da água) e luz solar reduzem naturalmente.",
            "norma": "ABNT NBR 10339",
        })

    # ── 3. ALCALINIDADE ───────────────────────────────────────────────────────
    if alc is not None:
        if alc < FAIXA_ALC_MIN:
            deficit_alc = round(META_ALC - alc, 1)
            dose_g = round((deficit_alc / 10) * DOSE_BICARBONATO * volume_m3)
            qtd_fmt = round(dose_g / 1000, 2) if dose_g >= 1000 else dose_g
            unid = "kg" if dose_g >= 1000 else "g"
            sugestoes.append({
                "prioridade": 3,
                "produto": "Bicarbonato de sódio",
                "quantidade": qtd_fmt,
                "unidade": unid,
                "acao": f"Elevar alcalinidade de {alc:.0f} → {META_ALC:.0f} ppm",
                "justificativa": f"Alcalinidade {alc:.0f} ppm abaixo de {FAIXA_ALC_MIN} ppm. "
                                 f"Déficit: {deficit_alc:.0f} ppm ÷ 10 × {DOSE_BICARBONATO}g × {volume_m3}m³.",
                "norma": "WHO / ABNT NBR 10339",
            })
        elif alc > FAIXA_ALC_MAX:
            excesso_alc = round(alc - META_ALC, 1)
            dose_ml = round((excesso_alc / 0.1) * DOSE_ACIDO_MURIATICO * volume_m3 * 0.5)
            sugestoes.append({
                "prioridade": 3,
                "produto": "Ácido muriático 31%",
                "quantidade": dose_ml,
                "unidade": "mL",
                "acao": f"Reduzir alcalinidade de {alc:.0f} → {META_ALC:.0f} ppm",
                "justificativa": f"Alcalinidade {alc:.0f} ppm acima de {FAIXA_ALC_MAX} ppm. "
                                 "Aplicar ácido muriático com bomba desligada.",
                "norma": "APSP",
            })

    # ── 4. DUREZA ─────────────────────────────────────────────────────────────
    if dc is not None:
        if dc < FAIXA_DC_MIN:
            deficit_dc = round(META_DUREZA - dc, 1)
            dose_g = round((deficit_dc / 10) * DOSE_CLORETO_CALCIO * volume_m3)
            qtd_fmt = round(dose_g / 1000, 2) if dose_g >= 1000 else dose_g
            unid = "kg" if dose_g >= 1000 else "g"
            sugestoes.append({
                "prioridade": 4,
                "produto": "Cloreto de cálcio",
                "quantidade": qtd_fmt,
                "unidade": unid,
                "acao": f"Elevar dureza de {dc:.0f} → {META_DUREZA:.0f} ppm",
                "justificativa": f"Dureza {dc:.0f} ppm abaixo de {FAIXA_DC_MIN} ppm. "
                                 "Água agressiva corrói equipamentos e pisos.",
                "norma": "APSP / WHO",
            })
        elif dc > FAIXA_DC_MAX:
            sugestoes.append({
                "prioridade": 4,
                "produto": "Troca parcial de água",
                "quantidade": round(volume_m3 * 0.2),
                "unidade": "m³",
                "acao": f"Reduzir dureza de {dc:.0f} ppm",
                "justificativa": f"Dureza {dc:.0f} ppm acima de {FAIXA_DC_MAX} ppm. "
                                 "Trocar ~20% da água e reequilibrar.",
                "norma": "APSP",
            })

    # ── 5. CYA — só monitoramento ─────────────────────────────────────────────
    if cya is not None:
        if cya > FAIXA_CYA_MAX:
            sugestoes.append({
                "prioridade": 5,
                "produto": "Troca parcial de água",
                "quantidade": round(volume_m3 * 0.3),
                "unidade": "m³",
                "acao": f"CYA {cya:.0f} ppm acima do limite",
                "justificativa": f"CYA {cya:.0f} ppm acima de {FAIXA_CYA_MAX} ppm. "
                                 "CYA alto inibe o cloro (efeito bloqueio). Trocar ~30% da água.",
                "norma": "WHO",
            })
        elif cya < FAIXA_CYA_MIN and cya > 0:
            sugestoes.append({
                "prioridade": 5,
                "produto": "Monitorar CYA",
                "quantidade": 0,
                "unidade": "",
                "acao": f"CYA {cya:.0f} ppm — abaixo do ideal",
                "justificativa": f"CYA {cya:.0f} ppm abaixo de {FAIXA_CYA_MIN} ppm. "
                                 "Sem ácido cianúrico puro no estoque — monitorar.",
                "norma": "APSP",
            })

    # Ordena por prioridade
    sugestoes.sort(key=lambda x: x["prioridade"])
    return sugestoes


def exibir_sugestoes_dosagem(sugestoes: list[dict]):
    """Exibe as sugestões de dosagem formatadas no Streamlit."""
    if not sugestoes:
        st.success("✅ Todos os parâmetros dentro da faixa ideal. Nenhuma correção necessária.")
        return

    st.markdown("**💊 Sugestões de correção (APSP/WHO):**")
    for s in sugestoes:
        prod = s["produto"]
        qtd  = s["quantidade"]
        unid = s["unidade"]
        acao = s["acao"]
        just = s["justificativa"]
        prio = s["prioridade"]

        if prio == 1:
            icon = "🔴"
        elif prio == 2:
            icon = "🟡"
        else:
            icon = "🔵"

        if qtd and qtd > 0:
            st.markdown(f"{icon} **{prod}** — **{qtd} {unid}** → _{acao}_")
        else:
            st.markdown(f"{icon} **{prod}** → _{acao}_")

        with st.expander("ℹ️ Detalhes técnicos", expanded=False):
            st.caption(f"📐 {just}")
            st.caption(f"📚 Norma: {s.get('norma','')}")

# =========================================
# CONFIGURAÇÃO GERAL
# =========================================

APP_TITLE = "Aqua Gestão – Controle Técnico de Piscinas"
RESPONSAVEL_TÉCNICO = "Thyago Fernando da Silveira"
RESPONSAVEL_TECNICO_ASSINATURA = "Thyago Fernando da Silveira | CRQ 024025748 | Técnico em Química"
CRQ = "CRQ-MG 2ª Região | CRQ 024025748"
CRQ_NUMERO = "024025748"
QUALIFICACAO_RT = "Técnico em Química"
CERTIFICACOES_RT = "NR-26 e NR-6"
EMPRESA_RT = "Aqua Gestão – Controle Técnico de Piscinas"

# ── Dados da Bem Star Piscinas ──────────────────────────────────────────────
EMPRESA_BEM_STAR     = "Bem Star Piscinas"
CNPJ_BEM_STAR        = "26.799.958/0001-88"
ENDERECO_BEM_STAR    = "Avenida Getúlio Vargas, 4411 — CEP 38.412-316 — Uberlândia/MG"
LOGO_BEM_STAR_CANDIDATOS = [
    BASE_DIR / "bem_star_logo.png",
    BASE_DIR / "bem_star_logo.jpg",
    BASE_DIR / "assets" / "bem_star_logo.png",
]

# ── Texto RT para relatório sem RT ──────────────────────────────────────────
TEXTO_RT_SEM_RT = """SOBRE RESPONSABILIDADE TÉCNICA (RT)

Este relatório foi elaborado pela Bem Star Piscinas como registro técnico das análises e dosagens realizadas durante a visita de manutenção.

A Responsabilidade Técnica (RT) é um serviço de supervisão especializada, regulamentado pela Resolução CFQ nº 332/2025 — publicada pelo Conselho Federal de Química em 24 de junho de 2025 — e complementada pela Resolução CFQ nº 345/2026, que tornam obrigatória a Anotação de Responsabilidade Técnica (ART) para o tratamento químico e controle de qualidade da água de piscinas de uso público e coletivo, abrangendo condomínios residenciais, clubes, academias, hotéis e escolas.

A RT consiste na supervisão e assinatura de profissional habilitado e registrado no Conselho Regional de Química (CRQ), garantindo que todos os procedimentos estejam em conformidade com os padrões técnicos e sanitários vigentes, incluindo a ABNT NBR 10339. A ART deve ser emitida anualmente e afixada em local visível, podendo o CRQ realizar fiscalizações preventivas e, em caso de irregularidade, acionar a Vigilância Sanitária municipal.

A Aqua Gestão — empresa parceira especializada em Responsabilidade Técnica — oferece o serviço completo de RT com profissional habilitado pelo CRQ-MG, emissão de laudos mensais, ART anual e total suporte documental, garantindo segurança jurídica e conformidade normativa ao seu condomínio.

Saiba mais sobre o serviço de RT:
Thyago Fernando da Silveira | CRQ-MG 2ª Região | CRQ 024025748
Aqua Gestão – Controle Técnico de Piscinas"""

BASE_DIR = Path(__file__).resolve().parent
GENERATED_DIR = BASE_DIR / "generated"
TEMPLATE_CONTRATO = BASE_DIR / "template.docx"
TEMPLATE_ADITIVO = BASE_DIR / "aditivo.docx"
TEMPLATE_RELATORIO = BASE_DIR / "relatorio_mensal.docx"
DADOS_JSON_NAME = "dados_condominio.json"
MANIFEST_JSON_NAME = "manifest.json"
ANALISES_PADRAO = 9
ANALISES_MAX_SUGERIDO = 40

ASSINATURA_RT_CANDIDATOS = [
    BASE_DIR / "assinatura_rt.png",
    BASE_DIR / "assinatura_rt.jpg",
    BASE_DIR / "assinatura_rt.jpeg",
    BASE_DIR / "assets" / "assinatura_rt.png",
    BASE_DIR / "images" / "assinatura_rt.png",
]

LOGO_CANDIDATOS = [
    BASE_DIR / "aqua_gestao_logo.png",
    BASE_DIR / "aqua_gestao_logo.jpg",
    BASE_DIR / "aqua_gestao_logo.jpeg",
    BASE_DIR / "logo.png",
    BASE_DIR / "logo.jpg",
    BASE_DIR / "logo.jpeg",
    BASE_DIR / "Aqua Gestão Logo.png",
    BASE_DIR / "Aqua_Gestao_Logo.png",
    BASE_DIR / "assets" / "aqua_gestao_logo.png",
    BASE_DIR / "assets" / "logo.png",
    BASE_DIR / "images" / "aqua_gestao_logo.png",
    BASE_DIR / "images" / "logo.png",
]

GENERATED_DIR.mkdir(exist_ok=True)

st.set_page_config(
    page_title="Aqua Gestão RT",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================
# ESTILO VISUAL
# =========================================

st.markdown(
    """
    <style>
        .main .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1360px;
        }

        .top-card {
            border: 1px solid rgba(20, 85, 160, 0.18);
            border-radius: 18px;
            padding: 18px 22px;
            background: linear-gradient(135deg, #ffffff 0%, #f6fbff 100%);
            box-shadow: 0 10px 25px rgba(10, 50, 100, 0.07);
            margin-bottom: 18px;
        }

        .top-title {
            font-size: 1.7rem;
            font-weight: 700;
            color: #0d3d75;
            margin: 0;
            line-height: 1.15;
        }

        .top-subtitle {
            font-size: 0.95rem;
            color: #3b5d85;
            margin-top: 8px;
        }

        .info-badge {
            display: inline-block;
            padding: 7px 12px;
            border-radius: 999px;
            background: #edf5ff;
            color: #134b8a;
            border: 1px solid #d3e6ff;
            font-size: 0.88rem;
            margin-right: 8px;
            margin-top: 8px;
        }

        .section-card {
            border: 1px solid rgba(20, 85, 160, 0.16);
            border-radius: 18px;
            padding: 18px;
            background: #ffffff;
            box-shadow: 0 8px 20px rgba(10, 50, 100, 0.05);
            margin-bottom: 16px;
        }

        .history-meta {
            font-size: 0.84rem;
            color: #6a7d92;
            margin-top: 2px;
        }

        .confirm-box {
            border: 1px solid rgba(220, 80, 80, 0.35);
            border-radius: 12px;
            padding: 10px;
            margin-top: 8px;
            margin-bottom: 10px;
            background: rgba(220, 80, 80, 0.08);
        }

        .quick-mode-box {
            border: 1px solid rgba(20, 85, 160, 0.18);
            border-radius: 14px;
            padding: 12px 14px;
            background: #f6fbff;
            margin-bottom: 14px;
        }

        .alert-vigente {
            border: 1px solid rgba(40, 140, 80, 0.25);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(40, 140, 80, 0.08);
            margin-bottom: 14px;
        }

        .alert-vencendo {
            border: 1px solid rgba(220, 150, 20, 0.35);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(255, 190, 40, 0.12);
            margin-bottom: 14px;
        }

        .alert-vencido {
            border: 1px solid rgba(220, 60, 60, 0.35);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(220, 60, 60, 0.10);
            margin-bottom: 14px;
        }

        .status-badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 0.75rem;
            margin-top: 6px;
            border: 1px solid rgba(0,0,0,0.08);
        }

        .status-vigente { background: #ecfff3; color: #1d7a43; }
        .status-vencendo { background: #fff8e8; color: #9c6200; }
        .status-vencido { background: #fff0f0; color: #b42318; }
        .status-indefinido { background: #f4f6f8; color: #52606d; }

        .venc-row {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 16px;
            padding: 14px;
            background: linear-gradient(135deg, #ffffff 0%, #fbfdff 100%);
            margin-bottom: 12px;
        }

        .venc-nome {
            font-size: 1rem;
            font-weight: 700;
            color: #153f73;
            margin-bottom: 4px;
        }

        .venc-meta {
            font-size: 0.88rem;
            color: #5d7288;
            margin-bottom: 3px;
        }

        .venc-empty {
            border: 1px dashed rgba(20, 85, 160, 0.20);
            border-radius: 14px;
            padding: 14px;
            background: #f9fcff;
            color: #59708b;
        }

        .legacy-note {
            border: 1px solid rgba(120, 120, 120, 0.18);
            border-radius: 12px;
            padding: 10px 12px;
            background: #fafbfd;
            color: #5e6f82;
            font-size: 0.87rem;
            margin-top: 8px;
            margin-bottom: 8px;
        }

        .dash-mini {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 16px;
            padding: 14px;
            background: #fbfdff;
            min-height: 120px;
        }

        .dash-title {
            font-size: 0.88rem;
            color: #5f7590;
            margin-bottom: 6px;
        }

        .dash-value {
            font-size: 1.55rem;
            font-weight: 700;
            color: #103f78;
            line-height: 1.1;
        }

        .dash-sub {
            font-size: 0.82rem;
            color: #6d8197;
            margin-top: 8px;
        }

        .docs-note {
            border: 1px solid rgba(20, 85, 160, 0.12);
            border-radius: 12px;
            padding: 10px;
            background: #fbfdff;
            color: #5c7188;
            font-size: 0.84rem;
            margin-top: 8px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================
# FUNÇÕES AUXILIARES GERAIS
# =========================================

def encontrar_assinatura_rt() -> Path | None:
    for caminho in ASSINATURA_RT_CANDIDATOS:
        if caminho.exists() and caminho.is_file():
            return caminho
    return None


def preparar_assinatura_rt_para_relatorio() -> Path | None:
    assinatura = encontrar_assinatura_rt()
    if not assinatura:
        return None

    destino = GENERATED_DIR / "_assinatura_rt_relatorio.png"
    try:
        with Image.open(assinatura) as img:
            img = img.convert("RGBA")
            fundo = Image.new("RGBA", img.size, (255, 255, 255, 255))
            fundo.alpha_composite(img)
            fundo = fundo.convert("RGB")
            fundo = ImageOps.expand(fundo, border=18, fill="white")
            fundo.save(destino, format="PNG")
        return destino
    except Exception:
        return assinatura


def encontrar_logo() -> Path | None:
    for caminho in LOGO_CANDIDATOS:
        if caminho.exists() and caminho.is_file():
            return caminho

    for extensao in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
        for pasta in [BASE_DIR, BASE_DIR / "assets", BASE_DIR / "images"]:
            if pasta.exists():
                encontrados = list(pasta.glob(extensao))
                for arq in encontrados:
                    if "logo" in arq.name.lower():
                        return arq
    return None


def slugify_nome(texto: str) -> str:
    texto = (texto or "").strip()
    texto = re.sub(r"[^\w\s-]", "", texto, flags=re.UNICODE)
    texto = re.sub(r"\s+", "_", texto)
    return texto[:120] if texto else "condominio"


def humanizar_nome_pasta(texto: str) -> str:
    texto = (texto or "").strip()
    texto = texto.replace("_", " ").replace("-", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def limpar_nome_arquivo(texto: str) -> str:
    texto = re.sub(r'[<>:"/\\\\|?*]+', "", texto)
    texto = re.sub(r"\s+", "_", texto.strip())
    return texto[:150]


def hoje_br() -> str:
    return date.today().strftime("%d/%m/%Y")


def apenas_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def formatar_data_hora_arquivo(ts: float) -> str:
    dt = datetime.fromtimestamp(ts)
    return dt.strftime("%d/%m/%Y %H:%M")


def classificar_arquivo(nome_arquivo: str) -> tuple[str, str]:
    nome_lower = nome_arquivo.lower()

    if "contrato" in nome_lower:
        tipo_doc = "Contrato"
    elif "aditivo" in nome_lower:
        tipo_doc = "Aditivo"
    elif "relatorio" in nome_lower:
        tipo_doc = "Relatório"
    else:
        tipo_doc = "Documento"

    if nome_lower.endswith(".pdf"):
        tipo_ext = "PDF"
    elif nome_lower.endswith(".docx"):
        tipo_ext = "DOCX"
    else:
        tipo_ext = "Arquivo"

    return tipo_doc, tipo_ext


def chave_segura(texto: str) -> str:
    return re.sub(r"[^a-zA-Z0-9_]+", "_", texto)


def parse_data_br(texto: str):
    try:
        return datetime.strptime((texto or "").strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def formatar_data_br(dt: date) -> str:
    return dt.strftime("%d/%m/%Y")


def adicionar_um_ano(dt: date) -> date:
    try:
        return dt.replace(year=dt.year + 1)
    except ValueError:
        return dt.replace(month=2, day=28, year=dt.year + 1)


def calcular_renovacao_anual(data_fim_texto: str):
    fim_atual = parse_data_br(data_fim_texto)
    if not fim_atual:
        return None, None

    novo_inicio = fim_atual + timedelta(days=1)
    novo_fim = adicionar_um_ano(novo_inicio) - timedelta(days=1)
    return novo_inicio, novo_fim


def status_vencimento(data_fim_texto: str, alerta_dias: int = 30):
    fim = parse_data_br(data_fim_texto)
    if not fim:
        return {
            "codigo": "indefinido",
            "rotulo": "Sem vigência válida",
            "mensagem": "Data final ausente ou inválida.",
            "dias": None,
            "css": "status-indefinido",
        }

    hoje = date.today()
    dias = (fim - hoje).days

    if dias < 0:
        return {
            "codigo": "vencido",
            "rotulo": "Vencido",
            "mensagem": f"Contrato vencido há {abs(dias)} dia(s).",
            "dias": dias,
            "css": "status-vencido",
        }

    if dias <= alerta_dias:
        return {
            "codigo": "vencendo",
            "rotulo": "Vence em breve",
            "mensagem": f"Contrato vence em {dias} dia(s).",
            "dias": dias,
            "css": "status-vencendo",
        }

    return {
        "codigo": "vigente",
        "rotulo": "Vigente",
        "mensagem": f"Contrato vigente. Restam {dias} dia(s) para o vencimento.",
        "dias": dias,
        "css": "status-vigente",
    }


def texto_dias_restantes(status: dict) -> str:
    dias = status.get("dias")
    if dias is None:
        return "Dias restantes: não disponível"
    if dias < 0:
        return f"Atrasado há {abs(dias)} dia(s)"
    return f"Restam {dias} dia(s)"


def sistema_e_windows() -> bool:
    return platform.system().lower().startswith("win")


def diagnostico_sistema() -> dict:
    return {
        "template_contrato_ok": TEMPLATE_CONTRATO.exists(),
        "template_aditivo_ok": TEMPLATE_ADITIVO.exists(),
        "template_relatorio_ok": TEMPLATE_RELATORIO.exists(),
        "generated_ok": GENERATED_DIR.exists(),
        "logo_ok": encontrar_logo() is not None,
        "windows_ok": sistema_e_windows(),
    }

# =========================================
# MÁSCARAS / FORMATAÇÃO
# =========================================

def formatar_cpf(texto: str) -> str:
    dig = apenas_digitos(texto)[:11]
    if len(dig) <= 3:
        return dig
    if len(dig) <= 6:
        return f"{dig[:3]}.{dig[3:]}"
    if len(dig) <= 9:
        return f"{dig[:3]}.{dig[3:6]}.{dig[6:]}"
    return f"{dig[:3]}.{dig[3:6]}.{dig[6:9]}-{dig[9:]}"


def formatar_cnpj(texto: str) -> str:
    dig = apenas_digitos(texto)[:14]
    if len(dig) <= 2:
        return dig
    if len(dig) <= 5:
        return f"{dig[:2]}.{dig[2:]}"
    if len(dig) <= 8:
        return f"{dig[:2]}.{dig[2:5]}.{dig[5:]}"
    if len(dig) <= 12:
        return f"{dig[:2]}.{dig[2:5]}.{dig[5:8]}/{dig[8:]}"
    return f"{dig[:2]}.{dig[2:5]}.{dig[5:8]}/{dig[8:12]}-{dig[12:]}"


def formatar_telefone(texto: str) -> str:
    dig = apenas_digitos(texto)

    if dig.startswith("55") and len(dig) > 11:
        dig = dig[2:]

    dig = dig[:11]

    if len(dig) <= 2:
        return dig
    if len(dig) <= 6:
        return f"({dig[:2]}) {dig[2:]}"
    if len(dig) <= 10:
        return f"({dig[:2]}) {dig[2:6]}-{dig[6:]}"
    return f"({dig[:2]}) {dig[2:7]}-{dig[7:]}"


def formatar_data_digitada(texto: str) -> str:
    dig = apenas_digitos(texto)[:8]
    if len(dig) <= 2:
        return dig
    if len(dig) <= 4:
        return f"{dig[:2]}/{dig[2:]}"
    return f"{dig[:2]}/{dig[2:4]}/{dig[4:]}"


def moeda_br_sem_prefixo(texto: str) -> str:
    if not texto:
        return ""

    dig = apenas_digitos(str(texto))
    if not dig:
        return ""

    if len(dig) == 1:
        valor = float(f"0.0{dig}")
    elif len(dig) == 2:
        valor = float(f"0.{dig}")
    else:
        valor = float(f"{dig[:-2]}.{dig[-2:]}")

    inteiro = int(valor)
    centavos = int(round((valor - inteiro) * 100))
    inteiro_fmt = f"{inteiro:,}".replace(",", ".")
    return f"{inteiro_fmt},{centavos:02d}"


def moeda_br(texto: str) -> str:
    fmt = moeda_br_sem_prefixo(texto)
    return f"R$ {fmt}" if fmt else ""


def valor_para_template(texto: str) -> str:
    texto = (texto or "").strip()
    if texto.startswith("R$"):
        return texto
    fmt = moeda_br_sem_prefixo(texto)
    return f"R$ {fmt}" if fmt else ""


def on_change_nome_condominio():
    """Salva automaticamente o JSON quando o usuário sai do campo nome_condominio."""
    nome = (st.session_state.get("nome_condominio") or "").strip()
    if nome:
        pasta = GENERATED_DIR / slugify_nome(nome)
        pasta.mkdir(parents=True, exist_ok=True)
        salvar_dados_condominio(pasta, salvar_snapshot_formulario())


def on_change_cpf():
    st.session_state.cpf_sindico = formatar_cpf(st.session_state.get("cpf_sindico", ""))


def on_change_cnpj():
    st.session_state.cnpj_condominio = formatar_cnpj(st.session_state.get("cnpj_condominio", ""))


def on_change_whatsapp():
    st.session_state.whatsapp_cliente = formatar_telefone(st.session_state.get("whatsapp_cliente", ""))


def on_change_valor_mensal():
    st.session_state.valor_mensal = moeda_br(st.session_state.get("valor_mensal", ""))


def on_change_valor_aditivo():
    st.session_state.valor_aditivo = moeda_br(st.session_state.get("valor_aditivo", ""))


def on_change_data_inicio():
    st.session_state.data_inicio = formatar_data_digitada(st.session_state.get("data_inicio", ""))


def on_change_data_fim():
    st.session_state.data_fim = formatar_data_digitada(st.session_state.get("data_fim", ""))


def on_change_data_assinatura():
    st.session_state.data_assinatura = formatar_data_digitada(st.session_state.get("data_assinatura", ""))


def on_change_rel_documento_representante():
    atual = st.session_state.get("rel_cpf_cnpj_representante", "")
    dig = apenas_digitos(atual)
    if len(dig) <= 11:
        st.session_state.rel_cpf_cnpj_representante = formatar_cpf(atual)
    else:
        st.session_state.rel_cpf_cnpj_representante = formatar_cnpj(atual)


# =========================================
# VALIDAÇÕES REAIS
# =========================================

def validar_cpf(cpf: str) -> bool:
    cpf = apenas_digitos(cpf)

    if len(cpf) != 11:
        return False
    if cpf == cpf[0] * 11:
        return False

    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1
    if dig1 != int(cpf[9]):
        return False

    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2
    return dig2 == int(cpf[10])


def validar_cnpj(cnpj: str) -> bool:
    cnpj = apenas_digitos(cnpj)

    if len(cnpj) != 14:
        return False
    if cnpj == cnpj[0] * 14:
        return False

    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma1 = sum(int(cnpj[i]) * pesos1[i] for i in range(12))
    resto1 = soma1 % 11
    dig1 = 0 if resto1 < 2 else 11 - resto1
    if dig1 != int(cnpj[12]):
        return False

    pesos2 = [6] + pesos1
    soma2 = sum(int(cnpj[i]) * pesos2[i] for i in range(13))
    resto2 = soma2 % 11
    dig2 = 0 if resto2 < 2 else 11 - resto2
    return dig2 == int(cnpj[13])


def validar_data_br(texto: str) -> bool:
    try:
        datetime.strptime(texto.strip(), "%d/%m/%Y")
        return True
    except Exception:
        return False


def validar_email(email: str) -> bool:
    email = (email or "").strip()
    if not email:
        return True
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return re.match(padrao, email) is not None


def validar_campos_formato(dados: dict, email_cliente: str) -> list[str]:
    erros = []

    if not validar_cpf(dados["CPF_SINDICO"]):
        erros.append("CPF do síndico/representante inválido.")
    if not validar_cnpj(dados["CNPJ_CONDOMINIO"]):
        erros.append("CNPJ do condomínio inválido.")
    if not validar_data_br(dados["DATA_ASSINATURA"]):
        erros.append("Data de assinatura inválida. Use o formato dd/mm/aaaa.")
    if not validar_data_br(dados["DATA_INICIO"]):
        erros.append("Data de início inválida. Use o formato dd/mm/aaaa.")
    if not validar_data_br(dados["DATA_FIM"]):
        erros.append("Data de fim inválida. Use o formato dd/mm/aaaa.")
    if email_cliente.strip() and not validar_email(email_cliente):
        erros.append("E-mail inválido.")

    return erros


# =========================================
# SNAPSHOT / MANIFEST / CADASTRO
# =========================================

def salvar_snapshot_formulario() -> dict:
    return {
        "nome_condominio": (st.session_state.get("nome_condominio") or "").strip(),
        "cnpj_condominio": (st.session_state.get("cnpj_condominio") or "").strip(),
        "endereco_condominio": (st.session_state.get("endereco_condominio") or "").strip(),
        "nome_sindico": (st.session_state.get("nome_sindico") or "").strip(),
        "cpf_sindico": (st.session_state.get("cpf_sindico") or "").strip(),
        "valor_mensal": (st.session_state.get("valor_mensal") or "").strip(),
        "valor_aditivo": (st.session_state.get("valor_aditivo") or "").strip(),
        "data_inicio": (st.session_state.get("data_inicio") or "").strip(),
        "data_fim": (st.session_state.get("data_fim") or "").strip(),
        "data_assinatura": (st.session_state.get("data_assinatura") or "").strip(),
        "whatsapp_cliente": (st.session_state.get("whatsapp_cliente") or "").strip(),
        "email_cliente": (st.session_state.get("email_cliente") or "").strip(),
        "observacoes_internas": (st.session_state.get("observacoes_internas") or "").strip(),
        "rel_art_status": (st.session_state.get("rel_art_status") or "Emitida").strip(),
        "rel_art_numero": (st.session_state.get("rel_art_numero") or "").strip(),
        "rel_art_inicio": (st.session_state.get("rel_art_inicio") or "").strip(),
        "rel_art_fim": (st.session_state.get("rel_art_fim") or "").strip(),
        "dosagens_ultimas": obter_dosagens_ultimas_relatorio(),
        "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }


def obter_dosagens_ultimas_relatorio() -> list[dict]:
    itens = []
    for i in range(7):
        itens.append({
            "produto": (st.session_state.get(f"rel_dos_produto_{i}") or "").strip(),
            "fabricante_lote": (st.session_state.get(f"rel_dos_lote_{i}") or "").strip(),
            "quantidade": (st.session_state.get(f"rel_dos_qtd_{i}") or "").strip(),
            "unidade": (st.session_state.get(f"rel_dos_un_{i}") or "").strip(),
            "finalidade": (st.session_state.get(f"rel_dos_finalidade_{i}") or "").strip(),
        })
    return itens


def aplicar_dosagens_ultimas_no_relatorio(dados_salvos: dict | None):
    dosagens = []
    if isinstance(dados_salvos, dict):
        dosagens = dados_salvos.get("dosagens_ultimas") or []

    for i in range(7):
        item = dosagens[i] if i < len(dosagens) and isinstance(dosagens[i], dict) else {}
        st.session_state[f"rel_dos_produto_{i}"] = (item.get("produto") or "").strip()
        st.session_state[f"rel_dos_lote_{i}"] = (item.get("fabricante_lote") or "").strip()
        st.session_state[f"rel_dos_qtd_{i}"] = (item.get("quantidade") or "").strip()
        st.session_state[f"rel_dos_un_{i}"] = (item.get("unidade") or "").strip()
        st.session_state[f"rel_dos_finalidade_{i}"] = (item.get("finalidade") or "").strip()


def obter_snapshot_relatorio_independente() -> dict:
    return {
        "nome_condominio": (st.session_state.get("rel_nome_condominio") or "").strip(),
        "cnpj_condominio": (st.session_state.get("rel_cnpj_condominio") or "").strip(),
        "endereco_condominio": (st.session_state.get("rel_endereco_condominio") or "").strip(),
        "nome_sindico": (st.session_state.get("rel_representante") or "").strip(),
        "cpf_sindico": (st.session_state.get("rel_cpf_cnpj_representante") or "").strip(),
        "rel_art_status": (st.session_state.get("rel_art_status") or "Emitida").strip(),
        "rel_art_numero": (st.session_state.get("rel_art_numero") or "").strip(),
        "rel_art_inicio": (st.session_state.get("rel_art_inicio") or "").strip(),
        "rel_art_fim": (st.session_state.get("rel_art_fim") or "").strip(),
        "ultima_origem_relatorio": (st.session_state.get("rel_tipo_atendimento") or "").strip(),
        "dosagens_ultimas": obter_dosagens_ultimas_relatorio(),
        "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }


def sincronizar_relatorio_com_cadastro():
    if not (st.session_state.get("rel_nome_condominio") or "").strip():
        st.session_state.rel_nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    if not (st.session_state.get("rel_cnpj_condominio") or "").strip():
        st.session_state.rel_cnpj_condominio = (st.session_state.get("cnpj_condominio") or "").strip()
    if not (st.session_state.get("rel_endereco_condominio") or "").strip():
        st.session_state.rel_endereco_condominio = (st.session_state.get("endereco_condominio") or "").strip()
    if not (st.session_state.get("rel_representante") or "").strip():
        st.session_state.rel_representante = (st.session_state.get("nome_sindico") or "").strip()
    if not (st.session_state.get("rel_cpf_cnpj_representante") or "").strip():
        cpf = (st.session_state.get("cpf_sindico") or "").strip()
        cnpj = (st.session_state.get("cnpj_condominio") or "").strip()
        st.session_state.rel_cpf_cnpj_representante = cpf or cnpj


def carregar_dados_cadastro_no_relatorio():
    st.session_state.rel_nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    st.session_state.rel_cnpj_condominio = (st.session_state.get("cnpj_condominio") or "").strip()
    st.session_state.rel_endereco_condominio = (st.session_state.get("endereco_condominio") or "").strip()
    st.session_state.rel_representante = (st.session_state.get("nome_sindico") or "").strip()
    st.session_state.rel_cpf_cnpj_representante = (st.session_state.get("cpf_sindico") or "").strip() or (st.session_state.get("cnpj_condominio") or "").strip()
    aplicar_dosagens_ultimas_no_relatorio(salvar_snapshot_formulario())


def salvar_relatorio_no_cadastro_principal():
    st.session_state.nome_condominio = (st.session_state.get("rel_nome_condominio") or "").strip()
    st.session_state.cnpj_condominio = (st.session_state.get("rel_cnpj_condominio") or "").strip()
    st.session_state.endereco_condominio = (st.session_state.get("rel_endereco_condominio") or "").strip()
    st.session_state.nome_sindico = (st.session_state.get("rel_representante") or "").strip()
    doc = (st.session_state.get("rel_cpf_cnpj_representante") or "").strip()
    dig = apenas_digitos(doc)
    if len(dig) == 11:
        st.session_state.cpf_sindico = formatar_cpf(doc)
    elif len(dig) == 14:
        st.session_state.cnpj_condominio = formatar_cnpj(doc)


def validar_campos_obrigatorios(dados: dict) -> list[str]:
    campos = {
        "Nome do condomínio": dados.get("NOME_CONDOMINIO", ""),
        "CNPJ do condomínio": dados.get("CNPJ_CONDOMINIO", ""),
        "Endereço do condomínio": dados.get("ENDERECO_CONDOMINIO", ""),
        "Nome do síndico / representante": dados.get("NOME_SINDICO", ""),
        "CPF do síndico / representante": dados.get("CPF_SINDICO", ""),
        "Valor mensal": dados.get("VALOR_MENSAL", ""),
        "Data de início": dados.get("DATA_INICIO", ""),
        "Data de fim": dados.get("DATA_FIM", ""),
        "Data de assinatura": dados.get("DATA_ASSINATURA", ""),
    }
    faltando = [nome for nome, valor in campos.items() if not (valor or "").strip()]
    return faltando


def carregar_manifest_condominio(pasta_condominio: Path) -> dict:
    caminho = pasta_condominio / MANIFEST_JSON_NAME
    if caminho.exists():
        try:
            return json.loads(caminho.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"condominio": pasta_condominio.name, "documentos": []}


def salvar_manifest_condominio(pasta_condominio: Path, manifest: dict):
    caminho = pasta_condominio / MANIFEST_JSON_NAME
    caminho.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def registrar_documento_manifest(pasta_condominio: Path, nome_condominio: str, tipo: str, arquivo_docx: Path | None, arquivo_pdf: Path | None, pdf_gerado: bool, erro_pdf: str | None, dados_utilizados: dict | None = None, extras: dict | None = None):
    manifest = carregar_manifest_condominio(pasta_condominio)
    documento = {
        "tipo": tipo,
        "nome_condominio": nome_condominio,
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "arquivo_docx": arquivo_docx.name if arquivo_docx and arquivo_docx.exists() else None,
        "arquivo_pdf": arquivo_pdf.name if arquivo_pdf and arquivo_pdf.exists() else None,
        "pdf_gerado": bool(pdf_gerado),
        "erro_pdf": erro_pdf,
        "dados_utilizados": dados_utilizados or {},
        "extras": extras or {},
    }
    manifest.setdefault("documentos", []).append(documento)
    manifest["ultimo_update"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    salvar_manifest_condominio(pasta_condominio, manifest)


def obter_ultimo_documento_manifest(pasta_condominio: Path, tipo: str) -> dict | None:
    manifest = carregar_manifest_condominio(pasta_condominio)
    docs = [d for d in manifest.get("documentos", []) if d.get("tipo") == tipo]
    return docs[-1] if docs else None


# =========================================
# PERSISTÊNCIA DE DADOS DO CONDOMÍNIO
# =========================================

def salvar_dados_condominio(pasta_condominio: Path, dados_para_salvar: dict):
    caminho_json = pasta_condominio / DADOS_JSON_NAME
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(dados_para_salvar, f, ensure_ascii=False, indent=2)


def carregar_dados_condominio(pasta_condominio: Path) -> dict | None:
    caminho_json = pasta_condominio / DADOS_JSON_NAME
    if not caminho_json.exists():
        return None
    try:
        with open(caminho_json, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def aplicar_dados_no_formulario(dados_salvos: dict):
    st.session_state.nome_condominio = dados_salvos.get("nome_condominio", "")
    st.session_state.cnpj_condominio = dados_salvos.get("cnpj_condominio", "")
    st.session_state.endereco_condominio = dados_salvos.get("endereco_condominio", "")
    st.session_state.nome_sindico = dados_salvos.get("nome_sindico", "")
    st.session_state.cpf_sindico = dados_salvos.get("cpf_sindico", "")
    st.session_state.valor_mensal = dados_salvos.get("valor_mensal", "")
    st.session_state.valor_aditivo = dados_salvos.get("valor_aditivo", "")
    st.session_state.data_inicio = dados_salvos.get("data_inicio", "")
    st.session_state.data_fim = dados_salvos.get("data_fim", "")
    st.session_state.data_assinatura = dados_salvos.get("data_assinatura", hoje_br())
    st.session_state.whatsapp_cliente = dados_salvos.get("whatsapp_cliente", "")
    st.session_state.email_cliente = dados_salvos.get("email_cliente", "")
    st.session_state.observacoes_internas = dados_salvos.get("observacoes_internas", "")
    st.session_state.rel_art_status = dados_salvos.get("rel_art_status", "Emitida")
    st.session_state.rel_art_numero = dados_salvos.get("rel_art_numero", "")
    st.session_state.rel_art_inicio = dados_salvos.get("rel_art_inicio", "")
    st.session_state.rel_art_fim = dados_salvos.get("rel_art_fim", "")
    st.session_state.origem_dados_carregados = dados_salvos.get("nome_condominio", "")

    carregar_dados_cadastro_no_relatorio()
    aplicar_dosagens_ultimas_no_relatorio(dados_salvos)


def preparar_cadastro_legado(nome_pasta: str):
    limpar_formulario()
    st.session_state.nome_condominio = humanizar_nome_pasta(nome_pasta)
    st.session_state.origem_dados_carregados = f"{humanizar_nome_pasta(nome_pasta)} (cadastro criado de pasta antiga)"


def preparar_renovacao_no_formulario(dados_salvos: dict) -> tuple[bool, str]:
    data_fim_atual = dados_salvos.get("data_fim", "")
    novo_inicio, novo_fim = calcular_renovacao_anual(data_fim_atual)

    if not novo_inicio or not novo_fim:
        return False, "Não foi possível renovar. A data final salva está ausente ou inválida."

    aplicar_dados_no_formulario(dados_salvos)
    st.session_state.data_inicio = formatar_data_br(novo_inicio)
    st.session_state.data_fim = formatar_data_br(novo_fim)
    st.session_state.data_assinatura = hoje_br()
    st.session_state.origem_dados_carregados = dados_salvos.get("nome_condominio", "")
    return True, "Nova vigência anual preenchida no formulário."


def obter_pasta_atual_do_formulario() -> Path | None:
    nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    if not nome_condominio:
        return None
    return GENERATED_DIR / slugify_nome(nome_condominio)


def salvar_cadastro_sem_gerar_documentos() -> tuple[bool, str]:
    nome_condominio = (st.session_state.get("nome_condominio") or "").strip()
    if not nome_condominio:
        return False, "Informe ao menos o nome do condomínio para salvar o cadastro."

    dados_base = {
        "DATA_ASSINATURA": (st.session_state.get("data_assinatura") or "").strip(),
        "NOME_CONDOMINIO": nome_condominio,
        "CNPJ_CONDOMINIO": (st.session_state.get("cnpj_condominio") or "").strip(),
        "ENDERECO_CONDOMINIO": (st.session_state.get("endereco_condominio") or "").strip(),
        "NOME_SINDICO": (st.session_state.get("nome_sindico") or "").strip(),
        "CPF_SINDICO": (st.session_state.get("cpf_sindico") or "").strip(),
        "VALOR_MENSAL": valor_para_template((st.session_state.get("valor_mensal") or "").strip()),
        "VALOR_ADITIVO": valor_para_template((st.session_state.get("valor_aditivo") or "").strip()),
        "DATA_INICIO": (st.session_state.get("data_inicio") or "").strip(),
        "DATA_FIM": (st.session_state.get("data_fim") or "").strip(),
    }

    erros = validar_para_geracao(dados_base, (st.session_state.get("email_cliente") or "").strip())
    if erros:
        return False, "Corrija os campos antes de salvar o cadastro: " + " | ".join(erros)

    pasta_condominio = obter_pasta_atual_do_formulario()
    assert pasta_condominio is not None
    pasta_condominio.mkdir(parents=True, exist_ok=True)
    salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())
    return True, f"Cadastro salvo/atualizado com sucesso em '{pasta_condominio.name}'."


# =========================================
# HISTÓRICO / PAINEL
# =========================================

def listar_arquivos_pasta(pasta: Path):
    if not pasta.exists():
        return []

    arquivos = []
    for p in pasta.iterdir():
        if p.is_file():
            if p.name in {DADOS_JSON_NAME, MANIFEST_JSON_NAME}:
                continue

            tipo_doc, tipo_ext = classificar_arquivo(p.name)
            arquivos.append(
                {
                    "path": p,
                    "name": p.name,
                    "tipo_doc": tipo_doc,
                    "tipo_ext": tipo_ext,
                    "modificado_em": formatar_data_hora_arquivo(p.stat().st_mtime),
                    "ts": p.stat().st_mtime,
                }
            )

    return sorted(arquivos, key=lambda x: x["ts"], reverse=True)


def localizar_ultimo_documento(arquivos: list, tipo_doc: str, extensao_preferida: str = "PDF"):
    preferidos = [a for a in arquivos if a["tipo_doc"] == tipo_doc and a["tipo_ext"] == extensao_preferida]
    if preferidos:
        return preferidos[0]
    alternativos = [a for a in arquivos if a["tipo_doc"] == tipo_doc]
    return alternativos[0] if alternativos else None


def listar_historico():
    if not GENERATED_DIR.exists():
        return []

    pastas = [p for p in GENERATED_DIR.iterdir() if p.is_dir()]
    pastas = sorted(pastas, key=lambda p: p.stat().st_mtime, reverse=True)

    historico = []
    for pasta in pastas:
        arquivos = listar_arquivos_pasta(pasta)
        dados_salvos = carregar_dados_condominio(pasta)
        data_fim = dados_salvos.get("data_fim", "") if dados_salvos else ""
        status = status_vencimento(data_fim)

        historico.append(
            {
                "nome": pasta.name,
                "pasta": pasta,
                "arquivos": arquivos,
                "total_arquivos": len(arquivos),
                "tem_dados_salvos": (pasta / DADOS_JSON_NAME).exists(),
                "status_vencimento": status,
                "data_fim": data_fim,
            }
        )
    return historico


def listar_painel_vencimentos(alerta_dias: int):
    if not GENERATED_DIR.exists():
        return []

    pastas = [p for p in GENERATED_DIR.iterdir() if p.is_dir()]
    itens = []

    for pasta in pastas:
        dados_salvos = carregar_dados_condominio(pasta)
        arquivos = listar_arquivos_pasta(pasta)

        nome_exibicao = pasta.name
        data_fim = ""
        status = {
            "codigo": "indefinido",
            "rotulo": "Sem vigência válida",
            "mensagem": "Data final ausente ou inválida.",
            "dias": None,
            "css": "status-indefinido",
        }
        origem = "legado_sem_json"

        if dados_salvos:
            nome_exibicao = dados_salvos.get("nome_condominio", pasta.name)
            data_fim = dados_salvos.get("data_fim", "")
            status = status_vencimento(data_fim, alerta_dias)
            origem = "json"

        ultimo_contrato = localizar_ultimo_documento(arquivos, "Contrato")
        ultimo_aditivo = localizar_ultimo_documento(arquivos, "Aditivo")
        ultimo_relatorio = localizar_ultimo_documento(arquivos, "Relatório")

        itens.append(
            {
                "nome_exibicao": nome_exibicao,
                "slug": pasta.name,
                "pasta": pasta,
                "dados": dados_salvos,
                "arquivos": arquivos,
                "data_fim": data_fim,
                "data_fim_dt": parse_data_br(data_fim),
                "status": status,
                "origem": origem,
                "tem_json": dados_salvos is not None,
                "ultimo_contrato": ultimo_contrato,
                "ultimo_aditivo": ultimo_aditivo,
                "ultimo_relatorio": ultimo_relatorio,
            }
        )

    def ordem_item(item):
        status_codigo = item["status"]["codigo"]
        dt_fim = item["data_fim_dt"] or date.max
        dias = item["status"]["dias"]
        dias_sort = dias if dias is not None else 999999

        prioridade = {
            "vencido": 0,
            "vencendo": 1,
            "vigente": 2,
            "indefinido": 3,
        }.get(status_codigo, 9)

        return (prioridade, dt_fim, dias_sort, item["nome_exibicao"].lower())

    return sorted(itens, key=ordem_item)


def filtrar_itens_painel(itens: list, termo: str, filtro_status: str):
    resultado = itens

    termo = (termo or "").strip().lower()
    if termo:
        filtrados = []
        for item in resultado:
            dados = item["dados"] or {}
            campos = [
                item["nome_exibicao"],
                item["slug"],
                dados.get("nome_condominio", ""),
                dados.get("cnpj_condominio", ""),
                dados.get("nome_sindico", ""),
                item["status"]["rotulo"],
            ]
            texto_unico = " ".join(campos).lower()
            if termo in texto_unico:
                filtrados.append(item)
        resultado = filtrados

    mapa_status = {
        "Todos": None,
        "Vigente": "vigente",
        "Vence em breve": "vencendo",
        "Vencido": "vencido",
        "Sem vigência válida": "indefinido",
    }

    codigo = mapa_status.get(filtro_status)
    if codigo:
        resultado = [i for i in resultado if i["status"]["codigo"] == codigo]

    return resultado


def abrir_pasta_windows(caminho: Path):
    try:
        os.startfile(str(caminho))
    except Exception as e:
        st.error(f"Não foi possível abrir a pasta no Windows: {e}")


def abrir_arquivo_windows(caminho: Path):
    try:
        os.startfile(str(caminho))
    except Exception as e:
        st.error(f"Não foi possível abrir o arquivo: {e}")


def deletar_pasta_condominio(pasta: Path):
    if pasta.exists() and pasta.is_dir():
        shutil.rmtree(pasta)


def deletar_arquivo_individual(arquivo: Path):
    pasta = arquivo.parent
    if arquivo.exists() and arquivo.is_file():
        arquivo.unlink()

    if pasta.exists() and pasta.is_dir():
        restantes = list(pasta.iterdir())
        if len(restantes) == 0:
            pasta.rmdir()


# =========================================
# DOCX / PLACEHOLDERS
# =========================================

def substituir_em_paragrafo(paragraph, mapa: dict[str, str]):
    texto_original = paragraph.text
    texto_novo = texto_original
    alterou = False

    for chave, valor in mapa.items():
        if chave in texto_novo:
            texto_novo = texto_novo.replace(chave, valor)
            alterou = True

    if alterou:
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = texto_novo
        else:
            paragraph.add_run(texto_novo)


def substituir_placeholders_doc(doc: Document, mapa: dict[str, str]):
    for paragraph in doc.paragraphs:
        substituir_em_paragrafo(paragraph, mapa)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    substituir_em_paragrafo(paragraph, mapa)

    for section in doc.sections:
        header = section.header
        footer = section.footer

        for paragraph in header.paragraphs:
            substituir_em_paragrafo(paragraph, mapa)

        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        substituir_em_paragrafo(paragraph, mapa)

        for paragraph in footer.paragraphs:
            substituir_em_paragrafo(paragraph, mapa)

        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        substituir_em_paragrafo(paragraph, mapa)


def adicionar_espaco(doc: Document, qtd: int = 1):
    for _ in range(qtd):
        doc.add_paragraph("")


def adicionar_bloco_assinaturas(doc: Document, nome_sindico: str, nome_condominio: str = "", cnpj_condominio: str = ""):
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    adicionar_espaco(doc, 2)

    # Data — parágrafo simples centralizado, fora de qualquer tabela
    p_local = doc.add_paragraph()
    p_local.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_data = p_local.add_run(f"Uberlândia/MG, {hoje_br()}.")
    run_data.font.size = Pt(11)

    adicionar_espaco(doc, 2)

    def preencher_celula(cell, linhas, negrito_idx=None):
        negrito_idx = negrito_idx or []
        cell.paragraphs[0].clear()
        primeira = True
        for i, linha in enumerate(linhas):
            p = cell.paragraphs[0] if primeira else cell.add_paragraph()
            primeira = False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(linha)
            run.font.size = Pt(9)
            run.bold = (i in negrito_idx)

    col_w = Cm(6)

    # ---- Tabela: CONTRATADA | CONTRATANTE ----
    tab1 = doc.add_table(rows=1, cols=2)
    tab1.autofit = False
    for row in tab1.rows:
        for cell in row.cells:
            cell.width = col_w

    linhas_contratada = [
        "_" * 28,
        "AQUA GESTÃO",
        "CONTROLE TÉCNICO DE PISCINAS",
        RESPONSAVEL_TÉCNICO,
        CRQ,
        "CONTRATADA",
    ]
    preencher_celula(tab1.cell(0, 0), linhas_contratada, negrito_idx=[1, 2, 5])

    # Monta linhas do CONTRATANTE — nome quebra por palavras com até 32 chars/linha
    linhas_contratante = ["_" * 28]
    if nome_condominio:
        palavras = nome_condominio.upper().split()
        linha_atual = ""
        for palavra in palavras:
            teste = (linha_atual + " " + palavra).strip()
            if len(teste) <= 32:
                linha_atual = teste
            else:
                if linha_atual:
                    linhas_contratante.append(linha_atual)
                linha_atual = palavra
        if linha_atual:
            linhas_contratante.append(linha_atual)
    if cnpj_condominio:
        linhas_contratante.append(f"CNPJ: {cnpj_condominio}")
    if nome_sindico:
        linhas_contratante.append(nome_sindico)
    linhas_contratante.append("CONTRATANTE")
    preencher_celula(tab1.cell(0, 1), linhas_contratante, negrito_idx=[1, len(linhas_contratante) - 1])

    adicionar_espaco(doc, 2)

    # ---- Tabela de testemunhas ----
    tab2 = doc.add_table(rows=1, cols=2)
    tab2.autofit = False
    for row in tab2.rows:
        for cell in row.cells:
            cell.width = col_w

    preencher_celula(tab2.cell(0, 0), ["_" * 28, "Testemunha 1", "Nome:", "CPF:"])
    preencher_celula(tab2.cell(0, 1), ["_" * 28, "Testemunha 2", "Nome:", "CPF:"])


def converter_docx_para_pdf(docx_path: Path, pdf_path: Path):
    try:
        import pythoncom
        from docx2pdf import convert

        pythoncom.CoInitialize()
        convert(str(docx_path), str(pdf_path))
        return True, None
    except Exception as e:
        return False, str(e)


def gerar_documento(
    template_path: Path,
    output_docx: Path,
    placeholders: dict[str, str],
    incluir_assinaturas: bool = True,
    nome_sindico: str = "",
    nome_condominio: str = "",
    cnpj_condominio: str = "",
):
    if not template_path.exists():
        raise FileNotFoundError(f"Template não encontrado: {template_path.name}")

    doc = Document(str(template_path))
    substituir_placeholders_doc(doc, placeholders)

    if incluir_assinaturas:
        adicionar_bloco_assinaturas(doc, nome_sindico=nome_sindico, nome_condominio=nome_condominio, cnpj_condominio=cnpj_condominio)

    doc.save(str(output_docx))


# =========================================
# EXPORTAÇÃO DE CADASTRO
# =========================================

def gerar_html_resumo_cadastro(item: dict) -> str:
    dados = item["dados"] or {}
    status = item["status"]
    nome = item["nome_exibicao"]

    def val(chave):
        return dados.get(chave, "Não informado") or "Não informado"

    html = f"""
    <html>
    <head>
        <meta charset="utf-8" />
        <title>Resumo de Cadastro - {nome}</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 28px;
                color: #15395f;
            }}
            h1 {{
                font-size: 24px;
                color: #0d3d75;
                margin-bottom: 6px;
            }}
            h2 {{
                font-size: 16px;
                color: #25598d;
                margin-top: 24px;
                margin-bottom: 10px;
            }}
            .sub {{
                color: #5e7691;
                margin-bottom: 20px;
            }}
            .box {{
                border: 1px solid #d7e6f7;
                border-radius: 12px;
                padding: 14px 16px;
                margin-bottom: 12px;
                background: #fbfdff;
            }}
            .line {{
                margin-bottom: 6px;
            }}
            .label {{
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <h1>{nome}</h1>
        <div class="sub">Resumo de cadastro exportado pelo sistema Aqua Gestão</div>

        <div class="box">
            <div class="line"><span class="label">Status:</span> {status['rotulo']}</div>
            <div class="line"><span class="label">Mensagem:</span> {status['mensagem']}</div>
            <div class="line"><span class="label">Data final:</span> {item.get('data_fim') or 'Não informada'}</div>
        </div>

        <h2>Dados do condomínio</h2>
        <div class="box">
            <div class="line"><span class="label">Nome do condomínio:</span> {val('nome_condominio')}</div>
            <div class="line"><span class="label">CNPJ:</span> {val('cnpj_condominio')}</div>
            <div class="line"><span class="label">Endereço:</span> {val('endereco_condominio')}</div>
            <div class="line"><span class="label">Síndico / representante:</span> {val('nome_sindico')}</div>
            <div class="line"><span class="label">CPF:</span> {val('cpf_sindico')}</div>
        </div>

        <h2>Dados contratuais</h2>
        <div class="box">
            <div class="line"><span class="label">Valor mensal:</span> {val('valor_mensal')}</div>
            <div class="line"><span class="label">Valor aditivo:</span> {val('valor_aditivo')}</div>
            <div class="line"><span class="label">Data de início:</span> {val('data_inicio')}</div>
            <div class="line"><span class="label">Data de fim:</span> {val('data_fim')}</div>
            <div class="line"><span class="label">Data de assinatura:</span> {val('data_assinatura')}</div>
        </div>

        <h2>Contato</h2>
        <div class="box">
            <div class="line"><span class="label">WhatsApp:</span> {val('whatsapp_cliente')}</div>
            <div class="line"><span class="label">E-mail:</span> {val('email_cliente')}</div>
            <div class="line"><span class="label">Última atualização:</span> {val('salvo_em')}</div>
        </div>
    </body>
    </html>
    """
    return html


# =========================================
# MENSAGENS / LINKS
# =========================================

def montar_mensagem_envio(
    nome_condominio: str,
    nome_sindico: str,
    caminho_contrato_pdf: Path | None = None,
    caminho_aditivo_pdf: Path | None = None,
) -> str:
    partes = [
        f"Prezado(a) {nome_sindico},",
        "",
        f"Segue em anexo a documentação referente ao condomínio {nome_condominio}:",
        "",
    ]

    if caminho_contrato_pdf and caminho_contrato_pdf.exists():
        partes.append("- Contrato de Responsabilidade Técnica (PDF)")
    if caminho_aditivo_pdf and caminho_aditivo_pdf.exists():
        partes.append("- Aditivo contratual (PDF)")

    partes += [
        "",
        "Permaneço à disposição para quaisquer esclarecimentos.",
        "",
        "Atenciosamente,",
        f"{RESPONSAVEL_TÉCNICO}",
        CRQ,
        "Aqua Gestão – Controle Técnico de Piscinas",
    ]

    return "\n".join(partes)


def link_whatsapp(telefone: str, mensagem: str) -> str:
    somente_numeros = apenas_digitos(telefone or "")
    if not somente_numeros.startswith("55") and somente_numeros:
        somente_numeros = "55" + somente_numeros
    return f"https://wa.me/{somente_numeros}?text={quote(mensagem)}"


def link_email(email: str, assunto: str, corpo: str) -> str:
    return f"mailto:{email}?subject={quote(assunto)}&body={quote(corpo)}"


def componente_copiar(texto: str):
    escaped = (
        texto.replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("${", "\\${")
    )

    st.components.v1.html(
        f"""
        <div style="margin-top:6px;">
            <button
                onclick="navigator.clipboard.writeText(`{escaped}`); this.innerText='Mensagem copiada';"
                style="
                    background:#0d5db8;
                    color:white;
                    border:none;
                    padding:10px 14px;
                    border-radius:10px;
                    cursor:pointer;
                    font-weight:600;
                "
            >
                Copiar mensagem
            </button>
        </div>
        """,
        height=55,
    )


# =========================================
# AÇÕES DO PAINEL
# =========================================

def gerar_aditivo_renovacao_por_painel(pasta: Path, alerta_dias: int) -> tuple[bool, str]:
    dados_salvos = carregar_dados_condominio(pasta)
    if not dados_salvos:
        return False, "Esta pasta ainda não possui dados_condominio.json."

    data_fim_atual = dados_salvos.get("data_fim", "")
    novo_inicio, novo_fim = calcular_renovacao_anual(data_fim_atual)
    if not novo_inicio or not novo_fim:
        return False, "Não foi possível calcular a renovação porque a data final salva está inválida."

    dados_atualizados = dict(dados_salvos)
    dados_atualizados["data_inicio"] = formatar_data_br(novo_inicio)
    dados_atualizados["data_fim"] = formatar_data_br(novo_fim)
    dados_atualizados["data_assinatura"] = hoje_br()
    dados_atualizados["salvo_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    placeholders = {
        "{{DATA_ASSINATURA}}": dados_atualizados.get("data_assinatura", ""),
        "{{NOME_CONDOMINIO}}": dados_atualizados.get("nome_condominio", ""),
        "{{CNPJ_CONDOMINIO}}": dados_atualizados.get("cnpj_condominio", ""),
        "{{ENDERECO_CONDOMINIO}}": dados_atualizados.get("endereco_condominio", ""),
        "{{NOME_SINDICO}}": dados_atualizados.get("nome_sindico", ""),
        "{{CPF_SINDICO}}": dados_atualizados.get("cpf_sindico", ""),
        "{{VALOR_MENSAL}}": valor_para_template(dados_atualizados.get("valor_mensal", "")),
        "{{VALOR_ADITIVO}}": valor_para_template(dados_atualizados.get("valor_aditivo", "")),
        "{{DATA_INICIO}}": dados_atualizados.get("data_inicio", ""),
        "{{DATA_FIM}}": dados_atualizados.get("data_fim", ""),
    }

    nome_condominio = dados_atualizados.get("nome_condominio", pasta.name)
    nome_sindico = dados_atualizados.get("nome_sindico", "")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_nome_aditivo = limpar_nome_arquivo(f"Aditivo_RT_{nome_condominio}_{timestamp}")
    aditivo_docx = pasta / f"{base_nome_aditivo}.docx"
    aditivo_pdf = pasta / f"{base_nome_aditivo}.pdf"

    gerar_documento(
        template_path=TEMPLATE_ADITIVO,
        output_docx=aditivo_docx,
        placeholders=placeholders,
        incluir_assinaturas=True,
        nome_sindico=nome_sindico,
        nome_condominio=nome_condominio,
        cnpj_condominio=dados_atualizados.get("cnpj_condominio", ""),
    )

    ok_pdf, erro_pdf = converter_docx_para_pdf(aditivo_docx, aditivo_pdf)
    salvar_dados_condominio(pasta, dados_atualizados)
    registrar_documento_manifest(
        pasta_condominio=pasta,
        nome_condominio=nome_condominio,
        tipo="Aditivo",
        arquivo_docx=aditivo_docx,
        arquivo_pdf=aditivo_pdf,
        pdf_gerado=ok_pdf,
        erro_pdf=erro_pdf,
        dados_utilizados=dados_atualizados,
    )
    aplicar_dados_no_formulario(dados_atualizados)

    if ok_pdf:
        return True, f"Aditivo de renovação gerado para '{nome_condominio}'. Nova vigência: {dados_atualizados['data_inicio']} até {dados_atualizados['data_fim']}."
    return True, f"Aditivo DOCX de renovação gerado para '{nome_condominio}', mas o PDF falhou: {erro_pdf}"



# =========================================
# RELATÓRIO MENSAL DE RT
# =========================================

def valor_float(texto: str):
    try:
        t = str(texto).replace(",", ".").strip()
        return float(t) if t else None
    except Exception:
        return None


def formatar_data_relatorio_chave(chave: str):
    st.session_state[chave] = formatar_data_digitada(st.session_state.get(chave, ""))


def formatar_art_numero(texto: str) -> str:
    texto = (texto or "").strip()
    return re.sub(r"[^A-Za-z0-9./-]", "", texto)[:40]


def on_change_rel_art_numero():
    st.session_state.rel_art_numero = formatar_art_numero(st.session_state.get("rel_art_numero", ""))


def on_change_rel_art_status():
    status = (st.session_state.get("rel_art_status") or "Emitida").strip()
    if status != "Emitida":
        st.session_state.rel_art_numero = ""
        st.session_state.rel_art_inicio = ""
        st.session_state.rel_art_fim = ""


def obter_status_art_texto(dados_relatorio: dict) -> str:
    status = (dados_relatorio.get("art_status") or "Emitida").strip()
    numero = (dados_relatorio.get("art_numero") or "").strip()
    inicio = (dados_relatorio.get("art_inicio") or "").strip()
    fim = (dados_relatorio.get("art_fim") or "").strip()
    if status == "Emitida":
        numero_final = numero or "N/A"
        if inicio and fim:
            return f"ART nº {numero_final} | Vigência: {inicio} a {fim}"
        return f"ART nº {numero_final}"
    if status == "Em tramitação":
        return "ART em tramitação administrativa"
    return "ART não emitida até a data de emissão deste relatório"


LIMITES_RELATORIO = {
    "ph": (7.2, 7.8, "pH"),
    "cloro_livre": (0.5, 3.0, "cloro residual livre"),
    "alcalinidade": (80.0, 120.0, "alcalinidade total"),
    "dureza": (150.0, 300.0, "dureza cálcica"),
    "cianurico": (30.0, 50.0, "ácido cianúrico"),
}


def avaliar_conformidade_analises(analises: list[dict]) -> dict:
    nao_conformes = []
    houve_leitura = False
    cloraminas_altas = []
    for idx, item in enumerate(analises, start=1):
        linha_tem = any((item.get(k) or "").strip() for k in ["ph", "cloro_livre", "cloro_total", "alcalinidade", "dureza", "cianurico"])
        if not linha_tem:
            continue
        houve_leitura = True
        for campo, (mn, mx, rotulo) in LIMITES_RELATORIO.items():
            v = valor_float(item.get(campo, ""))
            if v is None:
                continue
            if v < mn or v > mx:
                nao_conformes.append(f"Linha {idx}: {rotulo}={v} fora da faixa {mn}–{mx}")
        cl = valor_float(item.get("cloro_livre", ""))
        ct = valor_float(item.get("cloro_total", ""))
        if cl is not None and ct is not None:
            combinado = round(max(ct - cl, 0), 2)
            item["cloro_combinado"] = combinado
            if combinado > 0.2:
                cloraminas_altas.append((idx, combinado))
        else:
            item["cloro_combinado"] = None

    status = "NÃO CONFORME" if nao_conformes else ("CONFORME" if houve_leitura else "EM CORREÇÃO")
    return {
        "status": status,
        "detalhes": nao_conformes,
        "houve_leitura": houve_leitura,
        "cloraminas_altas": cloraminas_altas,
    }


def normalizar_texto(texto: str) -> str:
    texto = (texto or "").lower()
    mapa = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüç", "aaaaaeeeeiiiiooooouuuuc")
    return texto.translate(mapa)


def limpar_paragrafo(paragraph):
    for run in paragraph.runs:
        run.text = ""


def set_cell_text(cell, texto: str):
    texto = texto if texto is not None else ""
    if cell.paragraphs:
        primeira = True
        for p in cell.paragraphs:
            for run in p.runs:
                run.text = ""
            if primeira:
                if p.runs:
                    p.runs[0].text = str(texto)
                else:
                    p.add_run(str(texto))
                primeira = False
    else:
        cell.text = str(texto)


def encontrar_tabela_por_keywords(doc: Document, keywords: list[str]):
    kws = [normalizar_texto(k) for k in keywords]
    melhor = None
    melhor_score = -1
    for table in doc.tables:
        conteudo = " ".join(normalizar_texto(c.text) for row in table.rows[:6] for c in row.cells)
        score = sum(1 for k in kws if k in conteudo)
        if score > melhor_score:
            melhor = table
            melhor_score = score
    return melhor if melhor_score > 0 else None


def preencher_tabela_generica(table, rows_data: list[list[str]], start_row: int = 1):
    if table is None:
        return False
    total_cols = max((len(r.cells) for r in table.rows), default=0)
    if total_cols == 0:
        return False
    for idx, row_data in enumerate(rows_data, start=start_row):
        if idx >= len(table.rows):
            table.add_row()
        row = table.rows[idx]
        for c in range(total_cols):
            valor = row_data[c] if c < len(row_data) else ""
            set_cell_text(row.cells[c], valor)
    return True


def preencher_tabela_identificacao(doc: Document, dados_relatorio: dict):
    tabela = encontrar_tabela_por_keywords(doc, ["Responsável Técnico", "Registro CRQ", "Cliente", "Período de Referência", "Data de Emissão"])
    if tabela is None:
        return False
    mapa_linhas = {
        "responsavel tecnico": dados_relatorio["responsavel_tecnico"],
        "registro crq": dados_relatorio["crq"],
        "qualificacao": dados_relatorio["qualificacao"],
        "empresa": dados_relatorio["empresa_rt"],
        "cliente / estabelecimento": dados_relatorio["nome_condominio"],
        "endereco do local": dados_relatorio["endereco_condominio"],
        "periodo de referencia": f"Mês: {dados_relatorio['mes_referencia']} / Ano: {dados_relatorio['ano_referencia']}",
        "n° art – crq": obter_status_art_texto(dados_relatorio),
        "nº art – crq": obter_status_art_texto(dados_relatorio),
        "data de emissao": dados_relatorio["data_emissao"],
    }
    for row in tabela.rows:
        primeira = normalizar_texto(row.cells[0].text)
        for chave, valor in mapa_linhas.items():
            if chave in primeira and len(row.cells) > 1:
                set_cell_text(row.cells[1], valor)
    return True


def preencher_bloco_conformidades(doc: Document, dados_relatorio: dict):
    tabela_nbr = encontrar_tabela_por_keywords(doc, ["Requisito NBR 11238", "Evidência / Observação"])
    if tabela_nbr is not None and len(tabela_nbr.rows) > 1:
        observacao = dados_relatorio["conformidades"].get("nbr_11238", "") or "Sem observações adicionais registradas."
        for idx in range(1, len(tabela_nbr.rows)):
            row = tabela_nbr.rows[idx]
            if len(row.cells) > 1 and idx == 1:
                set_cell_text(row.cells[1], observacao)

    tabela_nr26 = encontrar_tabela_por_keywords(doc, ["NR-26", "GHS", "FISPQs"])
    if tabela_nr26 is not None and len(tabela_nr26.rows) > 1:
        observacao = dados_relatorio["conformidades"].get("nr_26", "") or "Checklist de GHS/FDS/FISPQ sem observações adicionais registradas."
        for idx in range(1, len(tabela_nr26.rows)):
            row = tabela_nr26.rows[idx]
            if len(row.cells) > 1 and idx == 1:
                set_cell_text(row.cells[1], observacao)
            if len(row.cells) > 2 and idx == 1:
                set_cell_text(row.cells[2], "OK" if observacao else "Pend.")

    tabela_nr6 = encontrar_tabela_por_keywords(doc, ["Equipamentos de Proteção Individual", "CA nº", "Luvas de nitrila"])
    if tabela_nr6 is not None and len(tabela_nr6.rows) > 1:
        ca_map = dados_relatorio["epis"]
        linhas_epi = [
            ("luvas", ca_map.get("luvas_status", "Conforme"), ca_map.get("luvas_ca", "")),
            ("oculos", ca_map.get("oculos_status", "Conforme"), ca_map.get("oculos_ca", "")),
            ("respirador", ca_map.get("respirador_status", "Conforme"), ca_map.get("respirador_ca", "")),
            ("botas", ca_map.get("botas_status", "Conforme"), ca_map.get("botas_ca", "")),
        ]
        mapa_status = {
            "Conforme": ("Sim", "Sim"),
            "Pendente": ("Não", "Não"),
            "Não informado": ("Não", "Não"),
            "N/A": ("Não", "Não"),
        }
        for row in tabela_nr6.rows[1:]:
            texto = normalizar_texto(row.cells[0].text)
            for chave, status_epi, ca in linhas_epi:
                if chave in texto and len(row.cells) > 1:
                    # Coluna 1 = CA nº: apenas o número ou N/A, NUNCA o status
                    ca_valor = ca.strip() if (ca or "").strip() else "N/A"
                    set_cell_text(row.cells[1], ca_valor)
                    fornecido, fiscalizado = mapa_status.get(status_epi, ("Não", "Não"))
                    if len(row.cells) > 2:
                        set_cell_text(row.cells[2], fornecido)
                    if len(row.cells) > 3:
                        set_cell_text(row.cells[3], fiscalizado)


def inserir_assinatura_rt_no_doc(doc: Document):
    assinatura = preparar_assinatura_rt_para_relatorio()
    texto_ass = RESPONSAVEL_TECNICO_ASSINATURA

    # Prioridade máxima: colocar a assinatura no fim do relatório,
    # exatamente acima do campo "Data" do bloco do Responsável Técnico.
    paragrafos = list(doc.paragraphs)
    for i, p in enumerate(paragrafos):
        txt = normalizar_texto(p.text)
        if "responsavel tecnico" in txt:
            for prox in paragrafos[i + 1:i + 8]:
                txt_prox = normalizar_texto(prox.text)
                if "representante do estabelecimento" in txt_prox:
                    break
                if "data:" in txt_prox:
                    try:
                        novo = prox.insert_paragraph_before("")
                    except Exception:
                        novo = doc.add_paragraph()
                    novo.alignment = 1
                    if assinatura:
                        try:
                            novo.add_run().add_picture(str(assinatura), width=Inches(1.45))
                        except Exception:
                            pass
                    return True

    # Fallback: insere ao final se não achar o bloco correto
    if assinatura:
        try:
            novo = doc.add_paragraph()
            novo.alignment = 1
            novo.add_run().add_picture(str(assinatura), width=Inches(1.45))
            novo2 = doc.add_paragraph(texto_ass)
            novo2.alignment = 1
            return True
        except Exception:
            pass

    doc.add_paragraph(texto_ass)
    return False


def salvar_uploads_relatorio(pasta_condominio: Path):
    """Salva fotos localmente E no Google Drive. Retorna caminhos locais para inserir no DOCX."""
    caminhos = []
    arquivos = st.session_state.get("rel_fotos_upload") or []
    pasta_fotos = pasta_condominio / "fotos_relatorio"
    pasta_fotos.mkdir(exist_ok=True)
    nome_cond = (st.session_state.get("rel_nome_condominio") or
                 st.session_state.get("nome_condominio") or
                 pasta_condominio.name)
    mes_ano = datetime.now().strftime("%Y-%m")
    for idx, arquivo in enumerate(arquivos, start=1):
        nome = limpar_nome_arquivo(f"foto_relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{idx}_{arquivo.name}")
        destino = pasta_fotos / nome
        foto_bytes = arquivo.getbuffer()
        with open(destino, "wb") as f:
            f.write(foto_bytes)
        caminhos.append(destino)
        # Upload paralelo para Google Drive
        try:
            drive_upload_foto(
                arquivo_bytes=bytes(foto_bytes),
                nome_arquivo=nome,
                nome_condominio=nome_cond,
                mes_ano=mes_ano,
            )
        except Exception:
            pass  # Falha no Drive não impede gerar o relatório
    return caminhos


def buscar_fotos_drive_para_relatorio(nome_condominio: str, mes_ano: str = None) -> list[Path]:
    """Baixa fotos do Drive para pasta temporária e retorna caminhos locais para inserir no DOCX."""
    import tempfile
    if not mes_ano:
        mes_ano = datetime.now().strftime("%Y-%m")
    try:
        service = conectar_drive()
        if not service:
            return []

        # Busca pasta do condomínio
        q_cond = f"name='{nome_condominio}' and '{DRIVE_FOTOS_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_cond = service.files().list(q=q_cond, fields="files(id,name)").execute()
        pastas_cond = res_cond.get("files", [])
        if not pastas_cond:
            return []
        pasta_cond_id = pastas_cond[0]["id"]

        # Busca pasta do mês
        q_mes = f"name='{mes_ano}' and '{pasta_cond_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_mes = service.files().list(q=q_mes, fields="files(id,name)").execute()
        pastas_mes = res_mes.get("files", [])
        if not pastas_mes:
            return []
        pasta_mes_id = pastas_mes[0]["id"]

        # Lista fotos
        q_fotos = f"'{pasta_mes_id}' in parents and trashed=false"
        res_fotos = service.files().list(q=q_fotos, fields="files(id,name,mimeType)").execute()
        fotos = res_fotos.get("files", [])

        # Baixa para pasta temp
        caminhos = []
        tmp_dir = Path(tempfile.mkdtemp())
        for foto in fotos:
            try:
                conteudo = service.files().get_media(fileId=foto["id"]).execute()
                dest = tmp_dir / foto["name"]
                with open(dest, "wb") as f:
                    f.write(conteudo)
                caminhos.append(dest)
            except Exception:
                pass
        return caminhos
    except Exception as e:
        _log_sheets_erro("buscar_fotos_drive_para_relatorio", e)
        return []




def garantir_campos_analises(qtd: int):
    qtd = max(ANALISES_PADRAO, int(qtd or ANALISES_PADRAO))
    qtd = min(qtd, ANALISES_MAX_SUGERIDO)
    st.session_state.rel_analises_total = qtd
    for i in range(qtd):
        for sufixo in ["data", "ph", "cl", "ct", "alc", "dc", "cya", "operador"]:
            chave = f"rel_analise_{sufixo}_{i}"
            if chave not in st.session_state:
                st.session_state[chave] = ""


def adicionar_analise_extra():
    atual = int(st.session_state.get("rel_analises_total", ANALISES_PADRAO) or ANALISES_PADRAO)
    garantir_campos_analises(atual + 1)


def coletar_analises_relatorio() -> list[dict]:
    itens = []
    qtd = int(st.session_state.get("rel_analises_total", ANALISES_PADRAO) or ANALISES_PADRAO)
    garantir_campos_analises(qtd)
    for i in range(qtd):
        itens.append({
            "data": (st.session_state.get(f"rel_analise_data_{i}") or "").strip(),
            "ph": (st.session_state.get(f"rel_analise_ph_{i}") or "").strip(),
            "cloro_livre": (st.session_state.get(f"rel_analise_cl_{i}") or "").strip(),
            "cloro_total": (st.session_state.get(f"rel_analise_ct_{i}") or "").strip(),
            "alcalinidade": (st.session_state.get(f"rel_analise_alc_{i}") or "").strip(),
            "dureza": (st.session_state.get(f"rel_analise_dc_{i}") or "").strip(),
            "cianurico": (st.session_state.get(f"rel_analise_cya_{i}") or "").strip(),
            "operador": (st.session_state.get(f"rel_analise_operador_{i}") or "").strip(),
        })
    return itens


def coletar_dosagens_relatorio() -> list[dict]:
    itens = []
    for i in range(7):
        itens.append({
            "produto": (st.session_state.get(f"rel_dos_produto_{i}") or "").strip(),
            "fabricante_lote": (st.session_state.get(f"rel_dos_lote_{i}") or "").strip(),
            "quantidade": (st.session_state.get(f"rel_dos_qtd_{i}") or "").strip(),
            "unidade": (st.session_state.get(f"rel_dos_un_{i}") or "").strip(),
            "finalidade": (st.session_state.get(f"rel_dos_finalidade_{i}") or "").strip(),
        })
    return itens


def coletar_recomendacoes_relatorio() -> list[dict]:
    itens = []
    for i in range(5):
        itens.append({
            "recomendacao": (st.session_state.get(f"rel_rec_texto_{i}") or "").strip(),
            "prazo": (st.session_state.get(f"rel_rec_prazo_{i}") or "").strip(),
            "responsavel": (st.session_state.get(f"rel_rec_resp_{i}") or "").strip(),
        })
    return itens


def coletar_observacoes_relatorio() -> list[str]:
    return [(st.session_state.get(f"rel_obs_{i}") or "").strip() for i in range(5)]


def coletar_conformidades_relatorio() -> dict:
    return {
        "nbr_11238": (st.session_state.get("rel_nbr_11238") or "").strip(),
        "nr_26": (st.session_state.get("rel_nr_26") or "").strip(),
        "nr_06": (st.session_state.get("rel_nr_06") or "").strip(),
    }


def coletar_epis_relatorio() -> dict:
    return {
        "luvas_status": (st.session_state.get("rel_epi_luvas_status") or "Conforme").strip(),
        "luvas_ca": (st.session_state.get("rel_epi_luvas_ca") or "").strip(),
        "oculos_status": (st.session_state.get("rel_epi_oculos_status") or "Conforme").strip(),
        "oculos_ca": (st.session_state.get("rel_epi_oculos_ca") or "").strip(),
        "respirador_status": (st.session_state.get("rel_epi_respirador_status") or "Conforme").strip(),
        "respirador_ca": (st.session_state.get("rel_epi_respirador_ca") or "").strip(),
        "botas_status": (st.session_state.get("rel_epi_botas_status") or "Conforme").strip(),
        "botas_ca": (st.session_state.get("rel_epi_botas_ca") or "").strip(),
    }


def gerar_textos_automaticos_relatorio(analises: list[dict], avaliacao: dict) -> dict:
    observacoes = []
    recomendacoes = []
    detalhes = avaliacao.get("detalhes", [])
    if detalhes:
        for detalhe in detalhes[:3]:
            observacoes.append(detalhe.replace("Linha", "Leitura"))
    if avaliacao.get("cloraminas_altas"):
        idx, valor = avaliacao["cloraminas_altas"][0]
        observacoes.append(f"Leitura {idx}: cloro combinado estimado em {valor} mg/L, sugerindo formação de cloraminas e maior carga orgânica.")
        recomendacoes.append("Executar oxidação complementar / supercloração controlada e reavaliar CT, CL e cloro combinado.")
    if any("ph=" in d for d in detalhes):
        recomendacoes.append("Ajustar pH para a faixa operacional de 7,2 a 7,8 e repetir a leitura após estabilização.")
    if any("cloro residual livre" in d for d in detalhes):
        recomendacoes.append("Revisar a dosagem de desinfetante e a demanda oxidante da água.")
    if any("ácido cianúrico" in d for d in detalhes):
        recomendacoes.append("Reavaliar o estabilizante e considerar renovação parcial de água quando tecnicamente indicada.")
    if any("alcalinidade total" in d for d in detalhes):
        recomendacoes.append("Corrigir alcalinidade total para aumentar a estabilidade do pH e reduzir flutuações operacionais.")
    if any("dureza cálcica" in d for d in detalhes):
        recomendacoes.append("Ajustar dureza cálcica para reduzir risco de corrosão ou incrustação.")

    if not observacoes:
        observacoes.append("Parâmetros avaliados sem desvios críticos nas leituras registradas no período.")
    if not recomendacoes:
        recomendacoes.append("Manter rotina de monitoramento, rastreabilidade de dosagens e controle técnico periódico.")

    diagnostico = []
    if avaliacao["status"] == "CONFORME":
        diagnostico.append("No período avaliado, os parâmetros registrados indicam condição satisfatória de controle físico-químico da água, sem desvios relevantes frente às faixas operacionais adotadas pelo sistema.")
    elif avaliacao["status"] == "EM CORREÇÃO":
        diagnostico.append("O sistema registra fase operacional de correção, recomendando continuidade das medidas técnicas, reforço de monitoramento e nova conferência analítica após estabilização.")
    else:
        diagnostico.append("Foram observadas não conformidades físico-químicas que exigem ação corretiva e reavaliação técnica sequencial.")
    if avaliacao.get("cloraminas_altas"):
        diagnostico.append("A presença de cloro combinado acima do desejável sugere formação de cloraminas, condição compatível com carga orgânica elevada, consumo do desinfetante e possível redução da eficiência sanitizante.")
    if detalhes:
        diagnostico.append("Resumo automático dos desvios: " + "; ".join(detalhes[:5]) + ".")

    return {
        "diagnostico": " ".join(diagnostico),
        "observacoes": observacoes[:5],
        "recomendacoes": recomendacoes[:5],
    }


def aplicar_textos_automaticos_relatorio():
    analises = coletar_analises_relatorio()
    avaliacao = avaliar_conformidade_analises(analises)
    textos = gerar_textos_automaticos_relatorio(analises, avaliacao)

    # Nunca apaga parâmetros digitados. Só recalcula os campos textuais do parecer.
    st.session_state.rel_diagnostico = textos["diagnostico"]

    for i in range(5):
        chave_obs = f"rel_obs_{i}"
        texto_obs = textos["observacoes"][i] if i < len(textos["observacoes"]) else ""
        st.session_state[chave_obs] = texto_obs

    for i in range(5):
        chave_rec = f"rel_rec_texto_{i}"
        chave_prazo = f"rel_rec_prazo_{i}"
        chave_resp = f"rel_rec_resp_{i}"
        texto_rec = textos["recomendacoes"][i] if i < len(textos["recomendacoes"]) else ""
        st.session_state[chave_rec] = texto_rec
        st.session_state[chave_prazo] = "Imediato" if texto_rec and i == 0 else ("Próxima rotina" if texto_rec else "")
        st.session_state[chave_resp] = "Operação / RT" if texto_rec else ""

    # Não escrever diretamente na chave do selectbox após o widget existir.
    if st.session_state.get("rel_status_agua") != "EM CORREÇÃO":
        st.session_state["rel_status_agua"] = avaliacao["status"]


def montar_dados_relatorio() -> dict:
    nome_condominio = (st.session_state.get("rel_nome_condominio") or "").strip()
    representante = (st.session_state.get("rel_representante") or "").strip()
    dados_base = obter_snapshot_relatorio_independente()
    analises = coletar_analises_relatorio()
    avaliacao = avaliar_conformidade_analises(analises)
    status_manual = (st.session_state.get("rel_status_agua") or "CONFORME").strip()
    status_final = "EM CORREÇÃO" if status_manual == "EM CORREÇÃO" else avaliacao["status"]
    textos_auto = gerar_textos_automaticos_relatorio(analises, avaliacao)
    diagnostico_base = (st.session_state.get("rel_diagnostico") or "").strip() or textos_auto["diagnostico"]

    recomendacoes = coletar_recomendacoes_relatorio()
    if not any(r["recomendacao"] for r in recomendacoes):
        recomendacoes = [{"recomendacao": t, "prazo": "Imediato" if i == 0 else "Próxima rotina", "responsavel": "Operação / RT"} for i, t in enumerate(textos_auto["recomendacoes"])]

    observacoes = coletar_observacoes_relatorio()
    if not any(observacoes):
        observacoes = textos_auto["observacoes"]

    return {
        "empresa_rt": EMPRESA_RT,
        "responsavel_tecnico": RESPONSAVEL_TÉCNICO,
        "assinatura_rt_texto": RESPONSAVEL_TECNICO_ASSINATURA,
        "crq": CRQ,
        "qualificacao": QUALIFICACAO_RT,
        "certificacoes": CERTIFICACOES_RT,
        "nome_condominio": nome_condominio,
        "cnpj_condominio": dados_base.get("cnpj_condominio", ""),
        "endereco_condominio": dados_base.get("endereco_condominio", ""),
        "representante": representante,
        "cpf_cnpj_representante": dados_base.get("cpf_sindico", ""),
        "tipo_atendimento": (st.session_state.get("rel_tipo_atendimento") or "Contrato ativo").strip(),
        "mes_referencia": (st.session_state.get("rel_mes_referencia") or "").strip(),
        "ano_referencia": (st.session_state.get("rel_ano_referencia") or "").strip(),
        "art_status": (st.session_state.get("rel_art_status") or "Emitida").strip(),
        "art_numero": (st.session_state.get("rel_art_numero") or "").strip() if (st.session_state.get("rel_art_status") or "Emitida").strip() == "Emitida" else "N/A",
        "art_inicio": (st.session_state.get("rel_art_inicio") or "").strip() if (st.session_state.get("rel_art_status") or "Emitida").strip() == "Emitida" else "N/A",
        "art_fim": (st.session_state.get("rel_art_fim") or "").strip() if (st.session_state.get("rel_art_status") or "Emitida").strip() == "Emitida" else "N/A",
        "art_texto": "",
        "data_emissao": (st.session_state.get("rel_data_emissao") or hoje_br()).strip(),
        "status_agua": status_final,
        "diagnostico": diagnostico_base,
        "analises": analises,
        "dosagens": coletar_dosagens_relatorio(),
        "recomendacoes": recomendacoes,
        "observacoes": observacoes,
        "conformidades": coletar_conformidades_relatorio(),
        "epis": coletar_epis_relatorio(),
        "avaliacao_automatica": avaliacao,
    }


def validar_relatorio_mensal(dados_relatorio: dict) -> list[str]:
    erros = []
    if not dados_relatorio.get("nome_condominio"):
        erros.append("Informe o nome do condomínio/estabelecimento no próprio relatório antes de gerar.")
    if not dados_relatorio.get("mes_referencia"):
        erros.append("Informe o mês de referência do relatório.")
    if not dados_relatorio.get("ano_referencia"):
        erros.append("Informe o ano de referência do relatório.")
    if not TEMPLATE_RELATORIO.exists():
        erros.append("O arquivo relatorio_mensal.docx não foi localizado na pasta do projeto.")
    if not validar_data_br(dados_relatorio.get("data_emissao", "")):
        erros.append("Data de emissão do relatório inválida.")
    if dados_relatorio.get("art_status") == "Emitida":
        if not dados_relatorio.get("art_numero"):
            erros.append("Informe o número da ART ou altere o status para Não emitida / Em tramitação.")
        if not validar_data_br(dados_relatorio.get("art_inicio", "")):
            erros.append("Vigência ART - início inválida.")
        if not validar_data_br(dados_relatorio.get("art_fim", "")):
            erros.append("Vigência ART - fim inválida.")
    return erros


def append_relatorio_fallback(doc: Document, dados_relatorio: dict):
    doc.add_page_break()
    doc.add_paragraph("COMPLEMENTO AUTOMÁTICO – DADOS ESTRUTURADOS DO RELATÓRIO MENSAL")
    doc.add_paragraph(f"Condomínio: {dados_relatorio['nome_condominio']}")
    doc.add_paragraph(f"Mês/Ano de referência: {dados_relatorio['mes_referencia']}/{dados_relatorio['ano_referencia']}")
    doc.add_paragraph(f"ART: {obter_status_art_texto(dados_relatorio)}")
    doc.add_paragraph(f"Data de emissão: {dados_relatorio['data_emissao']}")
    doc.add_paragraph(f"Responsável técnico: {dados_relatorio['assinatura_rt_texto']}")
    doc.add_paragraph(f"Certificações relevantes: {dados_relatorio['certificacoes']}")
    doc.add_paragraph(f"Status geral da água: {dados_relatorio['status_agua']}")
    doc.add_paragraph(f"Diagnóstico técnico: {dados_relatorio['diagnostico']}")
    doc.add_paragraph("Base normativa referencial do relatório: Lei nº 2.800/1956; Decreto nº 85.877/1981; Lei nº 6.839/1980; Resolução CFQ nº 332/2025; ABNT NBR 10339; NR-26; NR-06.")
    doc.add_paragraph("Recomenda-se a guarda e organização deste documento e de seus registros correlatos por prazo mínimo de 5 (cinco) anos, para fins de rastreabilidade, auditoria, controle documental e segurança jurídica.")
    doc.add_paragraph("Nota técnica: análises microbiológicas não são realizadas in loco e dependem de laboratório acreditado, sob responsabilidade de contratação do cliente.")

    doc.add_paragraph("Observações automáticas:")
    for obs in dados_relatorio["observacoes"]:
        if obs:
            doc.add_paragraph(f"• {obs}")

    doc.add_paragraph("Recomendações automáticas:")
    for item in dados_relatorio["recomendacoes"]:
        if item.get("recomendacao"):
            doc.add_paragraph(f"• {item['recomendacao']} | Prazo: {item.get('prazo','')} | Responsável: {item.get('responsavel','')}")


def atualizar_textos_normativos(doc: Document):
    # Mapeia TODAS as variações conhecidas de CRQ errado e normas defasadas
    trocas_ordenadas = [
        # CRQ – variações mais específicas primeiro para evitar substituição parcial
        ("CRQ CRQ 024025748 – 4ª Região", CRQ),
        ("CRQ 024025748 – 4ª Região", CRQ),
        ("CRQ-IV | CRQ 024025748", CRQ),
        ("CRQ-IV – CRQ 024025748", CRQ),
        ("CRQ-IV", "CRQ-MG 2ª Região"),
        # Normas
        ("Resolução CFQ nº 491/2020", "Resolução CFQ nº 332/2025"),
        ("Res. Normativa CFQ nº 01/1982", "Decreto nº 85.877/1981 e normas profissionais aplicáveis do Sistema CFQ/CRQs"),
        ("Portaria MS nº 888/2021", "Portaria GM/MS nº 888/2021 (referência sanitária complementar para água de consumo humano)"),
        ("NBR 10818", "ABNT NBR 10339"),
    ]
    # Aplica na ordem para evitar substituições aninhadas incorretas
    for de, para in trocas_ordenadas:
        substituir_placeholders_doc(doc, {de: para})


def preencher_relatorio_mensal_docx(template_path: Path, output_docx: Path, dados_relatorio: dict, fotos: list[Path] | None = None):
    doc = Document(str(template_path))
    dados_relatorio["art_texto"] = obter_status_art_texto(dados_relatorio)

    # ---- Placeholders primários (substituição em todo o documento) ----
    placeholders = {
        "{{NOME_CONDOMINIO}}": dados_relatorio["nome_condominio"],
        "{{CNPJ_CONDOMINIO}}": dados_relatorio["cnpj_condominio"],
        "{{ENDERECO_CONDOMINIO}}": dados_relatorio["endereco_condominio"],
        "{{NOME_SINDICO}}": dados_relatorio["representante"],
        "{{RESPONSAVEL_TÉCNICO}}": dados_relatorio["responsavel_tecnico"],
        "{{RESPONSAVEL_TECNICO_ASSINATURA}}": dados_relatorio["assinatura_rt_texto"],
        "{{CRQ}}": dados_relatorio["crq"],
        "{{QUALIFICACAO_RT}}": dados_relatorio["qualificacao"],
        "{{CERTIFICACOES_RT}}": dados_relatorio["certificacoes"],
        "{{EMPRESA_RT}}": dados_relatorio["empresa_rt"],
        "{{MES_REFERENCIA}}": dados_relatorio["mes_referencia"],
        "{{ANO_REFERENCIA}}": dados_relatorio["ano_referencia"],
        "{{ART_NUMERO}}": dados_relatorio["art_numero"],
        "{{ART_INICIO}}": dados_relatorio["art_inicio"],
        "{{ART_FIM}}": dados_relatorio["art_fim"],
        "{{ART_TEXTO}}": dados_relatorio["art_texto"],
        "{{DATA_EMISSAO}}": dados_relatorio["data_emissao"],
        "{{STATUS_AGUA}}": dados_relatorio["status_agua"],
        "{{DIAGNOSTICO_TÉCNICO}}": dados_relatorio["diagnostico"],
        "{{DIAGNOSTICO_TECNICO}}": dados_relatorio["diagnostico"],
    }
    substituir_placeholders_doc(doc, placeholders)
    atualizar_textos_normativos(doc)
    preencher_tabela_identificacao(doc, dados_relatorio)
    preencher_bloco_conformidades(doc, dados_relatorio)

    # ---- Tabela de análises ----
    tabela_analises = encontrar_tabela_por_keywords(doc, ["Data", "CRL", "Cl. Total", "Operador"])
    linhas_analises = []
    for item in dados_relatorio["analises"]:
        if any((item.get(k) or "").strip() for k in ["data", "ph", "cloro_livre", "cloro_total", "alcalinidade", "dureza", "cianurico", "operador"]):
            linhas_analises.append([item["data"], item["ph"], item["cloro_livre"], item["cloro_total"], item["alcalinidade"], item["dureza"], item["cianurico"], item["operador"]])
    if linhas_analises:
        preencher_tabela_generica(tabela_analises, linhas_analises, start_row=1)

    # ---- Tabela de dosagens ----
    tabela_dos = encontrar_tabela_por_keywords(doc, ["Produto Químico", "Fabricante / Lote", "Finalidade Técnica"])
    linhas_dos = []
    for item in dados_relatorio["dosagens"]:
        if any(item.values()):
            linhas_dos.append([item["produto"], item["fabricante_lote"], item["quantidade"], item["unidade"], item["finalidade"]])
    if linhas_dos:
        preencher_tabela_generica(tabela_dos, linhas_dos, start_row=1)

    # ---- Tabela de recomendações ----
    tabela_rec = encontrar_tabela_por_keywords(doc, ["Recomendação Técnica", "Prazo", "Responsável"])
    linhas_rec = []
    for idx, item in enumerate(dados_relatorio["recomendacoes"], start=1):
        if item.get("recomendacao"):
            linhas_rec.append([str(idx), item["recomendacao"], item.get("prazo", ""), item.get("responsavel", "")])
    if linhas_rec:
        preencher_tabela_generica(tabela_rec, linhas_rec, start_row=1)

    # ---- Parecer principal e aviso legal em parágrafos soltos ----
    # Também cobre o caso de o template ter esses textos fora de tabelas/placeholders.
    aviso_legal_texto = (
        "AVISO LEGAL: Recomenda-se a guarda e organização deste documento e de seus registros correlatos "
        "por prazo mínimo de 5 (cinco) anos, para fins de rastreabilidade, auditoria, controle documental "
        "e segurança jurídica. Emitido sob responsabilidade técnica do RT acima identificado. "
        "Análises microbiológicas in loco não integram este relatório e dependem de laboratório acreditado, "
        "sob responsabilidade de contratação do contratante."
    )

    # Rastreia se os quadros principais foram encontrados e preenchidos
    quadro_diagnostico_preenchido = False

    for p in doc.paragraphs:
        txt = normalizar_texto(p.text)
        # Quadro de diagnóstico – qualquer parágrafo que contenha o texto modelo do template
        if "diagnostico:" in txt and dados_relatorio["diagnostico"] not in p.text:
            # Preserva formatação dos runs, substituindo só o conteúdo
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = f"Diagnóstico: {dados_relatorio['diagnostico']}"
            else:
                p.add_run(f"Diagnóstico: {dados_relatorio['diagnostico']}")
            quadro_diagnostico_preenchido = True
        elif "conforme" in txt and ("nao conforme" in txt or "em correcao" in txt):
            # Linha de status geral – substitui o texto modelo
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = f"Status geral da água: {dados_relatorio['status_agua']}"
            else:
                p.add_run(f"Status geral da água: {dados_relatorio['status_agua']}")
        elif "aviso legal:" in txt:
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = aviso_legal_texto
            else:
                p.add_run(aviso_legal_texto)

    # ---- Também busca diagnóstico dentro de tabelas do template ----
    if not quadro_diagnostico_preenchido:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        txt = normalizar_texto(p.text)
                        if "diagnostico:" in txt and dados_relatorio["diagnostico"] not in p.text:
                            for run in p.runs:
                                run.text = ""
                            if p.runs:
                                p.runs[0].text = f"Diagnóstico: {dados_relatorio['diagnostico']}"
                            else:
                                p.add_run(f"Diagnóstico: {dados_relatorio['diagnostico']}")

    # ---- Assinatura automática REMOVIDA definitivamente. ----
    # O bloco oficial de assinatura existe no template relatorio_mensal.docx
    # e não deve ser duplicado ou complementado aqui.

    # ---- Fotos ----
    if fotos:
        doc.add_page_break()
        doc.add_paragraph("REGISTRO FOTOGRÁFICO")
        for foto in fotos:
            try:
                doc.add_paragraph(foto.name)
                doc.add_picture(str(foto), width=Inches(5.8))
            except Exception:
                doc.add_paragraph(f"Não foi possível inserir a foto: {foto.name}")

    # ---- Fallback CONDICIONAL: só adiciona se as tabelas principais não foram encontradas ----
    # O fallback NÃO deve ser gerado se o template já tem os quadros de análise, dosagem e recomendação.
    template_tem_analises = encontrar_tabela_por_keywords(doc, ["Data", "CRL", "Cl. Total", "Operador"]) is not None
    template_tem_dosagens = encontrar_tabela_por_keywords(doc, ["Produto Químico", "Fabricante / Lote", "Finalidade Técnica"]) is not None
    if not template_tem_analises and not template_tem_dosagens:
        append_relatorio_fallback(doc, dados_relatorio)

    doc.save(str(output_docx))



def gerar_html_relatorio_visita(lancamento: dict, nome_condominio: str) -> str:
    """Gera HTML profissional do relatório de visita para download como PDF."""

    def fmt(val, sufixo=""):
        return f"{val}{sufixo}" if val and str(val).strip() else "—"

    def param_box(nome, val, mn, mx, quinzenal=False):
        v = valor_float(val)
        if v is None:
            status_cls = "nd"
            status_txt = "Quinzenal" if quinzenal else "Não medido"
            val_txt = "—"
        elif v < mn or v > mx:
            status_cls = "warn"
            status_txt = "Fora da faixa"
            val_txt = str(val).replace(".", ",")
        else:
            status_cls = "ok"
            status_txt = "Conforme"
            val_txt = str(val).replace(".", ",")
        badge = ' <span class="q15">15d</span>' if quinzenal else ""
        return f"""
        <div class="param-box {status_cls}">
          <div class="pnm">{nome}{badge}</div>
          <div class="pval">{val_txt}</div>
          <div class="pst">{status_txt}</div>
        </div>"""

    # Alertas automáticos
    alertas = []
    checks = [
        (lancamento.get("ph",""), 7.2, 7.8, "pH", "Fora da faixa ideal (7,2–7,8). Corrigir imediatamente."),
        (lancamento.get("cloro_livre",""), 0.5, 3.0, "CRL", "Cloro livre fora da faixa (0,5–3,0 mg/L)."),
        (lancamento.get("alcalinidade",""), 80, 120, "Alcalinidade", "Abaixo do ideal (80–120 mg/L). Aplicar bicarbonato de sódio."),
        (lancamento.get("dureza",""), 150, 300, "Dureza DC", "Fora da faixa (150–300 mg/L)."),
        (lancamento.get("cianurico",""), 30, 50, "CYA", "Ácido cianúrico fora da faixa (30–50 mg/L)."),
    ]
    for val, mn, mx, rot, msg in checks:
        v = valor_float(val)
        if v is not None and (v < mn or v > mx):
            alertas.append(f"{rot}: {str(val).replace('.', ',')} mg/L — {msg}")
    cloraminas = lancamento.get("cloraminas", "")
    if cloraminas and valor_float(cloraminas) is not None and valor_float(cloraminas) > 0.2:
        alertas.append(f"Cloraminas {str(cloraminas).replace('.', ',')} mg/L — acima do limite (0,2 mg/L). Chocar piscina.")

    alertas_html = ""
    if alertas:
        for a in alertas:
            alertas_html += f'<div class="alerta"><div class="alerta-icon">!</div><div class="alerta-txt">{a}</div></div>'
    else:
        alertas_html = '<div class="alerta ok-all"><div class="alerta-txt">Todos os parâmetros medidos dentro da faixa ideal.</div></div>'

    # Dosagens
    dosagens = lancamento.get("dosagens", [])
    dos_html = ""
    for d in dosagens:
        if d.get("produto","").strip():
            qtd = f"{d.get('quantidade','')} {d.get('unidade','')}".strip()
            fin = d.get("finalidade","")
            dos_html += f'<div class="dos-row"><span class="dos-nome">{d["produto"]}</span><span class="dos-detalhe">{qtd}{(" · " + fin) if fin else ""}</span></div>'
    if not dos_html:
        dos_html = '<p style="font-size:12px;color:#8a9ab0;font-style:italic;">Nenhuma dosagem registrada.</p>'

    # Cloraminas box
    clor_box = ""
    if cloraminas and valor_float(cloraminas) is not None:
        v_cl = valor_float(cloraminas)
        cls_cl = "ok" if v_cl <= 0.2 else "warn"
        st_cl = "Conforme" if v_cl <= 0.2 else "Fora da faixa"
        clor_box = f"""
        <div class="param-box {cls_cl}">
          <div class="pnm">Cloraminas</div>
          <div class="pval">{str(cloraminas).replace(".", ",")}</div>
          <div class="pst">{st_cl}</div>
        </div>"""

    obs = lancamento.get("observacao","").strip()
    obs_html = f'<div class="obs-txt">"{obs}"</div>' if obs else '<p style="font-size:12px;color:#8a9ab0;font-style:italic;">Sem observações.</p>'

    data_hoje = date.today().strftime("%d/%m/%Y")
    operador = lancamento.get("operador","") or "—"

    # ── Seção de piscinas (múltiplas ou única) ────────────────────────────────
    piscinas_lista = lancamento.get("piscinas", [])
    if not piscinas_lista:
        # Compatibilidade com lançamentos antigos (sem múltiplas piscinas)
        piscinas_lista = [{
            "nome": "Piscina",
            "ph": lancamento.get("ph",""),
            "cloro_livre": lancamento.get("cloro_livre",""),
            "cloro_total": lancamento.get("cloro_total",""),
            "cloraminas": lancamento.get("cloraminas",""),
            "alcalinidade": lancamento.get("alcalinidade",""),
            "dureza": lancamento.get("dureza",""),
            "cianurico": lancamento.get("cianurico",""),
        }]

    piscinas_html_section = ""
    for pisc in piscinas_lista:
        p_clor_box = ""
        p_clor_val = pisc.get("cloraminas","")
        if p_clor_val and valor_float(p_clor_val) is not None:
            v_cl = valor_float(p_clor_val)
            cls_cl = "ok" if v_cl <= 0.2 else "warn"
            st_cl = "Conforme" if v_cl <= 0.2 else "Fora da faixa"
            p_clor_box = f'''<div class="param-box {cls_cl}"><div class="pnm">Cloraminas</div><div class="pval">{str(p_clor_val).replace(".", ",")}</div><div class="pst">{st_cl}</div></div>'''
        piscinas_html_section += f'''
  <div class="card">
    <div class="sec-title">🏊 {pisc.get("nome","Piscina")} — Parâmetros</div>
    <div class="param-grid">
      {param_box("pH", pisc.get("ph",""), 7.2, 7.8)}
      {param_box("CRL mg/L", pisc.get("cloro_livre",""), 0.5, 3.0)}
      {param_box("Alc. mg/L", pisc.get("alcalinidade",""), 80, 120, quinzenal=True)}
      {param_box("Dureza mg/L", pisc.get("dureza",""), 150, 300, quinzenal=True)}
      {param_box("CYA mg/L", pisc.get("cianurico",""), 30, 50, quinzenal=True)}
      {p_clor_box}
    </div>
  </div>'''

    # Problemas / ocorrências
    problemas = lancamento.get("problemas","").strip()
    problemas_html = f'''
  <div class="card">
    <div class="sec-title">Problemas / Ocorrências</div>
    <div class="obs-txt">"{problemas}"</div>
  </div>''' if problemas else ""

    # ── Fotos do Drive (base64 embutido por categoria) ───────────────────────
    import base64 as _b64

    def _ids_to_html(ids_list, titulo):
        if not ids_list:
            return ""
        imgs = ""
        for fid in ids_list:
            try:
                fb = drive_baixar_foto(fid)
                if fb:
                    b64 = _b64.b64encode(fb).decode("utf-8")
                    imgs += f'<div style="margin-bottom:6px;"><img src="data:image/jpeg;base64,{b64}" style="width:100%;border-radius:6px;border:1px solid #d0d8e4;" /></div>'
            except Exception:
                pass
        if not imgs:
            return ""
        return f'<div style="margin-bottom:12px;"><div style="font-size:10px;color:#1e4d8c;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">{titulo}</div>{imgs}</div>'

    fotos_antes_ids  = lancamento.get("fotos_antes_ids",  lancamento.get("fotos_drive_ids", []))
    fotos_depois_ids = lancamento.get("fotos_depois_ids", [])
    fotos_cmaq_ids   = lancamento.get("fotos_cmaq_ids",   [])

    _fotos_content = (
        _ids_to_html(fotos_antes_ids,  "Antes do tratamento") +
        _ids_to_html(fotos_depois_ids, "Depois do tratamento") +
        _ids_to_html(fotos_cmaq_ids,   "Casa de máquinas")
    )

    fotos_html_section = f'''
  <div class="card">
    <div class="sec-title">Registro fotográfico</div>
    {_fotos_content if _fotos_content else '<p style="font-size:12px;color:#8a9ab0;font-style:italic;">Nenhuma foto registrada nesta visita.</p>'}
  </div>'''

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Relatório de Visita — {nome_condominio}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{font-family:Arial,Helvetica,sans-serif;background:#f4f6f9;color:#1a2a4a;}}
  .page{{max-width:640px;margin:0 auto;padding:16px;}}
  .card{{background:#fff;border:1px solid #d0d8e4;border-radius:12px;padding:18px 20px;margin-bottom:12px;}}
  .hdr-top{{display:flex;justify-content:space-between;align-items:flex-start;}}
  .hdr-logo{{display:flex;align-items:center;gap:12px;}}
  .logo-ball{{width:48px;height:48px;border-radius:50%;background:#1e4d8c;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#fff;flex-shrink:0;}}
  .hdr-empresa{{font-size:16px;font-weight:700;color:#1a2a4a;letter-spacing:0.3px;}}
  .hdr-sub{{font-size:10px;color:#8a9ab0;letter-spacing:0.8px;text-transform:uppercase;margin-top:2px;}}
  .hdr-right{{text-align:right;}}
  .doc-titulo{{font-size:14px;font-weight:700;color:#1e4d8c;}}
  .doc-num{{font-size:10px;color:#8a9ab0;margin-top:2px;}}
  hr{{border:none;border-top:1px solid #d0d8e4;margin:12px 0;}}
  .info-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;}}
  .info-lbl{{font-size:10px;color:#8a9ab0;text-transform:uppercase;letter-spacing:0.5px;}}
  .info-val{{font-size:13px;color:#1a2a4a;font-weight:600;margin-top:2px;}}
  .sec-title{{font-size:10px;font-weight:700;color:#1e4d8c;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid #1e4d8c;}}
  .param-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;}}
  .param-box{{border:1px solid #d0d8e4;border-radius:8px;padding:10px 8px;text-align:center;}}
  .param-box.ok{{border-color:#2e7d32;background:#f1f8f1;}}
  .param-box.warn{{border-color:#e65100;background:#fff8f0;}}
  .param-box.nd{{border-color:#d0d8e4;background:#f8f9fb;}}
  .pnm{{font-size:9px;color:#8a9ab0;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:4px;}}
  .q15{{background:#e8f0fb;color:#1e4d8c;border-radius:4px;padding:1px 4px;font-size:8px;font-weight:700;}}
  .pval{{font-size:20px;font-weight:700;color:#1a2a4a;margin:2px 0;}}
  .param-box.ok .pval{{color:#2e7d32;}}
  .param-box.warn .pval{{color:#e65100;}}
  .param-box.nd .pval{{color:#b0bec5;}}
  .pst{{font-size:9px;}}
  .param-box.ok .pst{{color:#388e3c;}}
  .param-box.warn .pst{{color:#e65100;}}
  .param-box.nd .pst{{color:#b0bec5;font-style:italic;}}
  .alerta{{display:flex;align-items:flex-start;gap:10px;padding:10px 12px;border-radius:8px;background:#fff8f0;border:1px solid #e65100;margin-bottom:8px;}}
  .alerta.ok-all{{background:#f1f8f1;border-color:#2e7d32;}}
  .alerta-icon{{width:16px;height:16px;border-radius:50%;background:#e65100;display:flex;align-items:center;justify-content:center;font-size:10px;color:#fff;font-weight:700;flex-shrink:0;margin-top:1px;}}
  .alerta-txt{{font-size:12px;color:#b84200;line-height:1.5;}}
  .alerta.ok-all .alerta-txt{{color:#2e7d32;}}
  .dos-row{{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #eef1f5;}}
  .dos-row:last-child{{border-bottom:none;}}
  .dos-nome{{font-size:13px;color:#1a2a4a;font-weight:600;}}
  .dos-detalhe{{font-size:11px;color:#8a9ab0;text-align:right;}}
  .obs-txt{{font-size:13px;color:#4a5568;line-height:1.7;font-style:italic;padding:4px 0;}}
  .assin-bloco{{display:flex;justify-content:space-between;align-items:flex-end;padding-top:12px;margin-top:6px;border-top:1px solid #d0d8e4;}}
  .assin-esq{{font-size:11px;color:#4a5568;line-height:1.8;}}
  .assin-esq strong{{color:#1a2a4a;font-size:13px;}}
  .crq-badge{{display:inline-block;font-size:9px;color:#1e4d8c;background:#e8f0fb;border:1px solid #b5d0f0;border-radius:99px;padding:2px 8px;margin-top:4px;}}
  .assin-dir{{text-align:center;}}
  .assin-linha{{border-top:1px solid #1a2a4a;width:140px;margin-bottom:4px;}}
  .assin-nome{{font-size:10px;color:#8a9ab0;}}
  .rodape{{text-align:center;font-size:10px;color:#8a9ab0;padding:8px 0 4px;}}
  @media print{{
    body{{background:#fff;}}
    .page{{padding:0;}}
    .card{{border:1px solid #ccc;border-radius:0;box-shadow:none;page-break-inside:avoid;}}
  }}
</style>
</head>
<body>
<div class="page">

  <div class="card">
    <div class="hdr-top">
      <div class="hdr-logo">
        <div class="logo-ball">RT</div>
        <div>
          <div class="hdr-empresa">AQUA GESTÃO</div>
          <div class="hdr-sub">Controle Técnico de Piscinas</div>
        </div>
      </div>
      <div class="hdr-right">
        <div class="doc-titulo">Relatório de Visita</div>
        <div class="doc-num">Emitido em {data_hoje}</div>
      </div>
    </div>
    <hr>
    <div class="info-grid">
      <div><div class="info-lbl">Condomínio / Local</div><div class="info-val">{nome_condominio}</div></div>
      <div><div class="info-lbl">Data da visita</div><div class="info-val">{fmt(lancamento.get("data",""))}</div></div>
      <div><div class="info-lbl">Operador</div><div class="info-val">{operador}</div></div>
      <div><div class="info-lbl">Responsável técnico</div><div class="info-val">Thyago F. Silveira</div></div>
    </div>
  </div>

  {piscinas_html_section}

  <div class="card">
    <div class="sec-title">Alertas técnicos</div>
    {alertas_html}
  </div>

  <div class="card">
    <div class="sec-title">Dosagens aplicadas</div>
    {dos_html}
  </div>

  <div class="card">
    <div class="sec-title">Observações</div>
    {obs_html}
  </div>

  {fotos_html_section}

  {problemas_html}

  <div class="card">
    <div class="sec-title">Responsabilidade técnica</div>
    <div class="assin-bloco">
      <div class="assin-esq">
        <strong>Thyago Fernando da Silveira</strong><br>
        Técnico em Química · NR-26 · NR-6<br>
        <span class="crq-badge">CRQ-MG 2ª Região · CRQ 024025748</span>
      </div>
      <div class="assin-dir">
        <div class="assin-linha"></div>
        <div class="assin-nome">Assinatura / carimbo RT</div>
      </div>
    </div>
  </div>

  <div class="rodape">
    Aqua Gestão – Controle Técnico de Piscinas · Documento gerado automaticamente
  </div>

</div>
</body>
</html>"""
    return html


def gerar_pdf_relatorio_visita(lancamento: dict, nome_condominio: str) -> bytes:
    """Gera PDF do relatório de visita usando ReportLab. Retorna bytes do PDF."""
    import io
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    import base64 as _b64

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm,
    )

    # Cores da marca
    AZUL_ESCURO = colors.HexColor("#1a2a4a")
    AZUL_MEDIO  = colors.HexColor("#1e4d8c")
    AZUL_CLARO  = colors.HexColor("#e8f0fb")
    CINZA       = colors.HexColor("#8a9ab0")
    VERDE_OK    = colors.HexColor("#2e7d32")
    VERDE_BG    = colors.HexColor("#f1f8f1")
    LARANJA     = colors.HexColor("#e65100")
    LARANJA_BG  = colors.HexColor("#fff8f0")
    BORDA       = colors.HexColor("#d0d8e4")

    styles = getSampleStyleSheet()

    def estilo(nome, **kw):
        return ParagraphStyle(nome, **kw)

    s_titulo    = estilo("titulo",    fontSize=16, textColor=AZUL_ESCURO, fontName="Helvetica-Bold", spaceAfter=2)
    s_sub       = estilo("sub",       fontSize=8,  textColor=CINZA,       fontName="Helvetica",      spaceAfter=6, leading=10)
    s_sec       = estilo("sec",       fontSize=8,  textColor=AZUL_MEDIO,  fontName="Helvetica-Bold", spaceAfter=6, leading=10)
    s_body      = estilo("body",      fontSize=9,  textColor=AZUL_ESCURO, fontName="Helvetica",      leading=13)
    s_body_sm   = estilo("body_sm",   fontSize=8,  textColor=CINZA,       fontName="Helvetica",      leading=11)
    s_alerta    = estilo("alerta",    fontSize=8,  textColor=LARANJA,     fontName="Helvetica",      leading=11)
    s_ok        = estilo("ok",        fontSize=8,  textColor=VERDE_OK,    fontName="Helvetica",      leading=11)
    s_center    = estilo("center",    fontSize=8,  textColor=CINZA,       fontName="Helvetica",      alignment=TA_CENTER)
    s_bold      = estilo("bold",      fontSize=9,  textColor=AZUL_ESCURO, fontName="Helvetica-Bold", leading=13)

    elems = []

    # ── CABEÇALHO ─────────────────────────────────────────────────────────────
    data_hoje = date.today().strftime("%d/%m/%Y")
    operador  = lancamento.get("operador","") or "—"

    hdr_data = [
        [Paragraph("<b>AQUA GESTÃO</b>", estilo("hdr1", fontSize=14, textColor=AZUL_ESCURO, fontName="Helvetica-Bold")),
         Paragraph(f"<b>Relatório de Visita</b><br/><font size=8 color='#8a9ab0'>Emitido em {data_hoje}</font>", estilo("hdr2", fontSize=11, textColor=AZUL_MEDIO, fontName="Helvetica-Bold", alignment=TA_RIGHT))],
        [Paragraph("Controle Técnico de Piscinas", s_sub), ""],
    ]
    t_hdr = Table(hdr_data, colWidths=["60%","40%"])
    t_hdr.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LINEBELOW", (0,1), (-1,1), 0.5, BORDA),
        ("BOTTOMPADDING", (0,1), (-1,1), 8),
    ]))
    elems.append(t_hdr)
    elems.append(Spacer(1, 6))

    # Info básica
    info_data = [
        [Paragraph("<font size=7 color='#8a9ab0'>CONDOMÍNIO / LOCAL</font><br/>" + f"<b>{nome_condominio}</b>", s_body),
         Paragraph("<font size=7 color='#8a9ab0'>DATA DA VISITA</font><br/>" + f"<b>{lancamento.get('data','—')}</b>", s_body)],
        [Paragraph("<font size=7 color='#8a9ab0'>OPERADOR</font><br/>" + f"<b>{operador}</b>", s_body),
         Paragraph("<font size=7 color='#8a9ab0'>RESP. TÉCNICO</font><br/><b>Thyago F. Silveira</b>", s_body)],
    ]
    t_info = Table(info_data, colWidths=["50%","50%"])
    t_info.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, BORDA),
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f8fafd")),
        ("PADDING", (0,0), (-1,-1), 6),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, colors.HexColor("#f8fafd")]),
    ]))
    elems.append(t_info)
    elems.append(Spacer(1, 10))

    # ── PISCINAS ──────────────────────────────────────────────────────────────
    piscinas_lista = lancamento.get("piscinas", [])
    if not piscinas_lista:
        piscinas_lista = [{
            "nome": "Piscina", "ph": lancamento.get("ph",""),
            "cloro_livre": lancamento.get("cloro_livre",""), "cloro_total": lancamento.get("cloro_total",""),
            "cloraminas": lancamento.get("cloraminas",""), "alcalinidade": lancamento.get("alcalinidade",""),
            "dureza": lancamento.get("dureza",""), "cianurico": lancamento.get("cianurico",""),
        }]

    PARAMS = [
        ("pH",          "ph",          7.2, 7.8,  False),
        ("CRL mg/L",    "cloro_livre", 0.5, 3.0,  False),
        ("Alc. mg/L",   "alcalinidade",80, 120,   True),
        ("Dureza mg/L", "dureza",      150, 300,   True),
        ("CYA mg/L",    "cianurico",   30,  50,    True),
    ]

    for pisc in piscinas_lista:
        elems.append(Paragraph(f"🏊 {pisc.get('nome','Piscina')} — Parâmetros analisados", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))

        param_rows = []
        header_row = [Paragraph("<b>Parâmetro</b>", s_body_sm),
                      Paragraph("<b>Valor</b>", s_body_sm),
                      Paragraph("<b>Faixa ideal</b>", s_body_sm),
                      Paragraph("<b>Status</b>", s_body_sm),
                      Paragraph("<b>Obs</b>", s_body_sm)]
        param_rows.append(header_row)

        faixas_txt = {"pH":"7,2–7,8","CRL mg/L":"0,5–3,0","Alc. mg/L":"80–120","Dureza mg/L":"150–300","CYA mg/L":"30–50"}
        row_colors = []

        for label, key, mn, mx, quinzenal in PARAMS:
            val_raw = pisc.get(key, "")
            v = valor_float(val_raw)
            q_txt = " (15d)" if quinzenal else ""
            if v is None:
                status_txt = "Não medido"
                val_fmt = "—"
                bg = colors.white
            elif v < mn or v > mx:
                status_txt = "⚠ Fora da faixa"
                val_fmt = str(val_raw).replace(".", ",")
                bg = LARANJA_BG
            else:
                status_txt = "✓ Conforme"
                val_fmt = str(val_raw).replace(".", ",")
                bg = VERDE_BG
            row_colors.append(bg)
            param_rows.append([
                Paragraph(f"{label}{q_txt}", s_body_sm),
                Paragraph(f"<b>{val_fmt}</b>", s_body),
                Paragraph(faixas_txt.get(label,"—"), s_body_sm),
                Paragraph(status_txt, s_ok if "Conforme" in status_txt else (s_alerta if "Fora" in status_txt else s_body_sm)),
                Paragraph("", s_body_sm),
            ])

        # Cloraminas
        clor_raw = pisc.get("cloraminas","")
        v_cl = valor_float(clor_raw)
        if v_cl is not None:
            bg_cl = VERDE_BG if v_cl <= 0.2 else LARANJA_BG
            st_cl = "✓ Conforme" if v_cl <= 0.2 else "⚠ Fora da faixa"
            row_colors.append(bg_cl)
            param_rows.append([
                Paragraph("Cloraminas", s_body_sm),
                Paragraph(f"<b>{str(clor_raw).replace('.', ',')}</b>", s_body),
                Paragraph("≤ 0,2", s_body_sm),
                Paragraph(st_cl, s_ok if "Conforme" in st_cl else s_alerta),
                Paragraph("", s_body_sm),
            ])

        t_param = Table(param_rows, colWidths=["22%","15%","20%","28%","15%"])
        ts = [
            ("GRID", (0,0), (-1,-1), 0.3, BORDA),
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("TEXTCOLOR", (0,0), (-1,0), AZUL_MEDIO),
            ("PADDING", (0,0), (-1,-1), 5),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 7),
        ]
        for i, bg in enumerate(row_colors):
            ts.append(("BACKGROUND", (0, i+1), (-1, i+1), bg))
        t_param.setStyle(TableStyle(ts))
        elems.append(t_param)
        elems.append(Spacer(1, 10))

    # ── ALERTAS ───────────────────────────────────────────────────────────────
    alertas_gerais = []
    for pisc in piscinas_lista:
        for val_r, mn, mx, rot in [
            (pisc.get("ph",""), 7.2, 7.8, "pH"),
            (pisc.get("cloro_livre",""), 0.5, 3.0, "CRL"),
            (pisc.get("alcalinidade",""), 80, 120, "Alcalinidade"),
            (pisc.get("dureza",""), 150, 300, "Dureza DC"),
            (pisc.get("cianurico",""), 30, 50, "CYA"),
        ]:
            v = valor_float(val_r)
            if v is not None and (v < mn or v > mx):
                alertas_gerais.append(f"⚠ {pisc.get('nome','Piscina')} — {rot}: {str(val_r).replace('.', ',')} — fora da faixa ideal.")

    if alertas_gerais:
        elems.append(Paragraph("Alertas técnicos", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))
        for a in alertas_gerais:
            elems.append(Paragraph(a, s_alerta))
            elems.append(Spacer(1, 3))
        elems.append(Spacer(1, 6))

    # ── DOSAGENS ──────────────────────────────────────────────────────────────
    dosagens = lancamento.get("dosagens", [])
    dosagens = [d for d in dosagens if d.get("produto","").strip()]
    if dosagens:
        elems.append(Paragraph("Dosagens aplicadas", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))
        dos_rows = [[
            Paragraph("<b>Produto</b>", s_body_sm),
            Paragraph("<b>Quantidade</b>", s_body_sm),
            Paragraph("<b>Finalidade</b>", s_body_sm),
        ]]
        for d in dosagens:
            qtd = f"{d.get('quantidade','')} {d.get('unidade','')}".strip()
            dos_rows.append([
                Paragraph(d.get("produto",""), s_body),
                Paragraph(qtd, s_body_sm),
                Paragraph(d.get("finalidade",""), s_body_sm),
            ])
        t_dos = Table(dos_rows, colWidths=["40%","25%","35%"])
        t_dos.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.3, BORDA),
            ("BACKGROUND", (0,0), (-1,0), AZUL_CLARO),
            ("TEXTCOLOR", (0,0), (-1,0), AZUL_MEDIO),
            ("PADDING", (0,0), (-1,-1), 5),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafd")]),
        ]))
        elems.append(t_dos)
        elems.append(Spacer(1, 10))

    # ── PROBLEMAS ─────────────────────────────────────────────────────────────
    problemas = lancamento.get("problemas","").strip()
    if problemas:
        elems.append(Paragraph("Problemas / Ocorrências", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))
        t_prob = Table([[Paragraph(f"⚠ {problemas}", s_alerta)]], colWidths=["100%"])
        t_prob.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), LARANJA_BG),
            ("BOX", (0,0), (-1,-1), 0.5, LARANJA),
            ("PADDING", (0,0), (-1,-1), 8),
            ("RADIUS", (0,0), (-1,-1), 4),
        ]))
        elems.append(t_prob)
        elems.append(Spacer(1, 10))

    # ── OBSERVAÇÃO ────────────────────────────────────────────────────────────
    obs = lancamento.get("observacao","").strip()
    if obs:
        elems.append(Paragraph("Observações", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))
        elems.append(Paragraph(f'"{obs}"', s_body_sm))
        elems.append(Spacer(1, 10))

    # ── FOTOS ─────────────────────────────────────────────────────────────────
    def _add_fotos_b64(b64_list, titulo):
        """Adiciona fotos ao PDF a partir de lista de base64 — 1 por linha, tamanho máximo."""
        if not b64_list:
            return
        import io as _io
        from reportlab.platypus import Image as RLImage
        from PIL import Image as _PILR, ImageOps as _IOps
        elems.append(Paragraph(titulo, s_body_sm))
        LARGURA_MAX = 15 * cm   # largura útil da página
        ALTURA_MAX  = 18 * cm   # altura máxima por foto
        for b64_str in b64_list[:6]:
            try:
                fb = _b64.b64decode(b64_str)
                # Aplica rotação EXIF antes de medir
                _pil = _PILR.open(_io.BytesIO(fb))
                _pil = _IOps.exif_transpose(_pil)
                _iw, _ih = _pil.size
                # Salva versão corrigida
                _buf_corr = _io.BytesIO()
                _pil.convert("RGB").save(_buf_corr, format="JPEG", quality=85)
                _buf_corr.seek(0)
                # Calcula dimensões mantendo proporção
                ratio = _iw / _ih
                if ratio >= 1:  # paisagem
                    w = LARGURA_MAX
                    h = min(w / ratio, ALTURA_MAX)
                    w = h * ratio
                else:  # retrato
                    h = ALTURA_MAX
                    w = min(h * ratio, LARGURA_MAX)
                    h = w / ratio
                img = RLImage(_buf_corr, width=w, height=h)
                t_foto = Table([[img]], colWidths=[LARGURA_MAX])
                t_foto.setStyle(TableStyle([
                    ("ALIGN",   (0,0), (-1,-1), "CENTER"),
                    ("VALIGN",  (0,0), (-1,-1), "MIDDLE"),
                    ("PADDING", (0,0), (-1,-1), 2),
                ]))
                elems.append(t_foto)
                elems.append(Spacer(1, 8))
            except Exception:
                pass

    fotos_antes_b64  = lancamento.get("fotos_antes_b64",  [])
    fotos_depois_b64 = lancamento.get("fotos_depois_b64", [])
    fotos_cmaq_b64   = lancamento.get("fotos_cmaq_b64",   [])

    if fotos_antes_b64 or fotos_depois_b64 or fotos_cmaq_b64:
        elems.append(Paragraph("Registro fotográfico", s_sec))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=AZUL_MEDIO, spaceAfter=4))
        _add_fotos_b64(fotos_antes_b64,  "Antes do tratamento:")
        _add_fotos_b64(fotos_depois_b64, "Depois do tratamento:")
        _add_fotos_b64(fotos_cmaq_b64,   "Casa de máquinas:")
        elems.append(Spacer(1, 6))

    # ── ASSINATURA RT ─────────────────────────────────────────────────────────
    elems.append(HRFlowable(width="100%", thickness=0.5, color=BORDA, spaceAfter=8))
    ass_data = [[
        Paragraph("<b>Thyago Fernando da Silveira</b><br/>Técnico em Química · NR-26 · NR-6<br/><font size=7 color='#1e4d8c'>CRQ-MG 2ª Região · CRQ 024025748</font>", s_body),
        Paragraph("<br/><br/>___________________________<br/><font size=7 color='#8a9ab0'>Assinatura / carimbo RT</font>", s_center),
    ]]
    t_ass = Table(ass_data, colWidths=["60%","40%"])
    t_ass.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "BOTTOM"),
        ("PADDING", (0,0), (-1,-1), 4),
    ]))
    elems.append(t_ass)

    elems.append(Spacer(1, 4))
    elems.append(Paragraph("Aqua Gestão – Controle Técnico de Piscinas · Documento gerado automaticamente", s_center))

    doc.build(elems)
    buffer.seek(0)
    return buffer.read()


def gerar_relatorio_visita_docx(
    output_path: Path,
    nome_local: str,
    cnpj: str,
    endereco: str,
    responsavel: str,
    operador: str,
    mes: str,
    ano: str,
    lancamentos: list,
    obs_geral: str = "",
    incluir_rt: bool = True,
    fotos: list = None,
) -> tuple[bool, str]:
    """
    Gera relatório técnico de visitas em DOCX — unificado para RT e sem RT.
    Se incluir_rt=True: inclui assinatura e dados do RT.
    Se incluir_rt=False: omite dados do RT (relatório sem RT).
    """
    try:
        from docx import Document as _DocxDoc
        from docx.shared import Pt, Cm, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = _DocxDoc()
        for section in doc.sections:
            section.top_margin    = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        def _par(texto, bold=False, size=11, align=None, italic=False):
            p = doc.add_paragraph()
            if align: p.alignment = align
            r = p.add_run(texto)
            r.bold = bold; r.italic = italic
            r.font.size = Pt(size)
            return p

        def _tabela_info(dados: list):
            """Cria tabela de 2 colunas com rótulo → valor."""
            t = doc.add_table(rows=len(dados), cols=2)
            t.style = "Table Grid"
            for i, (rot, val) in enumerate(dados):
                t.cell(i,0).paragraphs[0].add_run(rot).bold = True
                t.cell(i,0).paragraphs[0].runs[0].font.size = Pt(10)
                t.cell(i,1).paragraphs[0].add_run(str(val or "—"))
                t.cell(i,1).paragraphs[0].runs[0].font.size = Pt(10)
            doc.add_paragraph()

        # ── CABEÇALHO ─────────────────────────────────────────────────────────
        _par("AQUA GESTÃO – CONTROLE TÉCNICO DE PISCINAS", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
        if incluir_rt:
            _par(f"Responsável Técnico: {RESPONSAVEL_TÉCNICO} | {CRQ}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
            _par(f"{QUALIFICACAO_RT} | Certificações: {CERTIFICACOES_RT}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _par("Relatório Técnico de Visitas", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()

        # ── IDENTIFICAÇÃO ─────────────────────────────────────────────────────
        _par("1. IDENTIFICAÇÃO DO LOCAL", bold=True, size=11)
        _tabela_info([
            ("Local / Condomínio", nome_local),
            ("CNPJ", cnpj or "Não informado"),
            ("Endereço", endereco or "Não informado"),
            ("Responsável / Síndico", responsavel or "Não informado"),
            ("Operador de campo", operador or "Não informado"),
            ("Período de referência", f"{mes}/{ano}"),
        ])

        # ── ANÁLISES ──────────────────────────────────────────────────────────
        _par("2. ANÁLISES FÍSICO-QUÍMICAS", bold=True, size=11)

        cabecalho = ["Data", "pH", "CRL mg/L", "CT mg/L", "Alc. mg/L", "Dureza mg/L", "CYA mg/L", "Operador"]
        t_anal = doc.add_table(rows=1 + len(lancamentos), cols=len(cabecalho))
        t_anal.style = "Table Grid"

        # Cabeçalho
        for j, cab in enumerate(cabecalho):
            cell = t_anal.cell(0, j)
            cell.paragraphs[0].add_run(cab).bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)

        # Dados
        for i, lc in enumerate(lancamentos):
            piscinas = lc.get("piscinas", [])
            if piscinas:
                lc_d = piscinas[0]
            else:
                lc_d = lc
            valores = [
                lc.get("data",""),
                lc_d.get("ph", lc.get("ph","")),
                lc_d.get("cloro_livre", lc.get("cloro_livre","")),
                lc_d.get("cloro_total", lc.get("cloro_total","")),
                lc_d.get("alcalinidade", lc.get("alcalinidade","")),
                lc_d.get("dureza", lc.get("dureza","")),
                lc_d.get("cianurico", lc.get("cianurico","")),
                lc.get("operador",""),
            ]
            for j, val in enumerate(valores):
                cell = t_anal.cell(i+1, j)
                cell.paragraphs[0].add_run(str(val or "—"))
                cell.paragraphs[0].runs[0].font.size = Pt(9)

        doc.add_paragraph()

        # Se múltiplas piscinas, adiciona tabelas extras
        todas_piscinas = set()
        for lc in lancamentos:
            for p in lc.get("piscinas",[]):
                if p.get("nome","") and p["nome"] != "Piscina":
                    todas_piscinas.add(p["nome"])

        for pisc_nome in sorted(todas_piscinas):
            if pisc_nome == lancamentos[0].get("piscinas",[{}])[0].get("nome","") if lancamentos and lancamentos[0].get("piscinas") else True:
                continue
            _par(f"Análises — {pisc_nome}", bold=True, size=10)
            t_p = doc.add_table(rows=1, cols=len(cabecalho))
            t_p.style = "Table Grid"
            for j, cab in enumerate(cabecalho):
                t_p.cell(0,j).paragraphs[0].add_run(cab).bold = True
                t_p.cell(0,j).paragraphs[0].runs[0].font.size = Pt(8)
            for lc in lancamentos:
                for p in lc.get("piscinas",[]):
                    if p.get("nome","") == pisc_nome:
                        row = t_p.add_row()
                        for j, val in enumerate([
                            lc.get("data",""), p.get("ph",""), p.get("cloro_livre",""),
                            p.get("cloro_total",""), p.get("alcalinidade",""),
                            p.get("dureza",""), p.get("cianurico",""), lc.get("operador",""),
                        ]):
                            row.cells[j].paragraphs[0].add_run(str(val or "—"))
                            row.cells[j].paragraphs[0].runs[0].font.size = Pt(8)
            doc.add_paragraph()

        # ── DOSAGENS ──────────────────────────────────────────────────────────
        _par("3. DOSAGENS APLICADAS", bold=True, size=11)
        todas_dosagens = []
        for lc in lancamentos:
            data_lc = lc.get("data","")
            for d in lc.get("dosagens",[]):
                if d.get("produto","").strip():
                    todas_dosagens.append({**d, "data": data_lc})

        if todas_dosagens:
            t_dos = doc.add_table(rows=1 + len(todas_dosagens), cols=5)
            t_dos.style = "Table Grid"
            for j, cab in enumerate(["Data", "Produto", "Quantidade", "Unidade", "Finalidade"]):
                t_dos.cell(0,j).paragraphs[0].add_run(cab).bold = True
                t_dos.cell(0,j).paragraphs[0].runs[0].font.size = Pt(9)
            for i, d in enumerate(todas_dosagens):
                for j, val in enumerate([
                    d.get("data",""), d.get("produto",""),
                    d.get("quantidade",""), d.get("unidade",""), d.get("finalidade",""),
                ]):
                    t_dos.cell(i+1,j).paragraphs[0].add_run(str(val or "—"))
                    t_dos.cell(i+1,j).paragraphs[0].runs[0].font.size = Pt(9)
        else:
            _par("Nenhuma dosagem registrada no período.", size=10, italic=True)
        doc.add_paragraph()

        # ── PROBLEMAS / OCORRÊNCIAS ───────────────────────────────────────────
        problemas_lista = [f"{lc.get('data','')}: {lc['problemas']}"
                           for lc in lancamentos if lc.get("problemas","").strip()]
        if problemas_lista:
            _par("4. PROBLEMAS / OCORRÊNCIAS", bold=True, size=11)
            for prob in problemas_lista:
                _par(f"⚠ {prob}", size=10)
            doc.add_paragraph()
            secao_obs = 5
        else:
            secao_obs = 4

        # ── OBSERVAÇÕES ───────────────────────────────────────────────────────
        obs_lista = [f"{lc.get('data','')}: {lc['observacao']}"
                     for lc in lancamentos if lc.get("observacao","").strip()]
        if obs_geral:
            obs_lista.insert(0, obs_geral)
        if obs_lista:
            _par(f"{secao_obs}. OBSERVAÇÕES GERAIS", bold=True, size=11)
            for obs in obs_lista:
                _par(obs, size=10)
            doc.add_paragraph()
            secao_fotos = secao_obs + 1
        else:
            secao_fotos = secao_obs

        # ── FOTOS ─────────────────────────────────────────────────────────────
        if fotos:
            _par(f"{secao_fotos}. REGISTRO FOTOGRÁFICO", bold=True, size=11)
            for foto_path in fotos:
                try:
                    _par(foto_path.name, size=9)
                    p_foto = doc.add_paragraph()
                    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_foto.add_run().add_picture(str(foto_path), width=Inches(5.5))
                except Exception:
                    _par(f"[Foto não inserida: {foto_path.name}]", size=9, italic=True)
            doc.add_paragraph()
            secao_aviso = secao_fotos + 1
        else:
            secao_aviso = secao_fotos

        # ── AVISO LEGAL ───────────────────────────────────────────────────────
        _par(f"{secao_aviso}. AVISO LEGAL E GUARDA DOCUMENTAL", bold=True, size=10)
        _par(
            "Recomenda-se a guarda e organização deste documento por prazo mínimo de 5 anos, "
            "para fins de rastreabilidade, auditoria e segurança jurídica. "
            "Este relatório é de caráter técnico-informativo e não substitui laudo laboratorial microbiológico.",
            size=9
        )
        doc.add_paragraph()

        # ── TEXTO RT (apenas no relatório sem RT — Bem Star) ──────────────────
        if not incluir_rt:
            doc.add_page_break()
            _par("SOBRE RESPONSABILIDADE TÉCNICA (RT)", bold=True, size=11)
            doc.add_paragraph()
            for linha in TEXTO_RT_SEM_RT.strip().split("\n\n"):
                if linha.startswith("SOBRE"):
                    continue
                _par(linha.strip(), size=10)
                doc.add_paragraph()
            # Logo Bem Star
            _logo_bs_path = None
            for _lp in LOGO_BEM_STAR_CANDIDATOS:
                if _lp.exists():
                    _logo_bs_path = _lp
                    break
            if _logo_bs_path:
                try:
                    p_logo = doc.add_paragraph()
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_logo.add_run().add_picture(str(_logo_bs_path), width=Inches(3.5))
                except Exception:
                    pass

        # ── ASSINATURA ────────────────────────────────────────────────────────
        _par(f"Uberlândia/MG, {hoje_br()}.", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()
        t_ass = doc.add_table(rows=1, cols=2)
        t_ass.autofit = True

        if incluir_rt:
            ass_rt = (
                f"___________________________\n"
                f"{RESPONSAVEL_TÉCNICO}\n"
                f"{CRQ}\n"
                f"Aqua Gestão – Controle Técnico de Piscinas"
            )
        else:
            ass_rt = (
                f"___________________________\n"
                f"Bem Star Piscinas\n"
                f"CNPJ: {CNPJ_BEM_STAR}"
            )

        ass_resp = (
            f"___________________________\n"
            f"{responsavel or 'Responsável local'}\n"
            f"{nome_local}"
        )

        for cell_a, texto_a in [(t_ass.cell(0,0), ass_rt), (t_ass.cell(0,1), ass_resp)]:
            cell_a.paragraphs[0].clear()
            cell_a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_a.paragraphs[0].add_run(texto_a).font.size = Pt(9)

        doc.save(str(output_path))
        return True, ""
    except Exception as e:
        return False, str(e)

def gerar_relatorio_mensal() -> tuple[bool, str]:
    dados_relatorio = montar_dados_relatorio()
    erros = validar_relatorio_mensal(dados_relatorio)
    if erros:
        return False, " | ".join(erros)

    nome_condominio = dados_relatorio["nome_condominio"]
    pasta_condominio = GENERATED_DIR / slugify_nome(nome_condominio)
    pasta_condominio.mkdir(parents=True, exist_ok=True)
    if st.session_state.get("rel_salvar_alteracoes_cadastro"):
        salvar_relatorio_no_cadastro_principal()
        salvar_dados_condominio(pasta_condominio, salvar_snapshot_formulario())
    else:
        salvar_dados_condominio(pasta_condominio, obter_snapshot_relatorio_independente())
    fotos_salvas = salvar_uploads_relatorio(pasta_condominio)

    # Se não há fotos no upload atual, busca do Google Drive (fotos do mês)
    if not fotos_salvas:
        mes_ano_rel = datetime.now().strftime("%Y-%m")
        fotos_drive = buscar_fotos_drive_para_relatorio(nome_condominio, mes_ano_rel)
        if fotos_drive:
            fotos_salvas = fotos_drive

    # Também busca fotos dos lançamentos de campo salvos localmente
    if not fotos_salvas and pasta_condominio.exists():
        pasta_fotos_campo = pasta_condominio / "fotos_campo"
        if pasta_fotos_campo.exists():
            fotos_campo = sorted(pasta_fotos_campo.glob("*"))
            fotos_campo = [f for f in fotos_campo if f.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")]
            if fotos_campo:
                fotos_salvas = fotos_campo

    data_nome = datetime.now().strftime("%Y%m%d")
    base_nome = limpar_nome_arquivo(f"{data_nome}_{nome_condominio}_RELATORIO_RT")
    relatorio_docx = pasta_condominio / f"{base_nome}.docx"
    relatorio_pdf = pasta_condominio / f"{base_nome}.pdf"

    preencher_relatorio_mensal_docx(TEMPLATE_RELATORIO, relatorio_docx, dados_relatorio, fotos=fotos_salvas)
    ok_pdf, erro_pdf = converter_docx_para_pdf(relatorio_docx, relatorio_pdf)

    registrar_documento_manifest(
        pasta_condominio=pasta_condominio,
        nome_condominio=nome_condominio,
        tipo="Relatório",
        arquivo_docx=relatorio_docx,
        arquivo_pdf=relatorio_pdf,
        pdf_gerado=ok_pdf,
        erro_pdf=erro_pdf,
        dados_utilizados={
            "TIPO_ATENDIMENTO": dados_relatorio.get("tipo_atendimento"),
            "REPRESENTANTE": dados_relatorio.get("representante"),
            "CPF_CNPJ_REPRESENTANTE": dados_relatorio.get("cpf_cnpj_representante"),
            "ART_STATUS": dados_relatorio.get("art_status"),
            "ART_TEXTO": obter_status_art_texto(dados_relatorio),
        },
        extras={"fotos": [p.name for p in fotos_salvas]},
    )

    st.session_state.ultimos_docs_gerados = st.session_state.get("ultimos_docs_gerados") or {}
    st.session_state.ultimos_docs_gerados.update({
        "relatorio_docx": str(relatorio_docx) if relatorio_docx.exists() else None,
        "relatorio_pdf": str(relatorio_pdf) if ok_pdf and relatorio_pdf.exists() else None,
    })

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    if dados_relatorio["avaliacao_automatica"]["detalhes"]:
        st.warning("Diagnóstico automático marcou NÃO CONFORME com base nos parâmetros fora de faixa.")
        for item in dados_relatorio["avaliacao_automatica"]["detalhes"]:
            st.write(f"- {item}")
    if dados_relatorio["avaliacao_automatica"].get("cloraminas_altas"):
        for idx, valor in dados_relatorio["avaliacao_automatica"]["cloraminas_altas"]:
            st.write(f"- Linha {idx}: cloro combinado estimado em {valor} mg/L.")
    c1, c2, c3 = st.columns(3)
    with c1:
        if relatorio_docx.exists():
            with open(relatorio_docx, "rb") as f:
                st.download_button("Baixar DOCX do relatório", data=f, file_name=relatorio_docx.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with c2:
        if ok_pdf and relatorio_pdf.exists():
            with open(relatorio_pdf, "rb") as f:
                st.download_button("Baixar PDF do relatório", data=f, file_name=relatorio_pdf.name, mime="application/pdf", use_container_width=True)
        else:
            st.warning(f"PDF do relatório não gerado. Erro: {erro_pdf}")
    with c3:
        if st.button("Abrir pasta do condomínio", key="abrir_pasta_relatorio", use_container_width=True):
            abrir_pasta_windows(pasta_condominio)
    st.markdown("</div>", unsafe_allow_html=True)
    return True, f"Relatório mensal registrado com sucesso para {nome_condominio}."


# =========================================
# PROCESSAMENTO DE VALIDAÇÃO
# =========================================

def validar_para_geracao(dados_base: dict, email_cliente: str) -> list[str]:
    faltando = validar_campos_obrigatorios(dados_base)
    if faltando:
        return [f"Preencha o campo obrigatório: {item}" for item in faltando]

    erros_formato = validar_campos_formato(dados_base, email_cliente)
    return erros_formato


# =========================================
# SESSÃO
# =========================================

def inicializar_campos():
    defaults = {
        "nome_condominio": "",
        "cnpj_condominio": "",
        "endereco_condominio": "",
        "nome_sindico": "",
        "cpf_sindico": "",
        "valor_mensal": "",
        "valor_aditivo": "",
        "data_inicio": "",
        "data_fim": "",
        "data_assinatura": hoje_br(),
        "whatsapp_cliente": "",
        "email_cliente": "",
        "observacoes_internas": "",
        "filtro_historico": "",
        "ultima_pasta_gerada": None,
        "confirm_delete_folder": "",
        "confirm_delete_file": "",
        "origem_dados_carregados": "",
        "alerta_vencimento_dias": 30,
        "painel_acao_msg": "",
        "busca_rapida": "",
        "filtro_status_central": "Todos",
        "rel_mes_referencia": datetime.now().strftime("%m"),
        "rel_tipo_atendimento": "Contrato ativo",
        "rel_nome_condominio": "",
        "rel_cnpj_condominio": "",
        "rel_endereco_condominio": "",
        "rel_representante": "",
        "rel_cpf_cnpj_representante": "",
        "rel_salvar_alteracoes_cadastro": False,
        "rel_ano_referencia": str(datetime.now().year),
        "rel_art_status": "Emitida",
        "rel_art_numero": "",
        "rel_art_inicio": "",
        "rel_art_fim": "",
        "rel_data_emissao": hoje_br(),
        "rel_epi_luvas_ca": "",
        "rel_epi_oculos_ca": "",
        "rel_epi_respirador_ca": "",
        "rel_epi_botas_ca": "",
        "rel_status_agua": "CONFORME",
        "rel_status_agua_select": "CONFORME",
        "rel_diagnostico": "",
        "rel_nbr_11238": "",
        "rel_nr_26": "",
        "rel_nr_06": "",
        "rel_analises_total": ANALISES_PADRAO,
    }
    garantir_campos_analises(st.session_state.get("rel_analises_total", ANALISES_PADRAO) if hasattr(st, "session_state") else ANALISES_PADRAO)
    for i in range(ANALISES_PADRAO):
        defaults[f"rel_analise_data_{i}"] = ""
        defaults[f"rel_analise_ph_{i}"] = ""
        defaults[f"rel_analise_cl_{i}"] = ""
        defaults[f"rel_analise_ct_{i}"] = ""
        defaults[f"rel_analise_alc_{i}"] = ""
        defaults[f"rel_analise_dc_{i}"] = ""
        defaults[f"rel_analise_cya_{i}"] = ""
        defaults[f"rel_analise_operador_{i}"] = ""
    for i in range(7):
        defaults[f"rel_dos_produto_{i}"] = ""
        defaults[f"rel_dos_lote_{i}"] = ""
        defaults[f"rel_dos_qtd_{i}"] = ""
        defaults[f"rel_dos_un_{i}"] = ""
        defaults[f"rel_dos_finalidade_{i}"] = ""
    for i in range(5):
        defaults[f"rel_rec_texto_{i}"] = ""
        defaults[f"rel_rec_prazo_{i}"] = ""
        defaults[f"rel_rec_resp_{i}"] = ""
        defaults[f"rel_obs_{i}"] = ""
    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def limpar_formulario():
    st.session_state.nome_condominio = ""
    st.session_state.cnpj_condominio = ""
    st.session_state.endereco_condominio = ""
    st.session_state.nome_sindico = ""
    st.session_state.cpf_sindico = ""
    st.session_state.valor_mensal = ""
    st.session_state.valor_aditivo = ""
    st.session_state.data_inicio = ""
    st.session_state.data_fim = ""
    st.session_state.data_assinatura = hoje_br()
    st.session_state.whatsapp_cliente = ""
    st.session_state.email_cliente = ""
    st.session_state.observacoes_internas = ""
    st.session_state.origem_dados_carregados = ""
    st.session_state.rel_tipo_atendimento = "Contrato ativo"
    st.session_state.rel_nome_condominio = ""
    st.session_state.rel_cnpj_condominio = ""
    st.session_state.rel_endereco_condominio = ""
    st.session_state.rel_representante = ""
    st.session_state.rel_cpf_cnpj_representante = ""
    st.session_state.rel_salvar_alteracoes_cadastro = False
    st.session_state.rel_mes_referencia = datetime.now().strftime("%m")
    st.session_state.rel_ano_referencia = str(datetime.now().year)
    st.session_state.rel_art_status = "Emitida"
    st.session_state.rel_art_numero = ""
    st.session_state.rel_art_inicio = ""
    st.session_state.rel_art_fim = ""
    st.session_state.rel_data_emissao = hoje_br()
    st.session_state.rel_epi_luvas_ca = ""
    st.session_state.rel_epi_oculos_ca = ""
    st.session_state.rel_epi_respirador_ca = ""
    st.session_state.rel_epi_botas_ca = ""
    st.session_state.rel_status_agua = "CONFORME"
    st.session_state.rel_status_agua_select = "CONFORME"
    st.session_state.rel_diagnostico = ""
    st.session_state.rel_nbr_11238 = ""
    st.session_state.rel_nr_26 = ""
    st.session_state.rel_nr_06 = ""
    st.session_state.rel_epi_luvas_status = "Conforme"
    st.session_state.rel_epi_oculos_status = "Conforme"
    st.session_state.rel_epi_respirador_status = "Conforme"
    st.session_state.rel_epi_botas_status = "Conforme"
    st.session_state.rel_epi_luvas_ca = ""
    st.session_state.rel_epi_oculos_ca = ""
    st.session_state.rel_epi_respirador_ca = ""
    st.session_state.rel_epi_botas_ca = ""
    total_analises_atual = int(st.session_state.get("rel_analises_total", ANALISES_PADRAO) or ANALISES_PADRAO)
    st.session_state.rel_analises_total = ANALISES_PADRAO
    for i in range(max(total_analises_atual, ANALISES_PADRAO)):
        st.session_state[f"rel_analise_data_{i}"] = ""
        st.session_state[f"rel_analise_ph_{i}"] = ""
        st.session_state[f"rel_analise_cl_{i}"] = ""
        st.session_state[f"rel_analise_ct_{i}"] = ""
        st.session_state[f"rel_analise_alc_{i}"] = ""
        st.session_state[f"rel_analise_dc_{i}"] = ""
        st.session_state[f"rel_analise_cya_{i}"] = ""
        st.session_state[f"rel_analise_operador_{i}"] = ""
    for i in range(7):
        st.session_state[f"rel_dos_produto_{i}"] = ""
        st.session_state[f"rel_dos_lote_{i}"] = ""
        st.session_state[f"rel_dos_qtd_{i}"] = ""
        st.session_state[f"rel_dos_un_{i}"] = ""
        st.session_state[f"rel_dos_finalidade_{i}"] = ""
    for i in range(5):
        st.session_state[f"rel_rec_texto_{i}"] = ""
        st.session_state[f"rel_rec_prazo_{i}"] = ""
        st.session_state[f"rel_rec_resp_{i}"] = ""
        st.session_state[f"rel_obs_{i}"] = ""


RASCUNHO_JSON_NAME = "rascunho_relatorio.json"


def salvar_rascunho_relatorio(pasta_condominio: Path):
    """Salva o estado atual do formulário de relatório como rascunho no JSON."""
    qtd = int(st.session_state.get("rel_analises_total", ANALISES_PADRAO) or ANALISES_PADRAO)
    rascunho = {
        "rel_nome_condominio": (st.session_state.get("rel_nome_condominio") or "").strip(),
        "rel_cnpj_condominio": (st.session_state.get("rel_cnpj_condominio") or "").strip(),
        "rel_endereco_condominio": (st.session_state.get("rel_endereco_condominio") or "").strip(),
        "rel_representante": (st.session_state.get("rel_representante") or "").strip(),
        "rel_cpf_cnpj_representante": (st.session_state.get("rel_cpf_cnpj_representante") or "").strip(),
        "rel_tipo_atendimento": (st.session_state.get("rel_tipo_atendimento") or "Contrato ativo"),
        "rel_mes_referencia": (st.session_state.get("rel_mes_referencia") or ""),
        "rel_ano_referencia": (st.session_state.get("rel_ano_referencia") or ""),
        "rel_data_emissao": (st.session_state.get("rel_data_emissao") or hoje_br()),
        "rel_art_status": (st.session_state.get("rel_art_status") or "Emitida"),
        "rel_art_numero": (st.session_state.get("rel_art_numero") or ""),
        "rel_art_inicio": (st.session_state.get("rel_art_inicio") or ""),
        "rel_art_fim": (st.session_state.get("rel_art_fim") or ""),
        "rel_status_agua": (st.session_state.get("rel_status_agua") or "CONFORME"),
        "rel_diagnostico": (st.session_state.get("rel_diagnostico") or ""),
        "rel_nbr_11238": (st.session_state.get("rel_nbr_11238") or ""),
        "rel_nr_26": (st.session_state.get("rel_nr_26") or ""),
        "rel_nr_06": (st.session_state.get("rel_nr_06") or ""),
        "rel_epi_luvas_status": (st.session_state.get("rel_epi_luvas_status") or "Conforme"),
        "rel_epi_luvas_ca": (st.session_state.get("rel_epi_luvas_ca") or ""),
        "rel_epi_oculos_status": (st.session_state.get("rel_epi_oculos_status") or "Conforme"),
        "rel_epi_oculos_ca": (st.session_state.get("rel_epi_oculos_ca") or ""),
        "rel_epi_respirador_status": (st.session_state.get("rel_epi_respirador_status") or "Conforme"),
        "rel_epi_respirador_ca": (st.session_state.get("rel_epi_respirador_ca") or ""),
        "rel_epi_botas_status": (st.session_state.get("rel_epi_botas_status") or "Conforme"),
        "rel_epi_botas_ca": (st.session_state.get("rel_epi_botas_ca") or ""),
        "rel_analises_total": qtd,
        "analises": [],
        "dosagens": obter_dosagens_ultimas_relatorio(),
        "observacoes": [(st.session_state.get(f"rel_obs_{i}") or "") for i in range(5)],
        "recomendacoes": [
            {
                "texto": (st.session_state.get(f"rel_rec_texto_{i}") or ""),
                "prazo": (st.session_state.get(f"rel_rec_prazo_{i}") or ""),
                "responsavel": (st.session_state.get(f"rel_rec_resp_{i}") or ""),
            }
            for i in range(5)
        ],
        "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    for i in range(qtd):
        rascunho["analises"].append({
            "data": (st.session_state.get(f"rel_analise_data_{i}") or ""),
            "ph": (st.session_state.get(f"rel_analise_ph_{i}") or ""),
            "cl": (st.session_state.get(f"rel_analise_cl_{i}") or ""),
            "ct": (st.session_state.get(f"rel_analise_ct_{i}") or ""),
            "alc": (st.session_state.get(f"rel_analise_alc_{i}") or ""),
            "dc": (st.session_state.get(f"rel_analise_dc_{i}") or ""),
            "cya": (st.session_state.get(f"rel_analise_cya_{i}") or ""),
            "operador": (st.session_state.get(f"rel_analise_operador_{i}") or ""),
        })
    caminho = pasta_condominio / RASCUNHO_JSON_NAME
    caminho.write_text(json.dumps(rascunho, ensure_ascii=False, indent=2), encoding="utf-8")
    return rascunho


def carregar_rascunho_relatorio(pasta_condominio: Path) -> dict | None:
    caminho = pasta_condominio / RASCUNHO_JSON_NAME
    if not caminho.exists():
        return None
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except Exception:
        return None


def aplicar_rascunho_no_formulario(rascunho: dict):
    """Restaura todos os campos do relatório a partir do rascunho salvo."""
    campos_simples = [
        "rel_nome_condominio", "rel_cnpj_condominio", "rel_endereco_condominio",
        "rel_representante", "rel_cpf_cnpj_representante", "rel_tipo_atendimento",
        "rel_mes_referencia", "rel_ano_referencia", "rel_data_emissao",
        "rel_art_status", "rel_art_numero", "rel_art_inicio", "rel_art_fim",
        "rel_status_agua", "rel_diagnostico", "rel_nbr_11238", "rel_nr_26", "rel_nr_06",
        "rel_epi_luvas_status", "rel_epi_luvas_ca", "rel_epi_oculos_status", "rel_epi_oculos_ca",
        "rel_epi_respirador_status", "rel_epi_respirador_ca", "rel_epi_botas_status", "rel_epi_botas_ca",
    ]
    for c in campos_simples:
        if c in rascunho:
            st.session_state[c] = rascunho[c]

    analises = rascunho.get("analises", [])
    qtd = max(len(analises), ANALISES_PADRAO)
    garantir_campos_analises(qtd)
    st.session_state.rel_analises_total = qtd
    for i, a in enumerate(analises):
        st.session_state[f"rel_analise_data_{i}"] = a.get("data", "")
        st.session_state[f"rel_analise_ph_{i}"] = a.get("ph", "")
        st.session_state[f"rel_analise_cl_{i}"] = a.get("cl", "")
        st.session_state[f"rel_analise_ct_{i}"] = a.get("ct", "")
        st.session_state[f"rel_analise_alc_{i}"] = a.get("alc", "")
        st.session_state[f"rel_analise_dc_{i}"] = a.get("dc", "")
        st.session_state[f"rel_analise_cya_{i}"] = a.get("cya", "")
        st.session_state[f"rel_analise_operador_{i}"] = a.get("operador", "")

    dosagens = rascunho.get("dosagens", [])
    for i in range(7):
        d = dosagens[i] if i < len(dosagens) else {}
        st.session_state[f"rel_dos_produto_{i}"] = d.get("produto", "")
        st.session_state[f"rel_dos_lote_{i}"] = d.get("fabricante_lote", "")
        st.session_state[f"rel_dos_qtd_{i}"] = d.get("quantidade", "")
        st.session_state[f"rel_dos_un_{i}"] = d.get("unidade", "")
        st.session_state[f"rel_dos_finalidade_{i}"] = d.get("finalidade", "")

    for i, obs in enumerate(rascunho.get("observacoes", [])[:5]):
        st.session_state[f"rel_obs_{i}"] = obs

    for i, rec in enumerate(rascunho.get("recomendacoes", [])[:5]):
        st.session_state[f"rel_rec_texto_{i}"] = rec.get("texto", "")
        st.session_state[f"rel_rec_prazo_{i}"] = rec.get("prazo", "")
        st.session_state[f"rel_rec_resp_{i}"] = rec.get("responsavel", "")


def excluir_rascunho_relatorio(pasta_condominio: Path):
    caminho = pasta_condominio / RASCUNHO_JSON_NAME
    if caminho.exists():
        caminho.unlink()


inicializar_campos()
sincronizar_relatorio_com_cadastro()

# Auto-restaurar rascunho ao iniciar sessão
if not st.session_state.get("_rascunho_restaurado"):
    nome_atual = (st.session_state.get("rel_nome_condominio") or st.session_state.get("nome_condominio") or "").strip()
    if nome_atual:
        pasta_rasc = GENERATED_DIR / slugify_nome(nome_atual)
        rasc = carregar_rascunho_relatorio(pasta_rasc)
        if rasc:
            aplicar_rascunho_no_formulario(rasc)
    st.session_state["_rascunho_restaurado"] = True

# Quando nome do condomínio do relatório muda, tenta restaurar rascunho automaticamente
_nome_rel_check = (st.session_state.get("rel_nome_condominio") or "").strip()
_ultimo_nome_rasc = st.session_state.get("_ultimo_nome_rasc_check", "")
if _nome_rel_check and _nome_rel_check != _ultimo_nome_rasc:
    st.session_state["_ultimo_nome_rasc_check"] = _nome_rel_check
    _pasta_check = GENERATED_DIR / slugify_nome(_nome_rel_check)
    if _pasta_check.exists():
        _rasc_auto = carregar_rascunho_relatorio(_pasta_check)
        if _rasc_auto:
            aplicar_rascunho_no_formulario(_rasc_auto)

pasta_formulario_atual = obter_pasta_atual_do_formulario()
if pasta_formulario_atual and pasta_formulario_atual.exists():
    dados_auto = carregar_dados_condominio(pasta_formulario_atual)
    if dados_auto:
        dosagens_preenchidas = any((st.session_state.get(f"rel_dos_produto_{i}") or "").strip() or (st.session_state.get(f"rel_dos_lote_{i}") or "").strip() or (st.session_state.get(f"rel_dos_qtd_{i}") or "").strip() or (st.session_state.get(f"rel_dos_un_{i}") or "").strip() or (st.session_state.get(f"rel_dos_finalidade_{i}") or "").strip() for i in range(7))
        if not dosagens_preenchidas:
            aplicar_dosagens_ultimas_no_relatorio(dados_auto)

# =========================================
# TOPO
# =========================================

logo = encontrar_logo()

col_top1, col_top2 = st.columns([1, 5])

with col_top1:
    if logo:
        st.image(str(logo), width=150)

with col_top2:
    st.markdown(
        f"""
        <div class="top-card">
            <div class="top-title">{APP_TITLE}</div>
            <div class="top-subtitle">
                Sistema profissional para geração automatizada de contrato e aditivo de RT
            </div>
            <div>
                <span class="info-badge">{RESPONSAVEL_TÉCNICO}</span>
                <span class="info-badge">{CRQ}</span>
                <span class="info-badge">Windows + Word + DOCX/PDF</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================
# SIDEBAR – CONFIG + HISTÓRICO
# =========================================

with st.sidebar:
    st.header("Histórico recente")

    st.number_input(
        "Lembrete de vencimento (dias)",
        min_value=1,
        max_value=180,
        step=1,
        key="alerta_vencimento_dias",
    )

    st.text_input(
        "Filtrar condomínio",
        key="filtro_historico",
        placeholder="Digite parte do nome...",
    )

    historico = listar_historico()
    filtro = st.session_state.filtro_historico.strip().lower()

    if filtro:
        historico = [h for h in historico if filtro in h["nome"].lower()]

    if not historico:
        st.caption("Nenhum histórico encontrado.")
    else:
        for item in historico:
            nome_cond = item["nome"]
            pasta = item["pasta"]
            arquivos = item["arquivos"]
            folder_key = chave_segura(str(pasta))
            status = status_vencimento(item["data_fim"], st.session_state.alerta_vencimento_dias)

            titulo = f"{nome_cond} ({item['total_arquivos']})"

            with st.expander(titulo, expanded=False):
                st.caption(str(pasta))
                st.markdown(
                    f"<span class='status-badge {status['css']}'>{status['rotulo']}</span>",
                    unsafe_allow_html=True,
                )
                st.caption(status["mensagem"])

                col_h1, col_h2 = st.columns(2)
                with col_h1:
                    if st.button(
                        "Carregar dados",
                        key=f"carregar_dados_{folder_key}",
                        use_container_width=True,
                    ):
                        dados_salvos = carregar_dados_condominio(pasta)
                        if dados_salvos:
                            aplicar_dados_no_formulario(dados_salvos)
                            st.success("Dados carregados no formulário.")
                            st.rerun()
                        else:
                            st.warning("Nenhum cadastro salvo encontrado para este condomínio.")

                with col_h2:
                    if st.button(
                        "Abrir pasta",
                        key=f"abrir_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        abrir_pasta_windows(pasta)

                if st.session_state.confirm_delete_folder == folder_key:
                    if st.button(
                        "Confirmar pasta",
                        key=f"confirmar_del_pasta_{folder_key}",
                        type="primary",
                        use_container_width=True,
                    ):
                        try:
                            deletar_pasta_condominio(pasta)
                            st.session_state.confirm_delete_folder = ""
                            st.success(f"Pasta de '{nome_cond}' excluída.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao excluir pasta: {e}")
                else:
                    if st.button(
                        "Excluir pasta",
                        key=f"pedir_del_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        st.session_state.confirm_delete_folder = folder_key
                        st.rerun()

                if st.session_state.confirm_delete_folder == folder_key:
                    st.markdown(
                        "<div class='confirm-box'><strong>Confirma excluir toda a pasta deste condomínio?</strong></div>",
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        "Cancelar exclusão da pasta",
                        key=f"cancelar_del_pasta_{folder_key}",
                        use_container_width=True,
                    ):
                        st.session_state.confirm_delete_folder = ""
                        st.rerun()

                st.markdown("**Arquivos:**")
                if not arquivos:
                    st.caption("Sem arquivos.")
                else:
                    for arq in arquivos:
                        caminho = arq["path"]
                        nome = arq["name"]
                        tipo_doc = arq["tipo_doc"]
                        tipo_ext = arq["tipo_ext"]
                        modificado = arq["modificado_em"]

                        file_key = chave_segura(str(caminho))

                        st.markdown(
                            f"**{tipo_doc} • {tipo_ext}**\n\n"
                            f"<div class='history-meta'>{nome}</div>"
                            f"<div class='history-meta'>Atualizado em: {modificado}</div>",
                            unsafe_allow_html=True,
                        )

                        col_a1, col_a2, col_a3 = st.columns([1.15, 1.0, 0.85])

                        with col_a1:
                            if st.button(
                                "Abrir arquivo",
                                key=f"abrir_arq_{file_key}",
                                use_container_width=True,
                            ):
                                abrir_arquivo_windows(caminho)

                        with col_a2:
                            if st.button(
                                "Abrir pasta",
                                key=f"abrir_pasta_arq_{file_key}",
                                use_container_width=True,
                            ):
                                abrir_pasta_windows(caminho.parent)

                        with col_a3:
                            if st.session_state.confirm_delete_file == file_key:
                                if st.button(
                                    "OK",
                                    key=f"confirmar_del_arq_{file_key}",
                                    type="primary",
                                    use_container_width=True,
                                ):
                                    try:
                                        deletar_arquivo_individual(caminho)
                                        st.session_state.confirm_delete_file = ""
                                        st.success(f"Arquivo excluído: {nome}")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Erro ao excluir arquivo: {e}")
                            else:
                                if st.button(
                                    "🗑️",
                                    key=f"pedir_del_arq_{file_key}",
                                    help="Excluir somente este arquivo",
                                    use_container_width=True,
                                ):
                                    st.session_state.confirm_delete_file = file_key
                                    st.rerun()

                        if st.session_state.confirm_delete_file == file_key:
                            st.markdown(
                                "<div class='confirm-box'>Excluir este arquivo?</div>",
                                unsafe_allow_html=True,
                            )
                            if st.button(
                                "Cancelar",
                                key=f"cancelar_del_arq_{file_key}",
                                use_container_width=True,
                            ):
                                st.session_state.confirm_delete_file = ""
                                st.rerun()

                        st.markdown("---")

# =========================================
# MODO DE OPERAÇÃO — TELA DE ENTRADA
# =========================================

# PIN padrão — altere aqui para trocar o PIN do operador
PIN_OPERADOR = "5010"

# Inicializa o modo se não estiver definido
if "modo_atual" not in st.session_state:
    st.session_state["modo_atual"] = "entrada"

# Compatibilidade com seletor antigo (Modo Campo ainda usa st.radio internamente)
_modo_interno = st.session_state.get("modo_atual", "entrada")

# ---- TELA DE ENTRADA ----
if _modo_interno == "entrada":
    st.markdown("""
    <style>
    .entrada-card {
        border: 1px solid rgba(20,85,160,0.15);
        border-radius: 20px;
        padding: 32px 24px;
        background: linear-gradient(135deg, #ffffff 0%, #f4f9ff 100%);
        box-shadow: 0 6px 24px rgba(10,50,100,0.08);
        margin: 12px 0;
        text-align: center;
    }
    .entrada-title { font-size: 1.3rem; font-weight: 700; color: #0d3d75; margin-bottom: 6px; }
    .entrada-sub { font-size: 0.9rem; color: #5d7288; margin-bottom: 20px; }
    .entrada-link { font-size: 0.75rem; color: #aab8c8; margin-top: 18px; }
    </style>
    """, unsafe_allow_html=True)

    col_e1, col_e2, col_e3 = st.columns([1, 2, 1])
    with col_e2:
        st.markdown('<div class="entrada-card">', unsafe_allow_html=True)

        # Seleção de empresa
        _empresa_sel = st.radio(
            "Empresa",
            ["🔵 Aqua Gestão", "⭐ Bem Star Piscinas"],
            key="empresa_selecionada",
            horizontal=True,
            label_visibility="collapsed",
        )
        _eh_bem_star = "Bem Star" in _empresa_sel
        st.session_state["empresa_ativa"] = "bem_star" if _eh_bem_star else "aqua_gestao"

        if _eh_bem_star:
            # Logo Bem Star
            _logo_bs = None
            for _lp in LOGO_BEM_STAR_CANDIDATOS:
                if _lp.exists():
                    _logo_bs = _lp
                    break
            if _logo_bs:
                st.image(str(_logo_bs), width=180)
            st.markdown('<div class="entrada-title">Bem Star Piscinas</div>', unsafe_allow_html=True)
            st.markdown('<div class="entrada-sub">Manutenção e Tratamento de Piscinas<br>CNPJ: 26.799.958/0001-88</div>', unsafe_allow_html=True)
        else:
            _logo_aq = encontrar_logo()
            if _logo_aq:
                st.image(str(_logo_aq), width=160)
            st.markdown('<div class="entrada-title">Aqua Gestão</div>', unsafe_allow_html=True)
            st.markdown('<div class="entrada-sub">Gestão de Água<br>Controle Técnico de Piscinas<br>Thyago Fernando da Silveira | CRQ-MG 2ª Região</div>', unsafe_allow_html=True)

        if st.button("📱 Acessar como Operador", type="primary", use_container_width=True):
            st.session_state["modo_atual"] = "operador"
            st.rerun()

        # Acesso ao escritório — com PIN administrativo
        st.markdown('<div class="entrada-link">Acesso administrativo</div>', unsafe_allow_html=True)
        if st.button("·  ·  ·", use_container_width=False, key="btn_escritorio_oculto"):
            st.session_state["mostrar_pin_admin"] = True

        if st.session_state.get("mostrar_pin_admin"):
            pin_admin = st.text_input("PIN administrativo", type="password",
                key="pin_admin_input", placeholder="Digite o PIN", label_visibility="collapsed")
            if st.button("Entrar", key="btn_pin_admin_ok", use_container_width=True):
                if pin_admin == "@Anajullya10":
                    st.session_state["modo_atual"] = "escritorio"
                    st.session_state["mostrar_pin_admin"] = False
                    st.rerun()
                else:
                    st.error("PIN incorreto.")

        st.markdown("</div>", unsafe_allow_html=True)

    st.stop()

# Alias para compatibilidade
modo = "Modo Escritório" if _modo_interno == "escritorio" else (
    "📱 Modo Operador (Campo / Celular)" if _modo_interno == "operador" else "Modo Escritório"
)

# Botão de voltar à tela inicial (aparece em todos os modos)
if _modo_interno in ("escritorio", "operador"):
    if st.button("← Voltar à tela inicial", key="btn_voltar_inicio"):
        st.session_state["modo_atual"] = "entrada"
        if _modo_interno == "operador":
            st.session_state["op_pin_ok"] = False
        st.rerun()

# =========================================
# MODO OPERADOR — LANÇAMENTO DE CAMPO
# =========================================

if modo == "📱 Modo Operador (Campo / Celular)":

    st.markdown("""
    <style>
    section[data-testid="stSidebar"] { display: none !important; }
    .main .block-container { padding: 0.5rem 0.8rem 2rem !important; max-width: 100% !important; }
    .op-card {
        border: 1px solid rgba(20,85,160,0.18);
        border-radius: 16px;
        padding: 16px;
        background: #ffffff;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .op-title { font-size: 1.2rem; font-weight: 700; color: #0d3d75; margin-bottom: 2px; }
    .op-sub { font-size: 0.82rem; color: #5d7288; margin-bottom: 10px; }
    .op-salvo {
        border: 1px solid rgba(30,140,70,0.3);
        border-radius: 12px;
        padding: 12px 14px;
        background: rgba(30,140,70,0.08);
        color: #1a6e3a;
        font-size: 0.95rem;
        margin-top: 8px;
    }
    .stTextInput input, .stTextArea textarea {
        font-size: 1rem !important;
        min-height: 48px !important;
        border-radius: 10px !important;
    }
    .stButton > button {
        min-height: 52px !important;
        font-size: 1.05rem !important;
        border-radius: 12px !important;
    }
    .stTextInput label, .stSelectbox label, .stTextArea label {
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: #1a3a5c !important;
    }
    .element-container { margin-bottom: 6px !important; }
    .pin-box {
        border: 2px solid rgba(20,85,160,0.25);
        border-radius: 20px;
        padding: 32px 20px;
        background: #ffffff;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin: 20px 0;
    }
    </style>
    """, unsafe_allow_html=True)

    # ---- TELA DE PIN ----
    if not st.session_state.get("op_pin_ok"):
        st.markdown('<div class="pin-box">', unsafe_allow_html=True)
        st.markdown("### 🔐 Área do Operador")
        st.markdown("**Aqua Gestão – Controle Técnico de Piscinas**")
        st.markdown("Digite o PIN para acessar o lançamento de campo.")
        pin_digitado = st.text_input("PIN", type="password", key="op_pin_input",
            placeholder="Digite o PIN", label_visibility="collapsed", max_chars=20)
        if st.button("Entrar", type="primary", use_container_width=True):
            op_dados = validar_pin_operador(pin_digitado.strip())
            if op_dados:
                st.session_state["op_pin_ok"] = True
                st.session_state["op_dados_atual"] = op_dados
                st.rerun()
            else:
                st.error("PIN incorreto. Tente novamente.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    # Dados do operador logado
    _op_atual = st.session_state.get("op_dados_atual", {"nome": "Operador", "acesso_total": True, "condomínios": []})
    _op_nome_logado = _op_atual.get("nome", "Operador")
    _op_acesso_total = _op_atual.get("acesso_total", False)
    _op_conds_permitidos = _op_atual.get("condomínios", [])

    if st.button("🔒 Sair / Trocar operador", use_container_width=False):
        st.session_state["op_pin_ok"] = False
        st.session_state.pop("op_dados_atual", None)
        st.rerun()

    st.markdown('<div class="op-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="op-title">📱 Lançamento de Campo — {_op_nome_logado}</div>', unsafe_allow_html=True)
    st.markdown('<div class="op-sub">Aqua Gestão – Controle Técnico de Piscinas | Thyago Fernando da Silveira</div>', unsafe_allow_html=True)

    _salvo = st.session_state.pop("op_salvo_sucesso", None)
    if _salvo:
        st.markdown(f"""
        <div class="op-salvo">
            ✅ <strong>Lançamento salvo!</strong><br>
            Condomínio: {_salvo['nome']}<br>
            Data: {_salvo['data']}<br>
            Operador: {_salvo['operador']}<br>
            Total este mês: {_salvo['total']}
        </div>
        """, unsafe_allow_html=True)
        # Botão de relatório do lançamento recém salvo
        _ult_lanc = st.session_state.get("_op_ultimo_lancamento")
        if _ult_lanc:
            nome_arq = limpar_nome_arquivo(f"Relatorio_Visita_{_salvo['nome']}_{_salvo['data'].replace('/','')}")
            with st.spinner("Gerando PDF..."):
                try:
                    pdf_bytes = gerar_pdf_relatorio_visita(_ult_lanc, _salvo["nome"])
                    st.download_button(
                        "📄 Baixar PDF desta visita",
                        data=pdf_bytes,
                        file_name=f"{nome_arq}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="btn_dl_relatorio_visita",
                    )
                    st.caption("Baixe e compartilhe diretamente pelo WhatsApp.")
                except Exception as _e:
                    st.warning(f"PDF não gerado: {_e}. Baixando versão HTML como alternativa.")
                    html_rel = gerar_html_relatorio_visita(_ult_lanc, _salvo["nome"])
                    st.download_button(
                        "📄 Baixar relatório (HTML)",
                        data=html_rel.encode("utf-8"),
                        file_name=f"{nome_arq}.html",
                        mime="text/html",
                        use_container_width=True,
                        key="btn_dl_relatorio_visita_html",
                    )

    # Busca clientes do Google Sheets primeiro, fallback para pasta local
    @st.cache_data(ttl=60)
    def _buscar_clientes_sheets():
        return sheets_listar_clientes()

    clientes_sheets = _buscar_clientes_sheets()

    # Combina clientes do Sheets com os locais
    pastas_disponiveis = sorted([
        p for p in GENERATED_DIR.iterdir() if p.is_dir()
    ], key=lambda p: p.name) if GENERATED_DIR.exists() else []

    opcoes_cond_local = []
    for p in pastas_disponiveis:
        dados_c = carregar_dados_condominio(p)
        nome_ex = dados_c.get("nome_condominio", humanizar_nome_pasta(p.name)) if dados_c else humanizar_nome_pasta(p.name)
        opcoes_cond_local.append(nome_ex)

    # Une as duas listas sem duplicar
    opcoes_cond_todas = list(dict.fromkeys(clientes_sheets + opcoes_cond_local))

    # Operador logado via PIN — nome vem do cadastro, não digitado
    _op_nome_logado_disp = _op_nome_logado if _op_nome_logado != "Operador" else ""
    op_operador = st.text_input(
        "Seu nome (operador)",
        key="op_operador",
        value=_op_nome_logado_disp,
        placeholder="Ex.: João Silva",
        help="Preenchido automaticamente pelo seu PIN de acesso."
    )

    # Filtra condomínios pelo PIN do operador logado
    if _op_acesso_total or not _op_conds_permitidos or "TODOS" in _op_conds_permitidos:
        opcoes_cond = opcoes_cond_todas
    else:
        opcoes_cond = [c for c in opcoes_cond_todas if any(
            perm.lower().strip() in c.lower() or c.lower() in perm.lower().strip()
            for perm in _op_conds_permitidos
        )]
        if opcoes_cond:
            st.caption(f"✅ Acesso liberado para {len(opcoes_cond)} condomínio(s).")
        else:
            st.warning("Nenhum condomínio disponível para seu acesso. Contate o administrador.")

    op_usar_novo = st.checkbox("Lançar para local não cadastrado", key="op_novo_cond")
    if op_usar_novo:
        op_nome_cond = st.text_input("Nome do local", key="op_nome_livre", placeholder="Ex.: Residencial Aquarela")
    else:
        if opcoes_cond:
            op_nome_cond = st.selectbox("Selecione o condomínio", opcoes_cond, key="op_sel_cond")
        else:
            st.warning("Nenhum condomínio disponível. Verifique seu nome ou contate o administrador.")
            op_nome_cond = ""

    def _fmt_data_op():
        v = apenas_digitos(st.session_state.get("op_data_visita", ""))[:8]
        if len(v) <= 2:
            st.session_state["op_data_visita"] = v
        elif len(v) <= 4:
            st.session_state["op_data_visita"] = f"{v[:2]}/{v[2:]}"
        else:
            st.session_state["op_data_visita"] = f"{v[:2]}/{v[2:4]}/{v[4:]}"

    st.text_input("Data (digite só números: ddmmaaaa)",
        key="op_data_visita", placeholder="06/04/2026", on_change=_fmt_data_op)

    st.markdown("</div>", unsafe_allow_html=True)

    if op_nome_cond:

        # ── Piscinas deste condomínio ─────────────────────────────────────────
        # Carrega configuração de piscinas salva ou usa padrão
        _pasta_cond_op = GENERATED_DIR / slugify_nome(op_nome_cond.strip())
        _dados_cond_op = carregar_dados_condominio(_pasta_cond_op) if _pasta_cond_op.exists() else {}
        _piscinas_config = _dados_cond_op.get("piscinas", ["Piscina Adulto"])

        # Admin pode definir piscinas pelo painel — operador vê as já configuradas
        with st.expander("🏊 Piscinas deste condomínio", expanded=False):
            st.caption("Configure as piscinas deste condomínio. Salvo automaticamente.")
            _pisc_adulto  = st.checkbox("Piscina Adulto",   value="Piscina Adulto"   in _piscinas_config, key="op_pisc_adulto")
            _pisc_infant  = st.checkbox("Piscina Infantil", value="Piscina Infantil" in _piscinas_config, key="op_pisc_infantil")
            _pisc_family  = st.checkbox("Piscina Family",   value="Piscina Family"   in _piscinas_config, key="op_pisc_family")
            _pisc_outra_check = st.checkbox("Outra piscina", value=any(p not in ["Piscina Adulto","Piscina Infantil","Piscina Family"] for p in _piscinas_config), key="op_pisc_outra_check")
            _pisc_outra_nome = ""
            if _pisc_outra_check:
                _outra_default = next((p for p in _piscinas_config if p not in ["Piscina Adulto","Piscina Infantil","Piscina Family"]), "")
                _pisc_outra_nome = st.text_input("Nome da outra piscina", value=_outra_default, key="op_pisc_outra_nome", placeholder="Ex.: Piscina Olímpica")
            _piscinas_ativas = []
            if _pisc_adulto: _piscinas_ativas.append("Piscina Adulto")
            if _pisc_infant: _piscinas_ativas.append("Piscina Infantil")
            if _pisc_family: _piscinas_ativas.append("Piscina Family")
            if _pisc_outra_check and _pisc_outra_nome.strip(): _piscinas_ativas.append(_pisc_outra_nome.strip())
            if not _piscinas_ativas: _piscinas_ativas = ["Piscina Adulto"]
            if st.button("💾 Salvar configuração de piscinas", key="btn_salvar_piscinas"):
                _pasta_cond_op.mkdir(parents=True, exist_ok=True)
                _dados_upd = carregar_dados_condominio(_pasta_cond_op) or {}
                _dados_upd["piscinas"] = _piscinas_ativas
                _dados_upd["nome_condominio"] = op_nome_cond.strip()
                salvar_dados_condominio(_pasta_cond_op, _dados_upd)
                st.success(f"✅ Piscinas salvas: {', '.join(_piscinas_ativas)}")
                st.rerun()
        piscinas_ativas = _piscinas_ativas

        # Limpa campos SE houver limpeza pendente
        if st.session_state.pop("op_limpar_campos", False):
            for pisc in ["adulto","infantil","family","outra"]:
                for k in ["ph","crl","ct","alc","dc","cya"]:
                    st.session_state[f"op_{pisc}_{k}"] = ""
            for i in range(5):
                for s in ["prod","qtd","un","fin"]:
                    st.session_state[f"op_dos_{s}_{i}"] = ""
            st.session_state["op_obs_campo"] = ""
            st.session_state["op_problemas"] = ""

        def _num_op(chave, label, placeholder, quinzenal=False):
            lbl = f"{label} ⏱ 15d" if quinzenal else label
            v = st.text_input(lbl, key=chave, placeholder=placeholder,
                help="Medição quinzenal — preencha somente nas visitas de medição completa." if quinzenal else None)
            return re.sub(r"[^0-9.,]", "", v).replace(",", ".")

        with st.expander("📋 Faixas de referência", expanded=False):
            st.markdown("""
| Parâmetro | Faixa ideal |
|---|---|
| pH | 7,2 – 7,8 |
| CRL mg/L | 0,5 – 3,0 |
| Alcalinidade mg/L | 80 – 120 |
| Dureza DC mg/L | 150 – 300 |
| CYA mg/L | 30 – 50 |
            """)

        # ── Parâmetros por piscina ────────────────────────────────────────────
        op_piscinas_dados = []
        _slug_map = {"Piscina Adulto":"adulto","Piscina Infantil":"infantil","Piscina Family":"family"}

        for pisc_nome in piscinas_ativas:
            pisc_slug = _slug_map.get(pisc_nome, slugify_nome(pisc_nome)[:12])
            st.markdown('<div class="op-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="op-title">🧪 {pisc_nome}</div>', unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            with c1:
                p_ph  = _num_op(f"op_{pisc_slug}_ph",  "pH", "ex: 7.4")
                p_alc = _num_op(f"op_{pisc_slug}_alc", "Alcalinidade mg/L", "ex: 95",  quinzenal=True)
                p_dc  = _num_op(f"op_{pisc_slug}_dc",  "Dureza DC mg/L",   "ex: 200", quinzenal=True)
            with c2:
                p_crl = _num_op(f"op_{pisc_slug}_crl", "CRL mg/L",            "ex: 1.5")
                p_ct  = _num_op(f"op_{pisc_slug}_ct",  "Cloro Total CT mg/L", "ex: 1.8")
                p_cya = _num_op(f"op_{pisc_slug}_cya", "CYA mg/L",            "ex: 40",  quinzenal=True)

            # Alertas em tempo real
            for val, mn, mx, rot in [
                (p_ph, 7.2, 7.8, "pH"), (p_crl, 0.5, 3.0, "CRL"),
                (p_alc, 80, 120, "Alcalinidade"), (p_dc, 150, 300, "Dureza DC"),
                (p_cya, 30, 50, "CYA"),
            ]:
                v = valor_float(val)
                if v is not None:
                    st.markdown(f"{'⚠️' if v < mn or v > mx else '✅'} **{rot}: {v}** {'— fora da faixa' if v < mn or v > mx else '— conforme'}")

            p_cloraminas = None
            v_crl2 = valor_float(p_crl); v_ct2 = valor_float(p_ct)
            if v_crl2 is not None and v_ct2 is not None:
                p_cloraminas = round(max(v_ct2 - v_crl2, 0), 2)
                st.markdown(f"{'⚠️' if p_cloraminas > 0.2 else '✅'} **Cloraminas: {p_cloraminas} mg/L**")

            op_piscinas_dados.append({
                "nome": pisc_nome,
                "ph": p_ph, "cloro_livre": p_crl, "cloro_total": p_ct,
                "cloraminas": str(p_cloraminas) if p_cloraminas is not None else "",
                "alcalinidade": p_alc, "dureza": p_dc, "cianurico": p_cya,
            })

            # ── Sugestões de dosagem em tempo real ───────────────────────────
            _v_ph  = valor_float(p_ph)
            _v_crl = valor_float(p_crl)
            _v_alc = valor_float(p_alc)
            _v_dc  = valor_float(p_dc)
            _v_cya = valor_float(p_cya)
            _tem_params = any(v is not None for v in [_v_ph, _v_crl, _v_alc, _v_dc, _v_cya])

            if _tem_params:
                # Busca volume m³ do condomínio no Sheets
                _vol_m3 = st.session_state.get(f"_vol_m3_{slugify_nome(op_nome_cond.strip())}", 0.0)
                if not _vol_m3:
                    try:
                        _clientes_vol = sheets_listar_clientes_completo()
                        for _cv in _clientes_vol:
                            if _cv["nome"].lower().strip() == op_nome_cond.strip().lower():
                                # Busca volume da planilha (col D = Volume_m3)
                                _sh_vol = conectar_sheets()
                                if _sh_vol:
                                    _aba_vol = _sh_vol.worksheet("👥 Clientes")
                                    _rows_vol = _aba_vol.get_all_values()
                                    for _rv in _rows_vol:
                                        if len(_rv) > 3 and _cv["nome"].lower() in str(_rv[2]).lower():
                                            try:
                                                _vol_m3 = float(str(_rv[3]).replace(",",".").strip() or 0)
                                            except Exception:
                                                _vol_m3 = 0.0
                                            break
                                break
                        st.session_state[f"_vol_m3_{slugify_nome(op_nome_cond.strip())}"] = _vol_m3
                    except Exception:
                        _vol_m3 = 0.0

                # Usa volume específico da piscina se disponível
                    _vol_pisc = 0.0
                    _slug_map2 = {"Piscina Adulto":"vol_adulto","Piscina Infantil":"vol_infantil","Piscina Family":"vol_family"}
                    _vol_key = _slug_map2.get(pisc_nome, "")
                    if _vol_key and _vol_m3:
                        # Busca volume individual
                        try:
                            _clv = sheets_listar_clientes_completo()
                            for _cv2 in _clv:
                                if _cv2["nome"].lower().strip() == op_nome_cond.strip().lower():
                                    _vol_pisc = float(_cv2.get(_vol_key, 0) or 0)
                                    break
                        except Exception:
                            pass
                    _vol_usar = _vol_pisc if _vol_pisc > 0 else _vol_m3

                    if _vol_usar > 0:
                        _sugestoes = calcular_sugestoes_dosagem(
                            ph=_v_ph, crl=_v_crl, alc=_v_alc, dc=_v_dc, cya=_v_cya,
                            volume_m3=_vol_usar
                        )
                    if _sugestoes:
                        st.markdown(f"**💊 Sugestões para {pisc_nome} ({_vol_m3:.0f} m³):**")
                        for _s in _sugestoes:
                            _icon = "🔴" if _s["prioridade"] == 1 else ("🟡" if _s["prioridade"] == 2 else "🔵")
                            if _s["quantidade"] and _s["quantidade"] > 0:
                                st.markdown(f"{_icon} **{_s['produto']}** — **{_s['quantidade']} {_s['unidade']}**")
                                st.caption(f"↳ {_s['acao']}")
                            else:
                                st.markdown(f"{_icon} **{_s['produto']}** — {_s['acao']}")
                            with st.expander("ℹ️ Base técnica", expanded=False):
                                st.caption(_s["justificativa"])
                                st.caption(f"📚 {_s.get('norma','')}")
                    else:
                        st.success("✅ Todos os parâmetros dentro da faixa ideal.")
                else:
                    st.caption("⚠️ Volume m³ não cadastrado — adicione na planilha para calcular doses.")

            st.markdown("</div>", unsafe_allow_html=True)

        # Compatibilidade com código legado (usa dados da primeira piscina)
        op_ph  = op_piscinas_dados[0]["ph"]        if op_piscinas_dados else ""
        op_crl = op_piscinas_dados[0]["cloro_livre"] if op_piscinas_dados else ""
        op_ct  = op_piscinas_dados[0]["cloro_total"] if op_piscinas_dados else ""
        op_alc = op_piscinas_dados[0]["alcalinidade"] if op_piscinas_dados else ""
        op_dc  = op_piscinas_dados[0]["dureza"]      if op_piscinas_dados else ""
        op_cya = op_piscinas_dados[0]["cianurico"]   if op_piscinas_dados else ""
        op_cloraminas = valor_float(op_piscinas_dados[0]["cloraminas"]) if op_piscinas_dados else None

        # ── Dosagens ──────────────────────────────────────────────────────────
        st.markdown('<div class="op-card">', unsafe_allow_html=True)
        st.markdown('<div class="op-title">⚗️ Dosagens aplicadas</div>', unsafe_allow_html=True)
        op_dosagens = []
        for i in range(5):
            with st.expander(f"Produto {i+1}", expanded=(i == 0)):
                d1, d2 = st.columns([2, 1])
                prod = d1.text_input("Produto", key=f"op_dos_prod_{i}", label_visibility="collapsed", placeholder="Nome do produto")
                qtd  = d2.text_input("Qtd", key=f"op_dos_qtd_{i}", label_visibility="collapsed", placeholder="Qtd")
                d3, d4 = st.columns([1, 2])
                un   = d3.text_input("Un", key=f"op_dos_un_{i}", label_visibility="collapsed", placeholder="kg/L")
                fin  = d4.text_input("Finalidade", key=f"op_dos_fin_{i}", label_visibility="collapsed", placeholder="Finalidade")
                if prod.strip():
                    op_dosagens.append({"produto": prod.strip(), "fabricante_lote": "", "quantidade": qtd.strip(), "unidade": un.strip(), "finalidade": fin.strip()})
        st.markdown("</div>", unsafe_allow_html=True)

        # ── Fotos categorizadas ───────────────────────────────────────────────
        st.markdown('<div class="op-card">', unsafe_allow_html=True)
        st.markdown('<div class="op-title">📸 Fotos da visita</div>', unsafe_allow_html=True)
        st.caption("Faça upload de cada foto na categoria correta.")

        op_fotos_antes  = st.file_uploader("🔵 Antes do tratamento", type=["jpg","jpeg","png","webp","heic"], accept_multiple_files=True, key="op_fotos_antes")
        op_fotos_depois = st.file_uploader("🟢 Depois do tratamento", type=["jpg","jpeg","png","webp","heic"], accept_multiple_files=True, key="op_fotos_depois")
        op_fotos_cmaq   = st.file_uploader("🔧 Casa de máquinas", type=["jpg","jpeg","png","webp","heic"], accept_multiple_files=True, key="op_fotos_cmaq")

        # Preview
        _todas_fotos_preview = [("Antes", op_fotos_antes), ("Depois", op_fotos_depois), ("Casa máq.", op_fotos_cmaq)]
        for _cat, _flist in _todas_fotos_preview:
            if _flist:
                st.caption(f"**{_cat}:**")
                _cols = st.columns(min(len(_flist), 3))
                for _i, _f in enumerate(_flist):
                    with _cols[_i % 3]:
                        st.image(_f, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        # ── Problemas / Ocorrências ───────────────────────────────────────────
        st.markdown('<div class="op-card">', unsafe_allow_html=True)
        st.markdown('<div class="op-title">⚠️ Problemas / Ocorrências</div>', unsafe_allow_html=True)
        op_problemas = st.text_area("Problemas", key="op_problemas", height=80,
            label_visibility="collapsed",
            placeholder="Ex.: Bomba com ruído, vazamento na casa de máquinas, pH instável, equipamento quebrado...")
        st.markdown("</div>", unsafe_allow_html=True)

        # ── Observação geral ──────────────────────────────────────────────────
        st.markdown('<div class="op-card">', unsafe_allow_html=True)
        st.markdown('<div class="op-title">📝 Observação geral</div>', unsafe_allow_html=True)
        op_obs = st.text_area("Obs", key="op_obs_campo", height=80,
            label_visibility="collapsed", placeholder="Ex.: condições gerais da água, recomendações...")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="op-card">', unsafe_allow_html=True)
        if st.button("💾 Salvar lançamento", type="primary", use_container_width=True):
            data_vis = (st.session_state.get("op_data_visita") or "").strip()
            if not op_nome_cond.strip():
                st.error("Informe o condomínio.")
            elif not data_vis:
                st.error("Informe a data da visita.")
            else:
                pasta_op = GENERATED_DIR / slugify_nome(op_nome_cond.strip())
                pasta_op.mkdir(parents=True, exist_ok=True)
                pasta_fotos_op = pasta_op / "fotos_campo"
                pasta_fotos_op.mkdir(exist_ok=True)
                ts_f = datetime.now().strftime("%Y%m%d_%H%M%S")
                mes_ano = datetime.now().strftime("%Y-%m")

                import base64 as _b64

                def _salvar_categoria(lista_uploads, categoria):
                    """Salva fotos localmente, no Drive e retorna bytes para PDF."""
                    nomes = []; ids = []; b64s = []
                    for idx_ff, foto_ff in enumerate(lista_uploads or [], 1):
                        nome_ff = limpar_nome_arquivo(f"{categoria}_{ts_f}_{idx_ff}_{foto_ff.name}")
                        foto_bytes = bytes(foto_ff.getbuffer())
                        # Salva localmente
                        with open(pasta_fotos_op / nome_ff, "wb") as ff:
                            ff.write(foto_bytes)
                        nomes.append(nome_ff)
                        # Guarda base64 para PDF — com rotação EXIF automática
                        try:
                            from PIL import Image as _PILImg, ImageOps as _IOps
                            import io as _io
                            _img = _PILImg.open(_io.BytesIO(foto_bytes))
                            # Corrige rotação EXIF (fotos tiradas na vertical pelo celular)
                            _img = _IOps.exif_transpose(_img)
                            # Reduz para não pesar mas mantém qualidade
                            _img.thumbnail((1200, 1200), _PILImg.LANCZOS)
                            _buf = _io.BytesIO()
                            _img.convert("RGB").save(_buf, format="JPEG", quality=82)
                            b64s.append(_b64.b64encode(_buf.getvalue()).decode("utf-8"))
                        except Exception:
                            b64s.append(_b64.b64encode(foto_bytes).decode("utf-8"))
                        # Upload para o Google Drive (async best-effort)
                        try:
                            drive_id = drive_upload_foto(
                                arquivo_bytes=foto_bytes,
                                nome_arquivo=f"{categoria}_{nome_ff}",
                                nome_condominio=op_nome_cond.strip(),
                                mes_ano=mes_ano,
                            )
                            if drive_id:
                                ids.append(drive_id)
                        except Exception:
                            pass
                    return nomes, ids, b64s

                fotos_antes_nomes,  fotos_antes_ids,  fotos_antes_b64  = _salvar_categoria(op_fotos_antes,  "antes")
                fotos_depois_nomes, fotos_depois_ids, fotos_depois_b64 = _salvar_categoria(op_fotos_depois, "depois")
                fotos_cmaq_nomes,   fotos_cmaq_ids,   fotos_cmaq_b64   = _salvar_categoria(op_fotos_cmaq,   "cmaq")

                fotos_salvas_op = fotos_antes_nomes + fotos_depois_nomes + fotos_cmaq_nomes
                fotos_drive_ids = fotos_antes_ids   + fotos_depois_ids   + fotos_cmaq_ids

                dados_ex = carregar_dados_condominio(pasta_op) or {}
                lancamento = {
                    "data": data_vis, "operador": op_operador.strip(),
                    "ph": op_ph, "cloro_livre": op_crl, "cloro_total": op_ct,
                    "cloraminas": str(op_cloraminas) if op_cloraminas is not None else "",
                    "alcalinidade": op_alc, "dureza": op_dc, "cianurico": op_cya,
                    "piscinas": op_piscinas_dados,
                    "problemas": op_problemas.strip(),
                    "observacao": op_obs.strip(), "dosagens": op_dosagens,
                    "fotos": fotos_salvas_op,
                    "fotos_antes": fotos_antes_nomes,
                    "fotos_depois": fotos_depois_nomes,
                    "fotos_cmaq": fotos_cmaq_nomes,
                    "fotos_drive_ids": fotos_drive_ids,
                    "fotos_antes_ids": fotos_antes_ids,
                    "fotos_depois_ids": fotos_depois_ids,
                    "fotos_cmaq_ids": fotos_cmaq_ids,
                    "fotos_antes_b64": fotos_antes_b64,
                    "fotos_depois_b64": fotos_depois_b64,
                    "fotos_cmaq_b64": fotos_cmaq_b64,
                    "condominio": op_nome_cond.strip(),
                    "salvo_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                }
                pendentes = dados_ex.get("lancamentos_campo", [])
                pendentes.append(lancamento)
                dados_ex["lancamentos_campo"] = pendentes
                dados_ex["nome_condominio"] = dados_ex.get("nome_condominio", op_nome_cond.strip())
                if op_dosagens:
                    dados_ex["dosagens_ultimas"] = (op_dosagens + [{"produto":"","fabricante_lote":"","quantidade":"","unidade":"","finalidade":""}]*7)[:7]
                salvar_dados_condominio(pasta_op, dados_ex)

                # Salva também no Google Sheets
                ok_sheets = sheets_salvar_lancamento_campo(lancamento, op_nome_cond.strip())
                if not ok_sheets:
                    erro_sh = st.session_state.get("_sheets_ultimo_erro", "")
                    if erro_sh:
                        st.warning(f"⚠️ Salvo localmente, mas falhou no Google Sheets.\n\nDiagnóstico:\n```\n{erro_sh[:600]}\n```")
                    else:
                        st.warning("⚠️ Salvo localmente, mas não foi possível enviar ao Google Sheets. Verifique a conexão.")
                st.session_state["op_salvo_sucesso"] = {
                    "nome": op_nome_cond, "data": data_vis,
                    "operador": op_operador.strip() or "Não informado",
                    "total": len(pendentes),
                }
                # Guarda último lançamento para gerar relatório
                st.session_state["_op_ultimo_lancamento"] = lancamento
                # Sinaliza limpeza para o próximo rerun — não toca nos widgets agora
                st.session_state["op_limpar_campos"] = True
                st.rerun()

        pasta_hc = GENERATED_DIR / slugify_nome(op_nome_cond.strip()) if op_nome_cond.strip() else None
        if pasta_hc and pasta_hc.exists():
            dados_hc = carregar_dados_condominio(pasta_hc)
            pend_hc = (dados_hc or {}).get("lancamentos_campo", [])
            if pend_hc:
                st.markdown(f"**{len(pend_hc)} lançamento(s) registrado(s):**")
                for lc in reversed(pend_hc[-3:]):
                    ft = f" | 📸 {len(lc.get('fotos',[]))} foto(s)" if lc.get("fotos") else ""
                    st.caption(f"📅 {lc.get('data','')} | {lc.get('operador','–')} | pH:{lc.get('ph','–')} CRL:{lc.get('cloro_livre','–')}{ft}")
        st.markdown("</div>", unsafe_allow_html=True)

    # Para o restante da página não renderizar no modo operador
    st.stop()

if modo == "Modo Campo":
    st.info("Fluxo compacto habilitado para operação em campo no Windows / tablet Windows.")

if st.session_state.origem_dados_carregados:
    st.markdown(
        f"""
        <div class="quick-mode-box">
            <strong>Dados carregados:</strong> {st.session_state.origem_dados_carregados}<br>
            Agora você pode ajustar valor do aditivo, vigência ou renovar automaticamente.
        </div>
        """,
        unsafe_allow_html=True,
    )

if st.session_state.painel_acao_msg:
    st.success(st.session_state.painel_acao_msg)
    st.session_state.painel_acao_msg = ""

status_formulario = status_vencimento(
    st.session_state.get("data_fim", ""),
    st.session_state.alerta_vencimento_dias,
)

if st.session_state.get("data_fim", "").strip():
    if status_formulario["codigo"] == "vencido":
        st.markdown(
            f"<div class='alert-vencido'><strong>Alerta de vencimento:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )
    elif status_formulario["codigo"] == "vencendo":
        st.markdown(
            f"<div class='alert-vencendo'><strong>Atenção:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )
    elif status_formulario["codigo"] == "vigente":
        st.markdown(
            f"<div class='alert-vigente'><strong>Status:</strong> {status_formulario['mensagem']}</div>",
            unsafe_allow_html=True,
        )

# =========================================
# DADOS DE PAINEL
# =========================================

painel_vencimentos = listar_painel_vencimentos(st.session_state.alerta_vencimento_dias)
painel_filtrado = filtrar_itens_painel(
    painel_vencimentos,
    st.session_state.busca_rapida,
    st.session_state.filtro_status_central,
)

total_monitorado = len(painel_vencimentos)
total_vencidos = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vencido"])
total_vencendo = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vencendo"])
total_vigentes = len([i for i in painel_vencimentos if i["status"]["codigo"] == "vigente"])
total_indefinidos = len([i for i in painel_vencimentos if i["status"]["codigo"] == "indefinido"])
total_com_json = len([i for i in painel_vencimentos if i["tem_json"]])

itens_vencidos = [i for i in painel_filtrado if i["status"]["codigo"] == "vencido"]
itens_vencendo = [i for i in painel_filtrado if i["status"]["codigo"] == "vencendo"]
itens_indefinidos = [i for i in painel_filtrado if i["status"]["codigo"] == "indefinido"]

# =========================================
# DASHBOARD EXECUTIVO
# =========================================

taxa_estrutura = (total_com_json / total_monitorado * 100) if total_monitorado else 0
criticos = [i for i in painel_vencimentos if i["status"]["codigo"] in ("vencido", "vencendo")][:5]

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Dashboard executivo")

d1, d2, d3, d4, d5 = st.columns(5)
with d1:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Total monitorado</div><div class='dash-value'>{total_monitorado}</div><div class='dash-sub'>Pastas acompanhadas pelo sistema</div></div>",
        unsafe_allow_html=True,
    )
with d2:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vigentes</div><div class='dash-value'>{total_vigentes}</div><div class='dash-sub'>Dentro da vigência regular</div></div>",
        unsafe_allow_html=True,
    )
with d3:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vencem em breve</div><div class='dash-value'>{total_vencendo}</div><div class='dash-sub'>Dentro da faixa de alerta</div></div>",
        unsafe_allow_html=True,
    )
with d4:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Vencidos</div><div class='dash-value'>{total_vencidos}</div><div class='dash-sub'>Exigem ação imediata</div></div>",
        unsafe_allow_html=True,
    )
with d5:
    st.markdown(
        f"<div class='dash-mini'><div class='dash-title'>Cadastros estruturados</div><div class='dash-value'>{taxa_estrutura:.0f}%</div><div class='dash-sub'>{total_com_json} de {total_monitorado} com JSON salvo</div></div>",
        unsafe_allow_html=True,
    )

cx1, cx2 = st.columns([1.15, 1])
with cx1:
    st.markdown("**Resumo crítico**")
    if not criticos:
        st.success("Nenhum condomínio em faixa crítica no momento.")
    else:
        for item in criticos:
            st.markdown(
                f"- **{item['nome_exibicao']}** — {item['status']['rotulo']} — {item['status']['mensagem']}"
            )

with cx2:
    st.markdown("**Leitura rápida**")
    st.write(f"Sem vigência válida: **{total_indefinidos}**")
    st.write(f"Com ação prioritária agora: **{total_vencidos + total_vencendo}**")
    st.write(f"Faixa de lembrete atual: **{st.session_state.alerta_vencimento_dias} dias**")

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# SAÚDE DO SISTEMA
# =========================================

diag = diagnostico_sistema()

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Saúde do sistema")

s1, s2, s3, s4, s5, s6 = st.columns(6)
with s1:
    st.markdown(f"Template do contrato<br><span class='{'health-ok' if diag['template_contrato_ok'] else 'health-no'}'>{'OK' if diag['template_contrato_ok'] else 'Ausente'}</span>", unsafe_allow_html=True)
with s2:
    st.markdown(f"Template do aditivo<br><span class='{'health-ok' if diag['template_aditivo_ok'] else 'health-no'}'>{'OK' if diag['template_aditivo_ok'] else 'Ausente'}</span>", unsafe_allow_html=True)
with s3:
    st.markdown(f"Template do relatório<br><span class='{'health-ok' if diag['template_relatorio_ok'] else 'health-no'}'>{'OK' if diag['template_relatorio_ok'] else 'Ausente'}</span>", unsafe_allow_html=True)
with s4:
    st.markdown(f"Pasta de documentos<br><span class='{'health-ok' if diag['generated_ok'] else 'health-no'}'>{'OK' if diag['generated_ok'] else 'Ausente'}</span>", unsafe_allow_html=True)
with s5:
    st.markdown(f"Logo institucional<br><span class='{'health-ok' if diag['logo_ok'] else 'health-no'}'>{'OK' if diag['logo_ok'] else 'Não localizada'}</span>", unsafe_allow_html=True)
with s6:
    st.markdown(f"Ambiente Windows<br><span class='{'health-ok' if diag['windows_ok'] else 'health-no'}'>{'OK' if diag['windows_ok'] else 'Fora do padrão'}</span>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# BUSCA RÁPIDA
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Busca rápida profissional")

b1, b2, b3 = st.columns([2.2, 1.2, 1])
with b1:
    st.text_input(
        "Buscar por nome, CNPJ, síndico ou status",
        key="busca_rapida",
        placeholder="Ex.: terra nova, 12.345.678/0001-90, Marcelo, vencido...",
    )
with b2:
    st.selectbox(
        "Filtrar por status",
        options=["Todos", "Vigente", "Vence em breve", "Vencido", "Sem vigência válida"],
        key="filtro_status_central",
    )
with b3:
    st.metric("Resultado da busca", len(painel_filtrado))

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# GESTÃO DE OPERADORES
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("👷 Gestão de Operadores")
st.caption("Cadastre operadores com PIN exclusivo e defina quais condomínios cada um pode acessar.")

# Cria aba no Sheets se não existir
if st.button("🔧 Inicializar aba de Operadores no Sheets", key="btn_init_op_aba"):
    ok_aba = sheets_criar_aba_operadores()
    if ok_aba:
        st.success("✅ Aba '👷 Operadores' criada/confirmada no Google Sheets!")
    else:
        st.error("❌ Erro ao criar aba. Verifique a conexão com o Sheets.")

# Lista operadores cadastrados
@st.cache_data(ttl=30)
def _listar_ops():
    return sheets_listar_operadores()

ops_cadastrados = _listar_ops()

if ops_cadastrados:
    st.success(f"✅ {len(ops_cadastrados)} operador(es) cadastrado(s):")
    for op in ops_cadastrados:
        conds_txt = ", ".join(op.get("condomínios", [])) or "Todos"
        status = "🟢 Ativo" if op.get("ativo") else "🔴 Inativo"
        with st.expander(f"{status} — {op['nome']} | PIN: {op['pin'][:2]}{'*' * (len(op['pin'])-2)}"):
            st.write(f"**Condomínios:** {conds_txt}")
            if st.button(f"🗑 Remover {op['nome']}", key=f"del_op_{op['nome']}"):
                if sheets_deletar_operador(op["nome"]):
                    st.success(f"Operador '{op['nome']}' removido.")
                    st.cache_data.clear()
                    st.rerun()
else:
    st.info("Nenhum operador cadastrado. Use o formulário abaixo. O PIN 5010 continua funcionando como acesso geral.")

with st.expander("➕ Cadastrar / editar operador", expanded=not bool(ops_cadastrados)):
    # Carrega lista de clientes para selecionar condomínios
    @st.cache_data(ttl=60)
    def _clientes_para_op():
        return sheets_listar_clientes()
    _clientes_op = _clientes_para_op()

    op_col1, op_col2 = st.columns(2)
    with op_col1:
        op_nome_novo = st.text_input("Nome do operador *", key="op_novo_nome", placeholder="Ex.: João Silva")
        op_pin_novo  = st.text_input("PIN exclusivo *", key="op_novo_pin", placeholder="Ex.: 1234", max_chars=10,
            help="Mínimo 4 caracteres. Não use 5010 (reservado para acesso geral).")
    with op_col2:
        op_ativo_novo = st.checkbox("Operador ativo", value=True, key="op_novo_ativo")
        op_acesso_total_novo = st.checkbox("Acesso a todos os condomínios", value=False, key="op_novo_acesso_total")

    if not op_acesso_total_novo and _clientes_op:
        op_conds_novo = st.multiselect(
            "Condomínios que este operador pode acessar",
            options=_clientes_op,
            key="op_novo_conds",
            help="Deixe vazio + marque 'Acesso total' para liberar tudo."
        )
    else:
        op_conds_novo = _clientes_op if op_acesso_total_novo else []

    if st.button("💾 Salvar operador", type="primary", use_container_width=True, key="btn_salvar_op"):
        if not op_nome_novo.strip():
            st.error("Informe o nome do operador.")
        elif not op_pin_novo.strip() or len(op_pin_novo.strip()) < 4:
            st.error("PIN deve ter pelo menos 4 caracteres.")
        elif op_pin_novo.strip() == "5010":
            st.error("O PIN 5010 é reservado para acesso geral. Escolha outro.")
        else:
            with st.spinner("Salvando operador..."):
                ok_op = sheets_salvar_operador(
                    nome=op_nome_novo.strip(),
                    pin=op_pin_novo.strip(),
                    condomínios=op_conds_novo,
                    ativo=op_ativo_novo,
                )
            if ok_op:
                st.success(f"✅ Operador '{op_nome_novo}' salvo! PIN: {op_pin_novo}")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error("❌ Erro ao salvar operador. Verifique a conexão com o Sheets.")

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
    for k in ["cc_nome","cc_cnpj","cc_endereco","cc_contato","cc_telefone",
              "cc_vol_adulto","cc_vol_infantil","cc_vol_family"]:
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
            st.session_state["cc_endereco"]     = _cc_cliente_editar.get("endereco","")
            st.session_state["cc_contato"]      = _cc_cliente_editar.get("contato","")
            st.session_state["cc_telefone"]     = _cc_cliente_editar.get("telefone","")
            st.session_state["cc_vol_adulto"]   = str(_cc_cliente_editar.get("vol_adulto","") or "")
            st.session_state["cc_vol_infantil"] = str(_cc_cliente_editar.get("vol_infantil","") or "")
            st.session_state["cc_vol_family"]   = str(_cc_cliente_editar.get("vol_family","") or "")
            st.rerun()
    else:
        st.info("Nenhum cliente para editar.")

# ── Formulário ───────────────────────────────────────────────────────────────
def _mask_cc_cnpj():
    st.session_state["cc_cnpj"] = formatar_cnpj(st.session_state.get("cc_cnpj",""))

def _mask_cc_telefone():
    st.session_state["cc_telefone"] = formatar_telefone(st.session_state.get("cc_telefone",""))

cc1, cc2 = st.columns(2)
with cc1:
    cc_nome     = st.text_input("Nome do condomínio / local *", key="cc_nome", placeholder="Ex.: Residencial Bella Vista")
    cc_endereco = st.text_area("Endereço completo", key="cc_endereco", height=70, placeholder="Rua, número, bairro, cidade")
with cc2:
    cc_cnpj     = st.text_input("CNPJ (opcional)", key="cc_cnpj", placeholder="00.000.000/0000-00", on_change=_mask_cc_cnpj)
    cc_contato  = st.text_input("Síndico / responsável", key="cc_contato", placeholder="Nome do responsável")
    cc_telefone = st.text_input("Telefone (opcional)", key="cc_telefone", placeholder="(34) 99999-9999", on_change=_mask_cc_telefone)

# ── Volumes das piscinas ─────────────────────────────────────────────────────
st.markdown("**🏊 Volume das piscinas (m³)**")
cv1, cv2, cv3 = st.columns(3)
with cv1:
    cc_vol_adulto   = st.text_input("Piscina Adulto (m³)", key="cc_vol_adulto", placeholder="ex: 150")
with cv2:
    cc_vol_infantil = st.text_input("Piscina Infantil (m³)", key="cc_vol_infantil", placeholder="ex: 30")
with cv3:
    cc_vol_family   = st.text_input("Piscina Family (m³)", key="cc_vol_family", placeholder="ex: 50")

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
                ok = sheets_editar_cliente(
                    id_cliente=_cc_cliente_editar["id"],
                    nome=cc_nome.strip(), cnpj=cc_cnpj.strip(),
                    endereco=cc_endereco.strip(), contato=cc_contato.strip(),
                    telefone=cc_telefone.strip(),
                    vol_adulto=_vol_a, vol_infantil=_vol_i, vol_family=_vol_f,
                )
                msg_ok = f"✅ Cliente '{cc_nome}' atualizado!"
            else:
                ok = sheets_salvar_cliente(
                    nome=cc_nome.strip(), cnpj=cc_cnpj.strip(),
                    endereco=cc_endereco.strip(), contato=cc_contato.strip(),
                    telefone=cc_telefone.strip(),
                    vol_adulto=_vol_a, vol_infantil=_vol_i, vol_family=_vol_f,
                )
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
        csr_endereco = st.text_area("Endereço", key="csr_endereco", height=70, placeholder="Rua, número, bairro, cidade")
    with csr2:
        csr_cnpj = st.text_input("CNPJ (opcional)", key="csr_cnpj", placeholder="00.000.000/0000-00", on_change=_mask_csr_cnpj)
        csr_contato = st.text_input("Responsável / contato", key="csr_contato", placeholder="Nome do responsável")
        csr_telefone = st.text_input("Telefone (opcional)", key="csr_telefone", placeholder="(34) 99999-9999", on_change=_mask_csr_telefone)

    if st.button("➕ Salvar cliente sem RT", use_container_width=True):
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
st.markdown("**Gerar relatório técnico simples (sem RT)**")

clientes_sem_rt_reload = carregar_clientes_sem_rt() if CLIENTES_SEM_RT_JSON.exists() else []
opcoes_csr = [c["nome"] for c in clientes_sem_rt_reload]

if not opcoes_csr:
    st.info("Cadastre um cliente sem RT acima para gerar o relatório técnico.")
else:
    rts1, rts2, rts3 = st.columns([2, 1, 1])
    with rts1:
        csr_sel = st.selectbox("Selecione o cliente", opcoes_csr, key="csr_sel_relatorio")
    with rts2:
        csr_mes = st.text_input("Mês", key="csr_mes_rel", placeholder="04")
    with rts3:
        csr_ano = st.text_input("Ano", key="csr_ano_rel", placeholder="2026")

    csr_dados_sel = next((c for c in clientes_sem_rt_reload if c["nome"] == csr_sel), {})

    # Carrega lançamentos de campo do cliente selecionado
    pasta_csr_sel = GENERATED_DIR / slugify_nome(csr_sel) if csr_sel else None
    lancamentos_csr = []
    if pasta_csr_sel and pasta_csr_sel.exists():
        dados_csr_json = carregar_dados_condominio(pasta_csr_sel)
        lancamentos_csr = (dados_csr_json or {}).get("lancamentos_campo", [])

    if lancamentos_csr:
        st.markdown(f"<div style='border:1px solid rgba(20,120,60,0.3);border-radius:10px;padding:10px;background:rgba(20,120,60,0.07);'>"
            f"📱 <strong>{len(lancamentos_csr)} lançamento(s) de campo disponível(is)</strong> — "
            f"período: {lancamentos_csr[0].get('data','?')} a {lancamentos_csr[-1].get('data','?')}</div>",
            unsafe_allow_html=True)

    csr_operador_nome = st.text_input("Operador responsável", key="csr_operador_rel", placeholder="Nome do operador")
    csr_obs_geral = st.text_area("Observações gerais", key="csr_obs_rel", height=80,
        placeholder="Condições gerais da piscina, ocorrências, recomendações...")

    if st.button("📄 Gerar relatório técnico simples", type="primary", use_container_width=True):
        if not csr_sel or not csr_mes.strip() or not csr_ano.strip():
            st.error("Selecione o cliente, mês e ano.")
        else:
            try:
                from docx import Document as DocxDoc
                from docx.shared import Pt, Cm, Inches
                from docx.enum.text import WD_ALIGN_PARAGRAPH

                doc_rt = DocxDoc()

                for section in doc_rt.sections:
                    section.top_margin = Cm(2)
                    section.bottom_margin = Cm(2)
                    section.left_margin = Cm(2.5)
                    section.right_margin = Cm(2.5)

                def add_par(doc, texto, bold=False, size=11, align=None):
                    p = doc.add_paragraph()
                    if align:
                        p.alignment = align
                    run = p.add_run(texto)
                    run.bold = bold
                    run.font.size = Pt(size)
                    return p

                # ── Cabeçalho ──
                add_par(doc_rt, "AQUA GESTÃO – CONTROLE TÉCNICO DE PISCINAS", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
                add_par(doc_rt, "RELATÓRIO TÉCNICO DE QUALIDADE DA ÁGUA", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
                add_par(doc_rt, f"Mês de Referência: {csr_mes.strip()}/{csr_ano.strip()}  |  Emissão: {hoje_br()}", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
                doc_rt.add_paragraph()

                # ── 1. Identificação ──
                add_par(doc_rt, "1. IDENTIFICAÇÃO DO LOCAL", bold=True, size=11)
                tab_id = doc_rt.add_table(rows=5, cols=2)
                tab_id.autofit = True
                dados_id = [
                    ("Local / Condomínio", csr_dados_sel.get("nome", csr_sel)),
                    ("CNPJ", csr_dados_sel.get("cnpj", "Não informado") or "Não informado"),
                    ("Endereço", csr_dados_sel.get("endereco", "Não informado") or "Não informado"),
                    ("Responsável", csr_dados_sel.get("contato", "Não informado") or "Não informado"),
                    ("Operador de campo", csr_operador_nome or "Não informado"),
                ]
                for i, (k, v) in enumerate(dados_id):
                    tab_id.cell(i, 0).text = k
                    tab_id.cell(i, 1).text = v
                    tab_id.cell(i, 0).paragraphs[0].runs[0].bold = True
                doc_rt.add_paragraph()

                # ── 2. Base normativa ──
                add_par(doc_rt, "2. BASE NORMATIVA E FAIXAS DE REFERÊNCIA", bold=True, size=11)
                normas_texto = (
                    "As análises físico-químicas registradas neste relatório seguem as faixas operacionais "
                    "estabelecidas pela ABNT NBR 10339 (Piscinas — Projeto, execução e manutenção), "
                    "com referência complementar à Portaria GM/MS nº 888/2021 (padrão sanitário de água). "
                    "Os procedimentos de segurança química observam a NR-26 (Sinalização de Segurança) "
                    "e NR-6 (Equipamentos de Proteção Individual — EPI)."
                )
                p_normas = doc_rt.add_paragraph()
                p_normas.add_run(normas_texto).font.size = Pt(10)

                # Tabela de faixas de referência
                faixas_header = ["Parâmetro", "Unidade", "Faixa Mínima", "Faixa Máxima", "Referência"]
                faixas_data = [
                    ["pH", "—", "7,2", "7,8", "ABNT NBR 10339"],
                    ["Cloro Residual Livre (CRL)", "mg/L", "0,5", "3,0", "ABNT NBR 10339"],
                    ["Cloro Total (CT)", "mg/L", "—", "CRL + 0,5", "ABNT NBR 10339"],
                    ["Alcalinidade Total", "mg/L", "80", "120", "ABNT NBR 10339"],
                    ["Dureza Cálcica (DC)", "mg/L", "150", "300", "ABNT NBR 10339"],
                    ["Ácido Cianúrico (CYA)", "mg/L", "30", "50", "ABNT NBR 10339"],
                ]
                tab_faixas = doc_rt.add_table(rows=1 + len(faixas_data), cols=5)
                tab_faixas.style = "Table Grid"
                for j, h in enumerate(faixas_header):
                    c = tab_faixas.cell(0, j)
                    c.text = h
                    c.paragraphs[0].runs[0].bold = True
                    c.paragraphs[0].runs[0].font.size = Pt(9)
                for i, row in enumerate(faixas_data, 1):
                    for j, v in enumerate(row):
                        tab_faixas.cell(i, j).text = v
                        tab_faixas.cell(i, j).paragraphs[0].runs[0].font.size = Pt(9)
                doc_rt.add_paragraph()

                # ── 3. Análises ──
                add_par(doc_rt, "3. ANÁLISES FÍSICO-QUÍMICAS REGISTRADAS", bold=True, size=11)
                if lancamentos_csr:
                    headers_an = ["Data", "pH", "CRL", "CT", "Alc.", "DC", "CYA", "Operador"]
                    tab_an = doc_rt.add_table(rows=1 + len(lancamentos_csr), cols=len(headers_an))
                    tab_an.style = "Table Grid"
                    for j, h in enumerate(headers_an):
                        c = tab_an.cell(0, j)
                        c.text = h
                        c.paragraphs[0].runs[0].bold = True
                        c.paragraphs[0].runs[0].font.size = Pt(9)
                    for i, lc in enumerate(lancamentos_csr, 1):
                        vals = [lc.get("data",""), lc.get("ph",""), lc.get("cloro_livre",""),
                                lc.get("cloro_total",""), lc.get("alcalinidade",""),
                                lc.get("dureza",""), lc.get("cianurico",""), lc.get("operador","")]
                        for j, v in enumerate(vals):
                            tab_an.cell(i, j).text = str(v)
                            tab_an.cell(i, j).paragraphs[0].runs[0].font.size = Pt(9)
                else:
                    add_par(doc_rt, "Nenhuma análise registrada neste período.", size=10)
                doc_rt.add_paragraph()

                # ── 4. Dosagens ──
                add_par(doc_rt, "4. DOSAGENS DE PRODUTOS QUÍMICOS", bold=True, size=11)
                todas_dosagens = []
                for lc in lancamentos_csr:
                    for d in lc.get("dosagens", []):
                        if d.get("produto") and d["produto"] not in [x.get("produto") for x in todas_dosagens]:
                            todas_dosagens.append(d)
                if todas_dosagens:
                    tab_dos = doc_rt.add_table(rows=1 + len(todas_dosagens), cols=4)
                    tab_dos.style = "Table Grid"
                    for j, h in enumerate(["Produto", "Quantidade", "Unidade", "Finalidade"]):
                        c = tab_dos.cell(0, j)
                        c.text = h
                        c.paragraphs[0].runs[0].bold = True
                        c.paragraphs[0].runs[0].font.size = Pt(9)
                    for i, d in enumerate(todas_dosagens, 1):
                        for j, v in enumerate([d.get("produto",""), d.get("quantidade",""), d.get("unidade",""), d.get("finalidade","")]):
                            tab_dos.cell(i, j).text = v
                            tab_dos.cell(i, j).paragraphs[0].runs[0].font.size = Pt(9)
                else:
                    add_par(doc_rt, "Nenhuma dosagem registrada neste período.", size=10)
                doc_rt.add_paragraph()

                # ── 5. Observações ──
                add_par(doc_rt, "5. OBSERVAÇÕES TÉCNICAS", bold=True, size=11)
                obs_unidas = csr_obs_geral.strip()
                for lc in lancamentos_csr:
                    if lc.get("observacao"):
                        obs_unidas += f"\n• {lc.get('data','')}: {lc['observacao']}"
                p_obs = doc_rt.add_paragraph()
                p_obs.add_run(obs_unidas or "Sem observações registradas.").font.size = Pt(10)
                doc_rt.add_paragraph()

                # ── 6. Registro fotográfico ──
                pasta_csr_out_pre = GENERATED_DIR / slugify_nome(csr_sel)
                pasta_fotos_csr = pasta_csr_out_pre / "fotos_campo"
                fotos_encontradas = []
                if pasta_fotos_csr.exists():
                    for lc in lancamentos_csr:
                        for nome_foto in lc.get("fotos", []):
                            caminho_foto = pasta_fotos_csr / nome_foto
                            if caminho_foto.exists():
                                fotos_encontradas.append((lc.get("data",""), caminho_foto))

                if fotos_encontradas:
                    add_par(doc_rt, "6. REGISTRO FOTOGRÁFICO", bold=True, size=11)
                    for data_foto, caminho_foto in fotos_encontradas:
                        try:
                            p_foto_leg = doc_rt.add_paragraph()
                            p_foto_leg.add_run(f"Visita: {data_foto}").font.size = Pt(9)
                            p_foto = doc_rt.add_paragraph()
                            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_foto = p_foto.add_run()
                            run_foto.add_picture(str(caminho_foto), width=Inches(5.5))
                        except Exception:
                            add_par(doc_rt, f"[Foto não pôde ser inserida: {caminho_foto.name}]", size=9)
                    doc_rt.add_paragraph()
                    secao_aviso = 7
                else:
                    secao_aviso = 6

                # ── Aviso legal ──
                add_par(doc_rt, f"{secao_aviso}. AVISO LEGAL E GUARDA DOCUMENTAL", bold=True, size=10)
                p_aviso = doc_rt.add_paragraph()
                p_aviso.add_run(
                    "Recomenda-se a guarda e organização deste documento e de seus registros correlatos por prazo "
                    "mínimo de 5 (cinco) anos, para fins de rastreabilidade, auditoria, controle documental e "
                    "segurança jurídica. Este relatório é de caráter técnico-informativo e não substitui laudo "
                    "laboratorial microbiológico, cuja contratação é de responsabilidade do estabelecimento."
                ).font.size = Pt(9)
                doc_rt.add_paragraph()

                # ── Assinatura ──
                add_par(doc_rt, f"Uberlândia/MG, {hoje_br()}.", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                doc_rt.add_paragraph()
                tab_ass = doc_rt.add_table(rows=1, cols=2)
                tab_ass.autofit = True
                c_ass1 = tab_ass.cell(0, 0)
                c_ass2 = tab_ass.cell(0, 1)
                for cell_a, texto_a in [
                    (c_ass1, f"___________________________\n{csr_operador_nome or 'Operador'}\nAqua Gestão – Controle Técnico de Piscinas"),
                    (c_ass2, f"___________________________\n{csr_dados_sel.get('contato','Responsável')}\n{csr_sel}"),
                ]:
                    cell_a.paragraphs[0].clear()
                    cell_a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell_a.paragraphs[0].add_run(texto_a).font.size = Pt(9)

                # ── Salva ──
                pasta_csr_out = GENERATED_DIR / slugify_nome(csr_sel)
                pasta_csr_out.mkdir(parents=True, exist_ok=True)
                data_nome_csr = datetime.now().strftime("%Y%m%d_%H%M%S")
                docx_csr = pasta_csr_out / f"{data_nome_csr}_{slugify_nome(csr_sel)}_RELATORIO_TECNICO.docx"
                pdf_csr  = pasta_csr_out / f"{data_nome_csr}_{slugify_nome(csr_sel)}_RELATORIO_TECNICO.pdf"
                doc_rt.save(str(docx_csr))
                ok_pdf_csr, err_pdf_csr = converter_docx_para_pdf(docx_csr, pdf_csr)

                registrar_documento_manifest(pasta_csr_out, csr_sel, "Relatório", docx_csr, pdf_csr, ok_pdf_csr, err_pdf_csr)
                st.session_state["ultimos_docs_gerados"] = st.session_state.get("ultimos_docs_gerados") or {}
                st.session_state["ultimos_docs_gerados"].update({
                    "relatorio_docx": str(docx_csr),
                    "relatorio_pdf": str(pdf_csr) if ok_pdf_csr else None,
                })

                st.success(f"✅ Relatório técnico gerado para {csr_sel}! {len(fotos_encontradas)} foto(s) incluída(s).")
                dl1, dl2 = st.columns(2)
                with dl1:
                    with open(docx_csr, "rb") as f:
                        st.download_button("⬇️ Baixar DOCX", data=f, file_name=docx_csr.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True)
                with dl2:
                    if ok_pdf_csr and pdf_csr.exists():
                        with open(pdf_csr, "rb") as f:
                            st.download_button("⬇️ Baixar PDF", data=f, file_name=pdf_csr.name,
                                mime="application/pdf", use_container_width=True)
                    else:
                        st.warning(f"PDF não gerado: {err_pdf_csr}")

            except Exception as e:
                st.error(f"Erro ao gerar relatório: {e}")

st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# FORMULÁRIO
# =========================================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
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

st.markdown('<div class="section-card">', unsafe_allow_html=True)
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

st.markdown('<div class="section-card">', unsafe_allow_html=True)
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
    return filtrados if filtrados else lancamentos

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