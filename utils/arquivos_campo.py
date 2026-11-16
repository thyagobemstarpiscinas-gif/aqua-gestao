"""Snapshots de campo no Sheets: partes completas, referência e hash verificados.

Usa a mesma planilha privada do app, sem depender da cota de arquivos do Drive.
Uma referência só é entregue depois da releitura integral. Revisões anteriores
nunca são sobrescritas; uma gravação interrompida não substitui o rascunho válido.
"""
import base64
from copy import deepcopy
import hashlib
import json
import re
from pathlib import Path
import uuid

ABA = '_Arquivos_Campo'
TAMANHO_PARTE = 38000
LIMITE_BYTES = 32 * 1024 * 1024


def _json(obj):
    # ASCII também limita corretamente as unidades UTF-16 de cada célula.
    return json.dumps(obj, ensure_ascii=True, sort_keys=True, separators=(',', ':'))


def _chave(contexto):
    return hashlib.sha256(_json(contexto).encode()).hexdigest()


def obter_aba(sh, criar=False):
    from gspread import WorksheetNotFound
    try:
        return sh.worksheet(ABA)
    except WorksheetNotFound:
        if not criar:
            raise
        try:
            return sh.add_worksheet(title=ABA, rows=1000, cols=6)
        except Exception:
            # Outra sessão pode ter criado a aba durante a requisição.
            return sh.worksheet(ABA)


def carregar(ref, contexto, linhas):
    if not isinstance(ref, dict) or ref.get('chave') != _chave(contexto):
        raise ValueError('Arquivo de outro contexto; recuperação suspensa.')
    rows = [r for r in linhas if len(r) >= 6 and r[0] == ref['chave'] and r[1] == ref.get('revisao')]
    try:
        total = int(ref['partes'])
        if not 1 <= total <= 2000 or len(rows) != total:
            raise ValueError('Arquivo incompleto na nuvem.')
        if {int(r[2]) for r in rows} != set(range(total)):
            raise ValueError('Partes ausentes ou duplicadas na nuvem.')
        payload = ''.join(r[5] for r in sorted(rows, key=lambda r: int(r[2])))
        digest = hashlib.sha256(payload.encode()).hexdigest()
        if digest != ref['sha256'] or any(r[3] != str(total) or r[4] != digest for r in rows):
            raise ValueError('Integridade do arquivo não confirmada.')
        obj = json.loads(payload)
        if obj['contexto'] != contexto or not isinstance(obj['dados'], dict):
            raise ValueError('Conteúdo de outro contexto.')
        return obj['dados']
    except (KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError('Referência de arquivo inválida.') from exc


def salvar(sh, contexto, dados):
    payload = _json({'contexto': contexto, 'dados': dados})
    if len(payload.encode()) > LIMITE_BYTES:
        raise ValueError('Fotos e dados excedem 32 MB. Reduza os anexos antes de salvar.')
    partes = [payload[i:i + TAMANHO_PARTE] for i in range(0, len(payload), TAMANHO_PARTE)]
    ref = {'chave': _chave(contexto), 'revisao': uuid.uuid4().hex,
           'partes': len(partes), 'sha256': hashlib.sha256(payload.encode()).hexdigest()}
    aba = obter_aba(sh, criar=True)
    linhas = [[ref['chave'], ref['revisao'], str(i), str(len(partes)), ref['sha256'], parte]
              for i, parte in enumerate(partes)]
    # Lotes pequenos evitam requisições excessivas; só depois de todos os lotes
    # completos a referência poderá ser gravada na visita/rascunho principal.
    ref['faixas'] = []
    for i in range(0, len(linhas), 20):
        resposta = aba.append_rows(linhas[i:i + 20], value_input_option='RAW')
        faixa = (resposta.get('updates') or {}).get('updatedRange', '').split('!')[-1]
        if not re.fullmatch(r'A[1-9][0-9]*:F[1-9][0-9]*', faixa):
            raise ValueError('Não foi possível localizar a cópia gravada na nuvem.')
        ref['faixas'].append(faixa)
    if carregar_varios(sh, [(ref, contexto)])[0] != dados:
        raise ValueError('A cópia na nuvem difere dos dados enviados.')
    return ref


def carregar_varios(sh, pedidos):
    """Uma leitura em lote apenas dos intervalos referenciados, não do arquivo inteiro."""
    faixas = []
    tamanhos = []
    for ref, contexto in pedidos:
        if not isinstance(ref, dict) or ref.get('chave') != _chave(contexto):
            raise ValueError('Arquivo de outro contexto.')
        intervalos = ref.get('faixas')
        if not isinstance(intervalos, list) or not 1 <= len(intervalos) <= 100:
            raise ValueError('Intervalos do arquivo indisponíveis.')
        total = 0
        for faixa in intervalos:
            m = re.fullmatch(r'A([1-9][0-9]*):F([1-9][0-9]*)', str(faixa))
            if not m or not 1 <= int(m[2])-int(m[1])+1 <= 20:
                raise ValueError('Intervalo do arquivo inválido.')
            total += int(m[2])-int(m[1])+1
        if total != ref.get('partes'):
            raise ValueError('Intervalos incompletos.')
        faixas.extend(intervalos); tamanhos.append(len(intervalos))
    if not pedidos:
        return []
    grupos = obter_aba(sh).batch_get(faixas)
    if len(grupos) != len(faixas):
        raise ValueError('Leitura incompleta dos arquivos.')
    saida, inicio = [], 0
    for (ref, contexto), quantidade in zip(pedidos, tamanhos):
        linhas = [r for grupo in grupos[inicio:inicio+quantidade] for r in grupo]
        saida.append(carregar(ref, contexto, linhas))
        inicio += quantidade
    return saida


def contexto_visita(lancamento):
    return ['visita', str(lancamento.get('condominio') or '').strip(),
            str(lancamento.get('id_visita') or '')]


def validar_fotos_visita(lancamento):
    for categoria in ('antes', 'depois', 'cmaq', 'extras'):
        nomes = lancamento.get(f'fotos_{categoria}') or []
        conteudos = lancamento.get(f'fotos_{categoria}_b64') or []
        if nomes and len(conteudos) != len(nomes):
            raise ValueError(f'As fotos de {categoria} ainda não têm cópia completa. Recupere ou anexe novamente antes de salvar.')
        for conteudo in conteudos:
            if not isinstance(conteudo, str) or not base64.b64decode(conteudo, validate=True):
                raise ValueError(f'Conteúdo de foto inválido: {categoria}.')


def contexto_rascunho(nome):
    return ['rascunho', str(nome).strip().casefold()]


def _nome_seguro(nome):
    if not isinstance(nome, str) or not nome.startswith('rasc_') or Path(nome).name != nome or '\\' in nome:
        raise ValueError('Nome de foto inválido no rascunho.')
    return nome


def anexar_fotos_rascunho(dados, pasta):
    obj = deepcopy(dados)
    fotos = {}
    for nomes in (dados.get('fotos_rascunho') or {}).values():
        for nome in nomes:
            nome = _nome_seguro(nome)
            path = Path(pasta) / nome
            if not path.is_file():
                raise ValueError(f'Foto do rascunho indisponível: {nome}.')
            fotos[nome] = base64.b64encode(path.read_bytes()).decode('ascii')
    obj['_fotos_campo'] = fotos
    return obj


def restaurar_fotos_rascunho(dados, pasta):
    obj = deepcopy(dados)
    anexos = obj.pop('_fotos_campo', {})
    conteudos = { _nome_seguro(nome): base64.b64decode(b64, validate=True)
                 for nome, b64 in anexos.items() }
    esperadas = {n for nomes in (obj.get('fotos_rascunho') or {}).values() for n in nomes}
    if set(conteudos) != esperadas:
        raise ValueError('Fotos do rascunho incompletas; recuperação suspensa.')
    if conteudos:
        Path(pasta).mkdir(parents=True, exist_ok=True)
        for nome, conteudo in conteudos.items():
            destino = Path(pasta) / nome
            if destino.exists() and destino.read_bytes() != conteudo:
                raise ValueError('Uma foto local mais nova foi preservada; confira o rascunho.')
        for nome, conteudo in conteudos.items():
            (Path(pasta) / nome).write_bytes(conteudo)
    return obj
