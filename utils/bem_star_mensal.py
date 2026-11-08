"""Dados do relatório operacional Bem Star. Sem acesso implícito à rede ou ao RT."""
from copy import deepcopy
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import io
import base64
import json
import math
import re
import uuid

PARAMETROS = [('ph', 'pH', ''), ('cloro_livre', 'Cloro livre', 'mg/L'),
              ('cloro_total', 'Cloro total', 'mg/L'), ('alcalinidade', 'Alcalinidade', 'mg/L'),
              ('dureza', 'Dureza cálcica', 'mg/L'), ('cianurico', 'Ácido cianúrico', 'mg/L')]
TIPOS = ['Não informado', 'Adulto', 'Infantil', 'Outros']
AVISO = ('Registro operacional de manutenção. Não constitui relatório de Responsabilidade Técnica (RT), '
         'laudo ou ART e não substitui o acompanhamento por profissional habilitado quando aplicável.')
REFERENCIA = 'Resolução CFQ nº 332/2025, alterada pela Resolução CFQ nº 345/2026.'
ABA = '_BemStar_Mensal'


def contexto_id(contexto):
    if not contexto or len(contexto) != 4 or not contexto[1] or not contexto[2].isdigit() or not contexto[3].isdigit():
        raise ValueError('Selecione um cliente e período válido.')
    if not 1 <= int(contexto[2]) <= 12 or not 2000 <= int(contexto[3]) <= 2100:
        raise ValueError('Mês ou ano inválido.')
    return 'bem_star:' + sha256(json.dumps(list(contexto), ensure_ascii=False).encode()).hexdigest()


def numero(valor):
    try:
        n = float(str(valor).strip().replace(',', '.'))
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None


def data_ordem(valor):
    from utils.datas_visitas import normalizar_data_visita
    try:
        return datetime.strptime(normalizar_data_visita(valor), '%d/%m/%Y')
    except (ValueError, TypeError):
        return datetime.max


def nome_piscina(pool):
    return re.sub(r'\s+', ' ', str(pool.get('nome') or pool.get('piscina') or 'Piscina principal')).strip()


# Variações de grafia que designam a mesma piscina ("Piscina adulta" e
# "Piscina Adulto"). Números, outras palavras e tipos diferentes continuam
# distinguindo piscinas: "Piscina Adulto 2" e "Piscina Infantil" não se unem.
_SINONIMOS_POOL = {'adulta': 'adulto', 'adultos': 'adulto', 'adultas': 'adulto',
                   'infantis': 'infantil', 'piscinas': 'piscina'}


def chave_piscina(nome):
    """Chave de comparação do nome da piscina (caixa, acento, espaços, gênero)."""
    import unicodedata
    texto = unicodedata.normalize('NFKD', str(nome or '')).casefold()
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    tokens = re.findall(r'[a-z0-9]+', texto)
    return ' '.join(_SINONIMOS_POOL.get(t, t) for t in tokens)


def canonizar_nomes_piscinas(nomes, preferidos=()):
    """Remove duplicatas equivalentes mantendo o nome preferido (cadastro) ou o primeiro."""
    escolhido = {}
    for nome in list(preferidos or []):
        nome = re.sub(r'\s+', ' ', str(nome or '')).strip()
        if nome:
            escolhido.setdefault(chave_piscina(nome), nome)
    saida = []
    for nome in nomes or []:
        nome = re.sub(r'\s+', ' ', str(nome or '')).strip()
        if not nome:
            continue
        final = escolhido.setdefault(chave_piscina(nome), nome)
        if final not in saida:
            saida.append(final)
    return saida


def _id_visita(lc):
    return re.sub(r'\s+', '', str(lc.get('id_visita') or lc.get('id') or '')).upper()


def _vazio(valor):
    return valor is None or valor == '' or valor == [] or valor == {}


def _assinatura_gravacao(lc):
    """Data + momento da gravação + operador: identifica a mesma visita sem ID.

    O modo campo grava a cópia local antes de o Sheets gerar o ID; o payload do
    Sheets guarda o mesmo "salvo_em". Sem "salvo_em" não há assinatura segura.
    """
    from utils.datas_visitas import normalizar_data_visita
    salvo = str(lc.get('salvo_em') or '').strip()
    if not salvo:
        return None
    return (normalizar_data_visita(lc.get('data', '')), salvo,
            re.sub(r'\s+', ' ', str(lc.get('operador') or '')).strip().casefold())


def _combinar(base, novo):
    for k, v in novo.items():
        if not _vazio(v):
            base[k] = deepcopy(v)


def deduplicar_lancamentos(lancamentos):
    """Uma entrada por visita: a mesma visita (local + Sheets) vira um único lançamento.

    Identidade: ID da visita; para a cópia local gravada antes do ID existir,
    data + "salvo_em" + operador. Os campos são combinados: valores não vazios
    das ocorrências posteriores (Sheets, fonte oficial) prevalecem; campos que
    só a cópia local tem (fotos em base64, nomes de arquivos) são mantidos.
    Visitas com IDs diferentes nunca são unidas; sem ID nem assinatura, só
    cópias idênticas são removidas.
    """
    saida, por_id, por_assinatura, vistos = [], {}, {}, set()
    for lc in lancamentos or []:
        if not isinstance(lc, dict):
            continue
        vid = _id_visita(lc)
        assinatura = _assinatura_gravacao(lc)
        existente = por_id.get(vid) if vid else None
        if existente is None and assinatura in por_assinatura:
            candidato = por_assinatura[assinatura]
            if not vid or not _id_visita(candidato):
                existente = candidato
        if existente is not None:
            _combinar(existente, lc)
            if vid:
                por_id.setdefault(vid, existente)
            continue
        if not vid and not assinatura:
            chave = json.dumps(lc, ensure_ascii=False, sort_keys=True, default=str)
            if chave in vistos:
                continue
            vistos.add(chave)
        item = deepcopy(lc)
        saida.append(item)
        if vid:
            por_id[vid] = item
        if assinatura:
            por_assinatura.setdefault(assinatura, item)
    return saida


def tipo_inferido(nome):
    n = nome.casefold()
    return 'Infantil' if 'infantil' in n else ('Adulto' if 'adult' in n else 'Não informado')


def expandir(lancamentos, preferidos=()):
    """Uma linha por leitura/piscina. Nunca herda a química de outra piscina.

    Visitas repetidas (mesmo ID) aparecem uma vez; grafias equivalentes da
    mesma piscina recebem um único nome (o do cadastro, quando informado).
    """
    from utils.datas_visitas import normalizar_data_visita
    saida = []
    nomes_canon = {}
    for nome in preferidos or []:
        nome = re.sub(r'\s+', ' ', str(nome or '')).strip()
        if nome:
            nomes_canon.setdefault(chave_piscina(nome), nome)
    for lc in deduplicar_lancamentos(lancamentos):
        pools = [p for p in (lc.get('piscinas') or []) if isinstance(p, dict)]
        for p in pools or [lc]:
            linha = {k: deepcopy(v) for k, v in lc.items() if k not in {'piscinas', *[x[0] for x in PARAMETROS]}}
            # Com pools explícitas, campos ausentes ficam ausentes, não recebem valores globais.
            linha.update({k: p.get(k, '') for k, _, _ in PARAMETROS})
            _nome = nomes_canon.setdefault(chave_piscina(nome_piscina(p)), nome_piscina(p))
            linha.update(data=normalizar_data_visita(lc.get('data', '')), piscina=_nome,
                         tipo_piscina=p.get('tipo_piscina') or p.get('tipo') or tipo_inferido(_nome))
            if len(pools) > 1 and 'dosagens' not in p:
                linha['dosagens'] = []
                if lc.get('dosagens'):
                    linha['dosagens_da_visita'] = deepcopy(lc['dosagens'])
            if 'dosagens' in p:
                linha['dosagens'] = deepcopy(p['dosagens'])
            if 'observacao' in p:
                linha['observacao'] = p['observacao']
            saida.append(linha)
    return sorted(saida, key=lambda x: (data_ordem(x['data']), x['piscina']))


def unir_agenda(originais, manuais, preferidos=()):
    """Edita por data + piscina; preserva as outras piscinas e as referências das fotos."""
    resultado = expandir(originais, preferidos)
    nomes = list(preferidos or []) + [l['piscina'] for l in resultado]
    for linha in expandir(manuais, nomes):
        chave = (linha['data'], chave_piscina(linha['piscina']))
        indice = next((i for i, item in enumerate(resultado)
                       if (item['data'], chave_piscina(item['piscina'])) == chave), None)
        anterior = resultado[indice] if indice is not None else {}
        novo = {**anterior, **linha}
        if anterior:
            novo['piscina'] = anterior['piscina']
        for k in ['fotos', 'fotos_b64', 'fotos_limpeza', 'fotos_parametros', 'fotos_dosagens']:
            if anterior.get(k):
                novo[k] = deepcopy(anterior[k])
        if indice is None:
            resultado.append(novo)
        else:
            # A agenda edita a primeira visita exibida, sem apagar outras coletas no mesmo dia.
            resultado[indice] = novo
    return sorted(resultado, key=lambda x: (data_ordem(x['data']), x['piscina']))


def token_pool(nome):
    return sha256(str(nome).encode()).hexdigest()[:16]


def chave_campo(nome, data, campo):
    return f'bs_agenda_pool_{token_pool(nome)}_{re.sub(r"[^0-9]", "", data)}_{campo}'


def coletar_agenda(estado, datas):
    saida = []
    for data in datas:
        for nome in estado.get('bs_mensal_piscinas', []) or []:
            valores = {k: str(estado.get(chave_campo(nome, data, k), '') or '').strip() for k, _, _ in PARAMETROS}
            obs = str(estado.get(chave_campo(nome, data, 'observacao'), '') or '').strip()
            dosagens = []
            for i in range(3):
                d = {k: str(estado.get(chave_campo(nome, data, f'dos_{i}_{k}'), '') or '').strip()
                     for k in ['produto', 'quantidade', 'unidade', 'finalidade']}
                if any(d.values()):
                    dosagens.append(d)
            if not any(valores.values()) and not obs and not dosagens and not estado.get('bs_rel_incluir_visitas_vazias', False):
                continue
            saida.append(dict(data=data, piscina=nome,
                              tipo_piscina=estado.get(f'bs_mensal_tipo_{token_pool(nome)}', tipo_inferido(nome)),
                              operador=estado.get('csr_operador_rel', ''), **valores,
                              observacao=obs, dosagens=dosagens, fonte='Agenda mensal Bem Star'))
    return saida


def normalizar_foto(conteudo):
    from PIL import Image, ImageOps
    if len(conteudo) > 10 * 1024 * 1024:
        raise ValueError('Cada imagem deve ter até 10 MB.')
    try:
        with Image.open(io.BytesIO(conteudo)) as im:
            if im.width * im.height > 40_000_000:
                raise ValueError('Imagem muito grande: limite de 40 megapixels.')
            im = ImageOps.exif_transpose(im)
            im.thumbnail((1800, 1800))
            if im.mode != 'RGB':
                fundo = Image.new('RGB', im.size, 'white')
                if 'A' in im.getbands():
                    fundo.paste(im, mask=im.getchannel('A'))
                else:
                    fundo.paste(im.convert('RGB'))
                im = fundo
            b = io.BytesIO()
            im.save(b, format='JPEG', quality=88, optimize=True)
            return b.getvalue()
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('Arquivo não reconhecido como imagem válida.') from exc


def anexar_foto(lista, dados, nome, contexto):
    ctx = contexto_id(contexto)
    dados = normalizar_foto(dados)
    digest = sha256(dados).hexdigest()
    if any(f['sha256'] == digest for f in lista):
        return False
    if len(lista) >= 12:
        raise ValueError('Limite de 12 imagens adicionais por relatório.')
    lista.append(dict(contexto=ctx, sha256=digest, bytes=dados, nome=Path(nome).name,
                      legenda='', piscina='Geral', data='', drive_id=''))
    return True


def materializar_fotos(lista, contexto, raiz, baixar=None):
    ctx = contexto_id(contexto)
    pasta = Path(raiz) / '_bem_star_mensal' / ctx.split(':')[1]
    pasta.mkdir(parents=True, exist_ok=True)
    paths, metadados = [], {}
    for foto in lista:
        if foto.get('contexto') != ctx:
            raise ValueError('Foto de outro cliente ou período: geração suspensa.')
        if foto.get('data'):
            d = data_ordem(foto['data'])
            if d.year != int(contexto[3]) or d.month != int(contexto[2]):
                raise ValueError('A data da imagem deve pertencer ao mês do relatório.')
        dados = foto.get('bytes')
        if not dados and baixar and foto.get('drive_id'):
            dados = baixar(foto['drive_id'])
        if not dados or sha256(dados).hexdigest() != foto['sha256']:
            raise ValueError(f'Não foi possível recuperar a imagem {foto.get("nome", "")}. Tente novamente.')
        path = pasta / (foto['sha256'] + '.jpg')
        path.write_bytes(dados)
        paths.append(path)
        metadados[str(path)] = {k: foto.get(k, '') for k in ['nome', 'legenda', 'piscina', 'data']}
    return paths, metadados


def salvar_complementos(estado, aba, enviar=None):
    """Commit append-only: fotos e campos divididos em células e verificados por hash.

    Não depende de permissões de gravação no Drive. O argumento enviar é mantido
    por compatibilidade com o chamador, mas não é utilizado para novos anexos.
    """
    contexto = list(estado.get('_bs_rel_contexto') or [])
    ctx = contexto_id(contexto)
    fotos = estado.get('bs_mensal_fotos', [])
    anexos = []
    for foto in fotos:
        if foto.get('contexto') != ctx:
            raise ValueError('Anexo de outro contexto.')
        item = {k: v for k, v in foto.items() if k != 'bytes'}
        dados = foto.get('bytes')
        if dados:
            if sha256(dados).hexdigest() != foto['sha256']:
                raise ValueError('Integridade da imagem não confirmada; anexe novamente.')
            item['conteudo_b64'] = base64.b64encode(dados).decode('ascii')
        elif not foto.get('drive_id'):
            raise ValueError('Conteúdo da imagem indisponível; anexe novamente.')
        anexos.append(item)
    campos = {k: v for k, v in dict(estado).items()
              if k.startswith('bs_agenda_') or k.startswith('bs_mensal_') or k in {
                  'csr_operador_rel', 'csr_obs_rel', 'bs_rel_parecer_tecnico', 'bs_rel_freq_semana',
                  'bs_rel_dias_semana', 'bs_rel_incluir_visitas_vazias'}}
    campos = deepcopy(campos)
    campos['bs_mensal_fotos'] = anexos
    payload = json.dumps({'contexto': contexto, 'campos': campos}, ensure_ascii=False, sort_keys=True)
    if len(payload) > 12_000_000:
        raise ValueError('Complementos excedem 12 MB; reduza o número ou a resolução das imagens.')
    partes = [payload[i:i+38000] for i in range(0, len(payload), 38000)]
    revisao, digest = uuid.uuid4().hex, sha256(payload.encode()).hexdigest()
    aba.append_rows([[ctx, revisao, str(i), str(len(partes)), digest, parte]
                     for i, parte in enumerate(partes)], value_input_option='RAW')
    return revisao


def carregar_complementos(contexto, linhas):
    ctx = contexto_id(contexto)
    grupos = {}
    for row in linhas:
        if len(row) >= 6 and row[0] == ctx:
            grupos.setdefault(row[1], []).append(row)
    if not grupos:
        return None
    # Uma revisão incompleta não deve fazer o app restaurar silenciosamente uma antiga.
    rows = list(grupos.values())[-1]
    try:
        total = int(rows[0][3])
        if len(rows) != total or {int(r[2]) for r in rows} != set(range(total)):
            raise ValueError('Rascunho incompleto na nuvem.')
        payload = ''.join(r[5] for r in sorted(rows, key=lambda r: int(r[2])))
        if any(r[3] != str(total) or r[4] != sha256(payload.encode()).hexdigest() for r in rows):
            raise ValueError('Integridade do rascunho não confirmada.')
        obj = json.loads(payload)
        if obj['contexto'] != list(contexto):
            raise ValueError('Identidade do rascunho divergente.')
        campos = obj['campos']
        permitidas = {'csr_operador_rel', 'csr_obs_rel', 'bs_rel_parecer_tecnico', 'bs_rel_freq_semana',
                      'bs_rel_dias_semana', 'bs_rel_incluir_visitas_vazias'}
        if any(not (k.startswith('bs_agenda_') or k.startswith('bs_mensal_') or k in permitidas) for k in campos):
            raise ValueError('Campos inválidos no rascunho.')
        for f in campos.get('bs_mensal_fotos', []):
            if f.get('contexto') != ctx:
                raise ValueError('Imagem de outro contexto.')
            if f.get('conteudo_b64'):
                try:
                    dados = base64.b64decode(f.pop('conteudo_b64'), validate=True)
                except Exception as exc:
                    raise ValueError('Imagem salva com conteúdo inválido.') from exc
                if sha256(dados).hexdigest() != f.get('sha256'):
                    raise ValueError('Integridade da imagem salva não confirmada.')
                f['bytes'] = dados
            elif not f.get('drive_id'):
                raise ValueError('Imagem sem conteúdo ou referência válida.')
        return campos
    except (KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError('Rascunho inválido; os dados da tela foram preservados.') from exc
