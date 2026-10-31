"""Dados do relatório operacional Bem Star. Sem acesso implícito à rede ou ao RT."""
from copy import deepcopy
from datetime import datetime
from hashlib import sha256
from pathlib import Path
import io
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
    return str(pool.get('nome') or pool.get('piscina') or 'Piscina principal').strip()


def tipo_inferido(nome):
    n = nome.casefold()
    return 'Infantil' if 'infantil' in n else ('Adulto' if 'adult' in n else 'Não informado')


def expandir(lancamentos):
    """Uma linha por leitura/piscina. Nunca herda a química de outra piscina."""
    from utils.datas_visitas import normalizar_data_visita
    saida = []
    for lc in lancamentos or []:
        pools = [p for p in (lc.get('piscinas') or []) if isinstance(p, dict)]
        for p in pools or [lc]:
            linha = {k: deepcopy(v) for k, v in lc.items() if k not in {'piscinas', *[x[0] for x in PARAMETROS]}}
            # Com pools explícitas, campos ausentes ficam ausentes, não recebem valores globais.
            linha.update({k: p.get(k, '') for k, _, _ in PARAMETROS})
            linha.update(data=normalizar_data_visita(lc.get('data', '')), piscina=nome_piscina(p),
                         tipo_piscina=p.get('tipo_piscina') or p.get('tipo') or tipo_inferido(nome_piscina(p)))
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


def unir_agenda(originais, manuais):
    """Edita por data + piscina; preserva as outras piscinas e as referências das fotos."""
    resultado = {}
    for linha in expandir(originais):
        chave = (linha['data'], linha['piscina'])
        if chave not in resultado:
            resultado[chave] = linha
    for linha in expandir(manuais):
        chave = (linha['data'], linha['piscina'])
        anterior = resultado.get(chave, {})
        novo = {**anterior, **linha}
        for k in ['fotos', 'fotos_b64', 'fotos_limpeza', 'fotos_parametros', 'fotos_dosagens']:
            if anterior.get(k):
                novo[k] = deepcopy(anterior[k])
        resultado[chave] = novo
    return sorted(resultado.values(), key=lambda x: (data_ordem(x['data']), x['piscina']))


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


def salvar_complementos(estado, aba, enviar):
    """Commit append-only: JSON dividido e verificado, imagens no Drive por ID exato."""
    contexto = list(estado.get('_bs_rel_contexto') or [])
    ctx = contexto_id(contexto)
    fotos = estado.get('bs_mensal_fotos', [])
    for foto in fotos:
        if foto.get('contexto') != ctx:
            raise ValueError('Anexo de outro contexto.')
        if not foto.get('drive_id'):
            fid = enviar(foto['bytes'], f'BEMSTAR_MENSAL_{ctx.split(":")[1][:16]}_{foto["sha256"]}.jpg', f'BemStar-Mensal_{ctx.split(":")[1]}', f'{contexto[3]}-{contexto[2]}')
            if not fid:
                raise ValueError('Não foi possível salvar a foto no Drive. Os dados continuam nesta sessão; tente novamente.')
            foto['drive_id'] = fid
    campos = {k: v for k, v in dict(estado).items()
              if k.startswith('bs_agenda_') or k.startswith('bs_mensal_') or k in {
                  'csr_operador_rel', 'csr_obs_rel', 'bs_rel_parecer_tecnico', 'bs_rel_freq_semana',
                  'bs_rel_dias_semana', 'bs_rel_incluir_visitas_vazias'}}
    campos = deepcopy(campos)
    campos['bs_mensal_fotos'] = [{k: v for k, v in f.items() if k != 'bytes'} for f in fotos]
    payload = json.dumps({'contexto': contexto, 'campos': campos}, ensure_ascii=False, sort_keys=True)
    if len(payload) > 800_000:
        raise ValueError('Complementos excedem o limite do rascunho; reduza os textos.')
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
            if f.get('contexto') != ctx or not f.get('drive_id'):
                raise ValueError('Imagem sem referência válida.')
        return campos
    except (KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError('Rascunho inválido; os dados da tela foram preservados.') from exc
