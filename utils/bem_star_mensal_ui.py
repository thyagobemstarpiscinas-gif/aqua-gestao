"""Campos mensais Bem Star, instanciados depois da seleção e troca do contexto."""
from utils.bem_star_mensal import (PARAMETROS, TIPOS, AVISO, contexto_id, expandir, tipo_inferido,
                                  token_pool, chave_campo, anexar_foto, data_ordem)


def renderizar_piscinas(st, lancamentos):
    ss = st.session_state
    # Protege os campos das piscinas recolhidas da limpeza automática de widgets.
    for k in list(ss):
        if k.startswith('bs_agenda_pool_'):
            ss[k] = ss[k]
    existentes = list(dict.fromkeys(l['piscina'] for l in expandir(lancamentos)))
    st.markdown('### Piscinas do relatório')
    st.caption('Cada piscina tem leituras, dosagens e gráficos próprios. Classifique a piscina principal como Adulto ou Infantil quando aplicável.')
    extras = st.text_input('Outras piscinas (separe os nomes por ponto e vírgula)', key='bs_mensal_outras',
                           placeholder='Ex.: Piscina aquecida; Spa')
    opcoes = list(dict.fromkeys(existentes + ['Piscina adulta', 'Piscina infantil', 'Piscina principal'] +
                               [n.strip() for n in extras.split(';') if n.strip()] + list(ss.get('bs_mensal_piscinas') or [])))
    if not ss.get('bs_mensal_piscinas'):
        ss['bs_mensal_piscinas'] = existentes or ['Piscina adulta']
    st.multiselect('Piscinas para preencher na agenda', opcoes, key='bs_mensal_piscinas')
    nomes = list(dict.fromkeys(existentes + list(ss.get('bs_mensal_piscinas') or [])))
    st.caption('As piscinas já presentes nas visitas salvas permanecem no relatório. Esta seleção define os campos manuais da agenda.')
    for nome in nomes:
        chave = f'bs_mensal_tipo_{token_pool(nome)}'
        if chave not in ss or ss[chave] not in TIPOS:
            ss[chave] = next((l['tipo_piscina'] for l in expandir(lancamentos)
                              if l['piscina'] == nome and l['tipo_piscina'] in TIPOS), tipo_inferido(nome))
        st.selectbox(f'Tipo — {nome}', TIPOS, key=chave)
    st.info(AVISO)


def renderizar_agenda(st, lancamentos, datas, preencher):
    ss = st.session_state
    originais = {(l['data'], l['piscina']): l for l in expandir(lancamentos)}
    for data in datas:
        with st.expander(f'Visita — {data}', expanded=False):
            for nome in ss.get('bs_mensal_piscinas', []) or []:
                st.markdown(f'**{nome} · {ss.get("bs_mensal_tipo_" + token_pool(nome), "Não informado")}**')
                base = originais.get((data, nome), {})
                # Somente a piscina de mesmo nome preenche seus campos.
                for k, _, _ in PARAMETROS:
                    preencher(chave_campo(nome, data, k), base.get(k, ''))
                preencher(chave_campo(nome, data, 'observacao'), base.get('observacao', ''))
                if data_ordem(data).weekday() == 0:
                    st.caption('Análise semanal completa — preencha somente os parâmetros efetivamente medidos.')
                else:
                    st.caption('Acompanhamento operacional. Leituras adicionais, quando realizadas, também entram nos gráficos.')
                colunas = st.columns(3)
                for i, (k, label, unidade) in enumerate(PARAMETROS):
                    with colunas[i % 3]:
                        st.text_input(label + (' (' + unidade + ')' if unidade else ''), key=chave_campo(nome, data, k))
                st.text_area('Observação — ' + nome, key=chave_campo(nome, data, 'observacao'), height=70)
                st.caption('Dosagens aplicadas — ' + nome)
                dosagens = base.get('dosagens') or []
                for i in range(3):
                    cols = st.columns([2, 1, 1, 2])
                    for j, (campo, label) in enumerate([('produto', 'Produto'), ('quantidade', 'Qtd.'), ('unidade', 'Unid.'), ('finalidade', 'Finalidade')]):
                        chave = chave_campo(nome, data, f'dos_{i}_{campo}')
                        preencher(chave, dosagens[i].get(campo, '') if i < len(dosagens) else '')
                        with cols[j]:
                            st.text_input(label, key=chave, label_visibility='visible' if i == 0 else 'collapsed')


def renderizar_fotos(st):
    ss = st.session_state
    ctx = ss.get('_bs_rel_contexto')
    cid = contexto_id(ctx)
    st.markdown('### Imagens do relatório mensal')
    st.caption('Até 12 imagens adicionais, de até 10 MB cada. JPG, PNG ou WebP. As fotos das visitas deste mês também são aproveitadas.')
    fotos = ss.setdefault('bs_mensal_fotos', [])
    # Uploader tem chave exclusiva do contexto e não participa do rascunho serializado.
    epoch = ss.get('_bs_mensal_upload_epoch', 0)
    uploads = st.file_uploader('Adicionar imagens', type=['jpg', 'jpeg', 'png', 'webp'],
                               accept_multiple_files=True, key=f'_bs_upload_{cid}_{epoch}')
    vistos = ss.setdefault('_bs_mensal_upload_vistos', {})
    import hashlib
    processados = vistos.setdefault(cid, [])
    for upload in uploads or []:
        bruto = upload.getvalue()
        digest = hashlib.sha256(bruto).hexdigest()
        if digest in processados:
            continue
        try:
            anexar_foto(fotos, bruto, upload.name, ctx)
            processados.append(digest)
        except ValueError as exc:
            st.error(f'{upload.name}: {exc}')
    nomes = ['Geral'] + list(dict.fromkeys(ss.get('bs_mensal_piscinas') or []))
    for foto in list(fotos):
        digest = foto['sha256']
        with st.expander('Imagem — ' + foto['nome']):
            if foto.get('bytes'):
                st.image(foto['bytes'], width=300)
            else:
                st.caption('Imagem salva no Drive. Será recuperada ao gerar o PDF.')
            # Chaves transitórias por contexto, valores canônicos no item do anexo.
            prefixo = f'_bs_foto_{cid}_{digest}_'
            opcoes = list(dict.fromkeys(nomes + [foto.get('piscina', 'Geral')]))
            foto['piscina'] = st.selectbox('Piscina da imagem', opcoes,
                                           index=opcoes.index(foto.get('piscina', 'Geral')), key=prefixo+'pool')
            foto['legenda'] = st.text_input('Legenda', value=foto.get('legenda', ''), max_chars=220, key=prefixo+'legenda')
            foto['data'] = st.text_input('Data da imagem (opcional, DD/MM/AAAA)', value=foto.get('data', ''), key=prefixo+'data')
            if foto['data']:
                d = data_ordem(foto['data'])
                if d.year != int(ctx[3]) or d.month != int(ctx[2]):
                    st.warning('A data da imagem precisa pertencer ao mês deste relatório; corrija antes de gerar.')
            if st.button('Remover imagem', key=prefixo+'remover'):
                fotos.remove(foto)
                vistos[cid] = []
                ss['_bs_mensal_upload_epoch'] = epoch + 1
                st.rerun()
    if fotos:
        st.caption(f'{len(fotos)} imagem(ns) adicional(is) anexada(s) a este cliente e mês.')
