"""Campos mensais Bem Star, instanciados depois da seleção e troca do contexto."""
from utils.bem_star_mensal import (PARAMETROS, TIPOS, AVISO, contexto_id, expandir, tipo_inferido,
                                  token_pool, chave_campo, anexar_foto, data_ordem, chave_piscina,
                                  canonizar_nomes_piscinas, acesso_analista_semanal,
                                  analise_semanal_no_formulario)


SUGESTOES_PISCINAS = ['Piscina Adulto', 'Piscina Infantil', 'Piscina principal']


def renderizar_analise_semanal_operador(st, salvar):
    ss = st.session_state
    if not acesso_analista_semanal(ss):
        ss.pop('op_analise_semanal_bem_star', None)
        ss.pop('_op_analise_semanal_widget', None)
        return
    ss['_op_analise_semanal_widget'] = analise_semanal_no_formulario(ss)

    def escolher():
        if acesso_analista_semanal(ss):
            ss['op_analise_semanal_bem_star'] = bool(ss.get('_op_analise_semanal_widget'))
            salvar()

    st.checkbox('Usar esta visita como análise semanal no mensal Bem Star',
                key='_op_analise_semanal_widget', on_change=escolher,
                help='Informe a data real da visita. Segunda-feira é o padrão, mas você pode '
                     'marcar a coleta semanal em outro dia. Desmarque nas visitas apenas de rotina.')
    st.caption('A análise entra no mês da data da visita, com as medições de cada piscina. '
               'A escolha fica guardada no rascunho e na visita salva.')


def renderizar_piscinas(st, lancamentos, cadastradas=()):
    """Piscinas do relatório. Nomes do cadastro e das visitas prevalecem.

    Sugestões genéricas e seleções restauradas que forem outra grafia de uma
    piscina existente ("Piscina adulta" × "Piscina Adulto") são convertidas
    para o nome existente, sem criar uma piscina extra no PDF.
    """
    ss = st.session_state
    # Protege os campos das piscinas recolhidas da limpeza automática de widgets.
    for k in list(ss):
        if k.startswith('bs_agenda_pool_'):
            ss[k] = ss[k]
    cadastradas = canonizar_nomes_piscinas(cadastradas)
    existentes = list(dict.fromkeys(l['piscina'] for l in expandir(lancamentos, cadastradas)))
    conhecidas = canonizar_nomes_piscinas(cadastradas + existentes)
    st.markdown('### Piscinas do relatório')
    st.caption('Cada piscina tem leituras, dosagens e gráficos próprios. Classifique a piscina principal como Adulto ou Infantil quando aplicável.')
    extras = st.text_input('Outras piscinas (separe os nomes por ponto e vírgula)', key='bs_mensal_outras',
                           placeholder='Ex.: Piscina aquecida; Spa')
    anteriores = list(ss.get('bs_mensal_piscinas') or [])
    selecionadas = canonizar_nomes_piscinas(anteriores, conhecidas)
    # A normalização do nome também precisa levar os campos da agenda.
    # As chaves antigas são mantidas; um valor já preenchido no nome atual prevalece.
    for anterior in anteriores:
        atual = canonizar_nomes_piscinas([anterior], conhecidas)[0]
        if token_pool(anterior) == token_pool(atual):
            continue
        prefixo = 'bs_agenda_pool_' + token_pool(anterior) + '_'
        for chave in list(ss):
            if chave.startswith(prefixo):
                destino = 'bs_agenda_pool_' + token_pool(atual) + '_' + chave[len(prefixo):]
                if ss.get(destino) in (None, ''):
                    ss[destino] = ss[chave]
        tipo_antigo = 'bs_mensal_tipo_' + token_pool(anterior)
        tipo_atual = 'bs_mensal_tipo_' + token_pool(atual)
        if ss.get(tipo_atual) not in TIPOS and ss.get(tipo_antigo) in TIPOS:
            ss[tipo_atual] = ss[tipo_antigo]
    if not selecionadas:
        selecionadas = list(existentes or cadastradas or SUGESTOES_PISCINAS[:1])
    sugestoes = SUGESTOES_PISCINAS if not conhecidas else []
    opcoes = canonizar_nomes_piscinas(conhecidas + sugestoes + [n.strip() for n in extras.split(';') if n.strip()]
                                      + selecionadas, conhecidas + selecionadas)
    if list(ss.get('bs_mensal_piscinas') or []) != selecionadas:
        ss['bs_mensal_piscinas'] = selecionadas
    st.multiselect('Piscinas para preencher na agenda', opcoes, key='bs_mensal_piscinas')
    nomes = canonizar_nomes_piscinas(existentes + list(ss.get('bs_mensal_piscinas') or []), conhecidas)
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
    originais = {}
    for linha in expandir(lancamentos, ss.get('bs_mensal_piscinas') or []):
        originais.setdefault((linha['data'], chave_piscina(linha['piscina'])), linha)
    for data in datas:
        with st.expander(f'Visita — {data}', expanded=False):
            for nome in ss.get('bs_mensal_piscinas', []) or []:
                st.markdown(f'**{nome} · {ss.get("bs_mensal_tipo_" + token_pool(nome), "Não informado")}**')
                base = originais.get((data, chave_piscina(nome)), {})
                # Somente a piscina de mesmo nome preenche seus campos.
                for k, _, _ in PARAMETROS:
                    preencher(chave_campo(nome, data, k), base.get(k, ''))
                preencher(chave_campo(nome, data, 'observacao'), base.get('observacao', ''))
                st.caption('Os gráficos mensais usam as visitas escolhidas por Thyago como análise semanal. '
                           'A segunda-feira é o padrão, com possibilidade de coleta em outro dia.')
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
                               accept_multiple_files=True, max_upload_size=10, key=f'_bs_upload_{cid}_{epoch}')
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
                st.caption('Imagem salva na nuvem. Será recuperada ao gerar o PDF.')
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
    renderizar_capa_mensal(st)


def renderizar_capa_mensal(st):
    """Escolha mensal isolada pelo contexto, salva junto dos complementos existentes."""
    ss = st.session_state
    cid = contexto_id(ss.get('_bs_rel_contexto'))
    fotos = {f['sha256']: f for f in ss.get('bs_mensal_fotos', [])
             if f.get('contexto') == cid}
    chave = 'bs_mensal_capa_hash'
    escolha = ss.get(chave, '')
    ss[chave] = escolha if escolha in fotos else ''
    st.selectbox('Foto da capa do relatório mensal', [''] + list(fotos), key=chave,
                 format_func=lambda valor: fotos[valor]['nome'] if valor else 'Sem foto na capa')
    st.caption('Escolha uma imagem do período anexada em Adicionar imagens. Ela também permanece no registro fotográfico. A escolha é salva com os complementos deste cliente e mês.')
    escolhida = fotos.get(ss[chave])
    if escolhida and escolhida.get('bytes'):
        st.image(escolhida['bytes'], width=300)


# Definido após as funções para sinalizar que a revisão terminou de carregar.
REVISAO_RESTAURACAO = 1
