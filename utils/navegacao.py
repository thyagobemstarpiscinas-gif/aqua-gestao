"""Navegação do escritório por âncoras: não troca páginas nem recria formulários."""
from html import escape


COMUNS = (
    ('Gestão', (
        ('nav-comercial', 'Gestão comercial'),
        ('nav-operadores', 'Operadores e acessos'),
        ('nav-diagnostico', 'Diagnóstico do sistema'),
    )),
)

MENUS = {
    'aqua_gestao': (
        ('Relatórios', (
            ('sec-relatorio-rt', 'Relatório mensal RT'),
            ('sec-preview-relatorio', 'Prévia do relatório'),
            ('nav-consolidados', 'Relatórios consolidados'),
        )),
        ('Clientes e documentos', (
            ('nav-clientes', 'Clientes e piscinas'),
            ('sec-formulario', 'Contrato de RT'),
            ('sec-aditivo-cnpj', 'Aditivos e renovação'),
            ('nav-aqua-manutencao', 'Contrato de manutenção'),
            ('sec-termos-ciencia', 'Termos de ciência'),
            ('sec-pops-adaptativos', 'Caderno de POPs'),
            ('nav-dossie', 'Dossiê de segurança química'),
            ('nav-documentos', 'Documentos e vencimentos'),
            ('nav-envio', 'Enviar documentos'),
        )),
    ) + COMUNS,
    'bem_star': (
        ('Relatórios', (
            ('nav-mensal-bs', 'Relatório mensal operacional'),
        )),
        ('Clientes e documentos', (
            ('nav-clientes', 'Clientes e piscinas'),
            ('nav-contrato-bs', 'Contrato de manutenção'),
            ('nav-proposta-bs', 'Proposta comercial'),
        )),
    ) + COMUNS,
}


def empresa_do_escritorio(estado):
    """Só expõe atalhos administrativos após login; respeita a empresa fixa."""
    if estado.get('modo_atual') != 'escritorio' or not estado.get('admin_logado'):
        return None
    empresa = estado.get('admin_empresa_fixa') or estado.get('empresa_ativa')
    return empresa if empresa in MENUS else None


def links_menu(empresa):
    if empresa not in MENUS:
        return ()
    return (('nav-inicio', 'Área de trabalho'), ('nav-painel', 'Indicadores'),
            *(link for _, links in MENUS[empresa] for link in links))


def ancora(st, identificador):
    st.markdown(f'<div id="{escape(identificador, quote=True)}" class="aq-nav-anchor"></div>',
                unsafe_allow_html=True)


def _link(alvo, rotulo, classe='aq-nav-link'):
    return (f'<a class="{classe}" href="#{escape(alvo, quote=True)}" target="_self">'
            f'{escape(rotulo)}</a>')


def html_menu(empresa):
    if empresa not in MENUS:
        return ''
    nome = 'Bem Star Piscinas' if empresa == 'bem_star' else 'ACQUA GESTÃO'
    sigla = 'BS' if empresa == 'bem_star' else 'AG'
    grupos = []
    for titulo, links in MENUS[empresa]:
        aberto = ' open' if titulo == 'Relatórios' else ''
        grupos.append(f'<details class="aq-nav-group"{aberto}><summary>{escape(titulo)}</summary>'
                      + ''.join(_link(*link) for link in links) + '</details>')
    return (f'<nav class="aq-nav {empresa}" aria-label="Navegação do escritório">'
            f'<div class="aq-nav-brand"><span class="aq-nav-monogram">{sigla}</span>'
            f'<div><strong>{nome}</strong><small>Acesso administrativo</small></div></div>'
            + _link('nav-inicio', 'Área de trabalho') + _link('nav-painel', 'Indicadores')
            + ''.join(grupos) + '</nav>')


def renderizar_menu(st, estado):
    empresa = empresa_do_escritorio(estado)
    if empresa:
        st.markdown(html_menu(empresa), unsafe_allow_html=True)


def renderizar_inicio(st, estado, logo=None):
    empresa = empresa_do_escritorio(estado)
    if not empresa:
        return
    ancora(st, 'nav-inicio')
    bs = empresa == 'bem_star'
    nome = 'Bem Star Piscinas' if bs else 'ACQUA GESTÃO'
    titulo = 'Operação e manutenção' if bs else 'Controle técnico de piscinas'
    cabecalho = (f'<div class="aq-workspace {empresa}"><span class="aq-eyebrow">{nome} / ESCRITÓRIO</span>'
                 f'<h1>{titulo}</h1><p>Acesse seus relatórios, clientes e documentos.</p></div>')
    if logo:
        texto, marca = st.columns([5, 1])
        with texto:
            st.markdown(cabecalho, unsafe_allow_html=True)
        with marca:
            st.image(str(logo), width=135)
    else:
        st.markdown(cabecalho, unsafe_allow_html=True)
    mensal = ('nav-mensal-bs', 'Relatório mensal', 'Consolide as visitas e análises do período.') if bs else (
        'sec-relatorio-rt', 'Relatório mensal RT', 'Prepare medições, parecer e fotos do período.')
    cards = (mensal, ('nav-clientes', 'Clientes e piscinas', 'Consulte e atualize os cadastros.'),
             ('nav-comercial', 'Gestão comercial', 'Acompanhe contatos, propostas e retornos.'))
    html = '<div class="aq-shortcuts" role="navigation" aria-label="Atalhos principais">'
    for indice, (alvo, label, descricao) in enumerate(cards, 1):
        html += (f'<a class="aq-shortcut" href="#{alvo}" target="_self">'
                 f'<span class="aq-shortcut-number">0{indice}</span><span class="aq-shortcut-arrow" aria-hidden="true">↗</span>'
                 f'<strong>{label}</strong><span>{descricao}</span></a>')
    st.markdown(html + '</div>', unsafe_allow_html=True)
    st.caption('Os atalhos levam à seção nesta página. Os campos em preenchimento continuam no lugar.')


def aplicar_estilo(st):
    ativos = '\n'.join(
        f'body:has(#{alvo}:target) .aq-nav a[href="#{alvo}"] '
        '{background:#e7eff8;color:#123f6b;border-left-color:#1a5b94;font-weight:650;}'
        for alvo in sorted({alvo for empresa in MENUS for alvo, _ in links_menu(empresa)})
    )
    st.markdown('''<style>
    .aq-nav-anchor {scroll-margin-top:5rem;}
    [data-testid="stMain"] {scroll-behavior:smooth;}
    .aq-nav {color:#22374b;font-size:.92rem;background:#f7faff;border:1px solid #dce5ee;border-radius:14px;padding:14px;}
    .aq-nav-brand {display:flex;align-items:center;gap:12px;padding:4px 0 22px;}
    .aq-nav-brand strong {display:block;font-size:.98rem;letter-spacing:.03em;}
    .aq-nav-brand small {display:block;color:#64748b;font-size:.76rem;margin-top:3px;}
    .aq-nav-monogram {display:grid;place-items:center;width:42px;height:42px;border-radius:12px;background:#123f6b;color:white;font-weight:750;}
    .bem_star .aq-nav-monogram {background:#174c7a;box-shadow:inset 0 -4px #e5b747;}
    .aq-nav a.aq-nav-link {display:block;padding:10px 12px;margin:3px 0;border-radius:7px;border-left:3px solid transparent;color:#304b65;text-decoration:none;line-height:1.35;}
    .aq-nav a.aq-nav-link:hover {background:#e7eff8;color:#123f6b;}
    .aq-nav-group {border-top:1px solid #dce5ee;margin-top:12px;padding-top:10px;}
    .aq-nav-group summary {cursor:pointer;color:#60758a;font-size:.75rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:8px 3px;}
    .aq-nav a:focus-visible,.aq-nav summary:focus-visible,.aq-shortcut:focus-visible {outline:3px solid #367bbc;outline-offset:2px;}
    .aq-workspace {padding:20px;background:#f7faff;border:1px solid #dce5ee;border-radius:14px;}
    .aq-eyebrow {font-size:.73rem;font-weight:700;letter-spacing:.12em;color:#54728e;}
    .aq-workspace h1 {font-size:clamp(1.65rem,3vw,2.45rem);letter-spacing:-.035em;color:#153a5d;padding:10px 0 5px;}
    .aq-workspace p {color:#63798d;font-size:.97rem;margin:0 0 10px;}
    .aq-shortcuts {display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin:12px 0 6px;}
    a.aq-shortcut {display:block;position:relative;padding:20px;border:1px solid #d9e4ee;border-radius:14px;background:#fff;color:#163d60;text-decoration:none;box-shadow:0 4px 16px #183c6010;transition:background .15s,border-color .15s;}
    a.aq-shortcut:hover {background:#f2f7fc;border-color:#7ba1c5;}
    .aq-shortcut-number {display:block;color:#708ca5;font-size:.76rem;font-weight:700;letter-spacing:.06em;margin-bottom:19px;}
    .aq-shortcut-arrow {position:absolute;top:16px;right:19px;font-size:1.25rem;color:#567d9f;}
    .aq-shortcut strong {display:block;font-size:1.05rem;margin-bottom:6px;}
    .aq-shortcut span:last-child {display:block;color:#62788c;font-size:.85rem;line-height:1.5;}
    @media(max-width:760px){.aq-shortcuts{grid-template-columns:1fr;gap:8px;}a.aq-shortcut{padding:15px;}.aq-shortcut-number{margin-bottom:8px;}}
    @media(prefers-reduced-motion:reduce){[data-testid="stMain"]{scroll-behavior:auto;}a.aq-shortcut{transition:none;}}
    </style>''' + '<style>' + ativos + '</style>', unsafe_allow_html=True)
