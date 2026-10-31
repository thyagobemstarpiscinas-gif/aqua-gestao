"""PDF operacional Bem Star: piscinas independentes, dados informados e gráficos vetoriais."""
from collections import OrderedDict
from html import escape
from pathlib import Path
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer,
                               PageBreak, KeepTogether, Image, Flowable)
from utils.bem_star_mensal import (PARAMETROS, AVISO, REFERENCIA, expandir, numero, data_ordem)
from utils.capa_bem_star import desenhar_capa_bem_star

AZUL = colors.HexColor('#104888')
AGUA = colors.HexColor('#0090D0')
CLARO = colors.HexColor('#EDF6FC')
CINZA = colors.HexColor('#536779')
LINHA = colors.HexColor('#D5E3ED')
MESES = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']


def serie_grafico(linhas, campo):
    """Preserva todas as datas, inclusive as lacunas; nunca interpola ausências."""
    return [(l['data'], numero(l.get(campo))) for l in sorted(linhas, key=lambda l: data_ordem(l['data']))]


class GraficoParametro(Flowable):
    def __init__(self, linhas, campo, titulo, unidade, width=82*mm, height=39*mm):
        super().__init__()
        self.width, self.height = width, height
        self.serie = serie_grafico(linhas, campo)
        self.titulo, self.unidade = titulo, unidade

    def draw(self):
        c, w, h = self.canv, self.width, self.height
        c.saveState()
        c.setFillColor(CLARO); c.roundRect(0, 0, w, h, 4, fill=1, stroke=0)
        c.setFillColor(AZUL); c.setFont('Helvetica-Bold', 8)
        c.drawString(9, h-14, self.titulo + (f' ({self.unidade})' if self.unidade else ''))
        valores = [v for _, v in self.serie if v is not None]
        if not valores:
            c.setFont('Helvetica', 8); c.setFillColor(CINZA)
            c.drawCentredString(w/2, h/2-5, 'Sem medição informada')
            c.restoreState(); return
        ymin, ymax = min(valores), max(valores)
        amplitude = ymax-ymin
        pad = amplitude*.18 if amplitude else max(abs(ymin)*.04, .1)
        ymin, ymax = ymin-pad, ymax+pad
        x0, x1, y0, y1 = 34, w-14, 23, h-29
        c.setFont('Helvetica', 6.5)
        for i in range(3):
            y = y0+(y1-y0)*i/2; val = ymin+(ymax-ymin)*i/2
            c.setStrokeColor(LINHA); c.setLineWidth(.4); c.line(x0, y, x1, y)
            c.setFillColor(CINZA); c.drawRightString(x0-4, y-2, f'{val:.1f}'.replace('.', ','))
        primeiro, ultimo = data_ordem(self.serie[0][0]), data_ordem(self.serie[-1][0])
        span = max((ultimo-primeiro).days, 1)
        def xdata(data):
            return x0+(x1-x0)*((data_ordem(data)-primeiro).days/span) if primeiro != ultimo else (x0+x1)/2
        anterior = None
        for data, valor in self.serie:
            x = xdata(data)
            if valor is None:
                anterior = None
                c.setStrokeColor(CINZA); c.line(x-2, y0-3, x+2, y0+1); c.line(x-2, y0+1, x+2, y0-3)
                continue
            y = y0+(y1-y0)*(valor-ymin)/(ymax-ymin)
            c.setStrokeColor(AGUA); c.setLineWidth(1.2)
            if anterior:
                c.line(*anterior, x, y)
            c.setFillColor(AZUL); c.circle(x, y, 2, fill=1, stroke=0)
            anterior = (x, y)
        indices = sorted(set([0, len(self.serie)//2, len(self.serie)-1]))
        c.setFont('Helvetica', 6.5); c.setFillColor(CINZA)
        for i in indices:
            data = self.serie[i][0]
            c.drawCentredString(xdata(data), 10, data[:5])
        c.restoreState()


def gerar_pdf(output_path, ctx, *, cnpj, responsavel, logo_path):
    try:
        from calendar import monthrange
        from utils.relogio import hoje_brasilia
        path = Path(output_path); path.parent.mkdir(parents=True, exist_ok=True)
        mes, ano = int(ctx['mes']), int(ctx['ano'])
        nome_mes = MESES[mes]
        cliente = str((ctx.get('dados_cliente') or {}).get('nome') or ctx['cliente'])
        emissao = hoje_brasilia()
        ini, fim = f'01/{mes:02d}/{ano}', f'{monthrange(ano, mes)[1]:02d}/{mes:02d}/{ano}'
        linhas = expandir(ctx.get('lancamentos') or [])
        grupos = OrderedDict()
        for l in linhas:
            for campo, label, _ in PARAMETROS:
                if str(l.get(campo) or '').strip() and numero(l[campo]) is None:
                    raise ValueError(f'Valor inválido em {label}: {l["piscina"]}, {l["data"]}.')
            if data_ordem(l['data']).month != mes or data_ordem(l['data']).year != ano:
                raise ValueError('Leitura fora do período selecionado.')
            grupos.setdefault(l['piscina'], []).append(l)
        for nome in ctx.get('piscinas') or []:
            grupos.setdefault(nome, [])
        styles = {
            'body': ParagraphStyle('body', fontName='Helvetica', fontSize=9, leading=13, textColor=colors.HexColor('#263B4E'), spaceAfter=5),
            'small': ParagraphStyle('small', fontName='Helvetica', fontSize=7.5, leading=10.2, textColor=CINZA),
            'h1': ParagraphStyle('h1', fontName='Helvetica-Bold', fontSize=19, leading=23, textColor=AZUL, spaceAfter=10, keepWithNext=True),
            'h2': ParagraphStyle('h2', fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=AZUL, spaceBefore=8, spaceAfter=6, keepWithNext=True),
            'th': ParagraphStyle('th', fontName='Helvetica-Bold', fontSize=7, leading=9, textColor=colors.white),
            'td': ParagraphStyle('td', fontName='Helvetica', fontSize=7.4, leading=9.5, textColor=AZUL),
        }
        def p(txt, style='body'):
            return Paragraph(escape(str(txt or '')).replace('\n', '<br/>'), styles[style])
        def textos(txt):
            return [p(b.strip()) for b in re.split(r'\n\s*\n', str(txt or '')) if b.strip()]
        def fmt(v):
            n = numero(v)
            return '-' if n is None else f'{n:g}'.replace('.', ',')
        def tabela(rows, widths, header=True):
            t = Table([[p(c, 'th' if header and i == 0 else 'td') for c in row] for i, row in enumerate(rows)],
                      colWidths=widths, repeatRows=1 if header else 0, hAlign='LEFT')
            rules = [('VALIGN', (0,0), (-1,-1), 'TOP'), ('GRID', (0,0), (-1,-1), .35, LINHA),
                     ('LEFTPADDING',(0,0),(-1,-1),6), ('RIGHTPADDING',(0,0),(-1,-1),6),
                     ('TOPPADDING',(0,0),(-1,-1),5), ('BOTTOMPADDING',(0,0),(-1,-1),5)]
            if header:
                rules += [('BACKGROUND',(0,0),(-1,0), AZUL), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, CLARO])]
            t.setStyle(TableStyle(rules)); return t
        def caixa(titulo, texto):
            t = Table([[p(titulo, 'h2')], [p(texto)]], colWidths=[171*mm])
            t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),CLARO), ('LINEBEFORE',(0,0),(0,-1),3,AGUA),
                                   ('LEFTPADDING',(0,0),(-1,-1),12), ('RIGHTPADDING',(0,0),(-1,-1),12),
                                   ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),8)]))
            return t
        def capa(c, doc):
            desenhar_capa_bem_star(c, cliente=cliente, mes_nome=nome_mes, ano=str(ano), periodo_ini=ini,
                                   periodo_fim=fim, data_emissao=emissao, logo_path=logo_path, cnpj=cnpj)
        def rodape(c, doc):
            c.saveState(); w, h = A4
            c.setFillColor(AZUL); c.setFont('Helvetica-Bold', 8)
            c.drawString(19*mm, h-12*mm, 'BEM STAR PISCINAS | REGISTRO OPERACIONAL')
            c.setFillColor(CINZA); c.setFont('Helvetica', 7)
            c.drawRightString(w-19*mm, h-12*mm, f'{nome_mes} / {ano}')
            c.setStrokeColor(LINHA); c.line(19*mm, 16*mm, w-19*mm, 16*mm)
            c.setFont('Helvetica', 6.8)
            c.drawString(19*mm, 11*mm, 'Não é relatório de RT. Não constitui laudo ou ART.')
            c.drawRightString(w-19*mm, 11*mm, f'Página {doc.page}')
            c.restoreState()
        story = [Spacer(1,1), PageBreak(), p('O mês em perspectiva', 'h1'), p(f'{cliente} | {nome_mes} / {ano}', 'h2'), p(AVISO)]
        dias = len({l['data'] for l in linhas})
        story.append(tabela([['Datas com registro', 'Piscinas identificadas', 'Período'],
                              [str(dias), str(len(grupos)), f'{ini} a {fim}']], [48*mm, 51*mm,72*mm]))
        story.append(Spacer(1,5*mm))
        resumo = [['Piscina / setor', 'Tipo', 'Datas com leituras']]
        for nome, dados in grupos.items():
            tipos = list(dict.fromkeys(l['tipo_piscina'] for l in dados)) or [(ctx.get('tipos_piscinas') or {}).get(nome, 'Não informado')]
            cont = len({l['data'] for l in dados if any(numero(l.get(k)) is not None for k,_,_ in PARAMETROS)})
            resumo.append([nome, ', '.join(tipos), str(cont)])
        story.append(tabela(resumo, [81*mm,43*mm,47*mm]))
        story.extend([p('Como ler este relatório', 'h2'), p('As tabelas e os gráficos apresentam somente os valores informados, individualmente por piscina. '
                'A rotina de análise semanal completa é prevista às segundas-feiras; medições adicionais também são exibidas. '
                'Campos vazios permanecem sem resultado. Ausência de medição não significa resultado adequado.'),
                p('As linhas dos gráficos apenas conectam medições consecutivas informadas. Lacunas interrompem o traçado; '
                'não há estimativa de valores ausentes. Os gráficos mostram tendências, sem atestar potabilidade, segurança ou conformidade normativa.', 'small'),
                Spacer(1,5*mm), caixa('Manutenção bem registrada. Decisões mais bem apoiadas.',
                'O histórico organizado ajuda o condomínio a acompanhar os serviços, identificar recorrências e planejar intervenções. '
                'Um serviço de Responsabilidade Técnica, com escopo próprio, pode complementar essa rotina com interpretação especializada e orientação à gestão. '
                'Veja ao final como esse acompanhamento pode beneficiar o condomínio.')])
        for nome, dados in grupos.items():
            story.extend([PageBreak(), p(nome, 'h1')])
            tipo = next((l['tipo_piscina'] for l in dados), (ctx.get('tipos_piscinas') or {}).get(nome, 'Não informado'))
            story.append(p(f'Tipo: {tipo} | Leituras informadas por data', 'small'))
            rows = [['Data', 'pH', 'CL', 'CT', 'Alc.', 'Dur.', 'CYA']]
            for l in dados:
                rows.append([l['data']] + [fmt(l.get(k)) for k,_,_ in PARAMETROS])
            if not dados:
                rows.append(['Sem registro'] + ['-']*6)
            story.append(tabela(rows, [30*mm]+[23.5*mm]*6))
            story.extend([Spacer(1,2*mm), p('CL: cloro livre | CT: cloro total | Alc.: alcalinidade | Dur.: dureza cálcica | CYA: ácido cianúrico. '
                'Concentrações em mg/L; pH sem unidade. Traço: não medido ou não informado.', 'small'),
                p('Evolução dos parâmetros', 'h2')])
            graficos = [GraficoParametro(dados,k,label,unit,width=54.5*mm,height=36*mm) for k,label,unit in PARAMETROS]
            for i in range(0,6,3):
                t = Table([[graficos[i], graficos[i+1], graficos[i+2]]], colWidths=[57*mm]*3, hAlign='LEFT')
                t.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
                                      ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)]))
                story.append(t)
            story.append(p('Ponto azul: medição informada. X na base: ausência de medição em uma data registrada. '
                           'Escalas próprias por parâmetro; não compare a altura entre gráficos.', 'small'))
            story.append(p('Resumo numérico da piscina', 'h2'))
            stats = [['Parâmetro', 'Mínimo', 'Média', 'Máximo', 'Nº de leituras']]
            labels = {'cloro_livre':'Cloro Livre', 'cloro_total':'Cloro Total'}
            for k,label,_ in PARAMETROS:
                vals = [numero(l.get(k)) for l in dados if numero(l.get(k)) is not None]
                stats.append([labels.get(k,label)] + ([fmt(min(vals)), fmt(round(sum(vals)/len(vals),2)), fmt(max(vals)), str(len(vals))] if vals else ['-']*4))
            cc = [numero(l.get('cloro_total'))-numero(l.get('cloro_livre')) for l in dados
                  if numero(l.get('cloro_total')) is not None and numero(l.get('cloro_livre')) is not None
                  and numero(l.get('cloro_total')) >= numero(l.get('cloro_livre'))]
            stats.append(['Cloro Combinado'] + ([fmt(min(cc)),fmt(round(sum(cc)/len(cc),2)),fmt(max(cc)),str(len(cc))] if cc else ['-']*4))
            story.append(tabela(stats,[54*mm,28*mm,28*mm,28*mm,33*mm]))
        story.extend([PageBreak(), p('Registro da operação', 'h1')])
        for titulo, chave in [('Observações da equipe','obs_geral'),('Resumo e recomendações operacionais','parecer_tecnico')]:
            if ctx.get(chave):
                story.append(p(titulo,'h2')); story.extend(textos(ctx[chave]))
        registros = [l for l in linhas if l.get('observacao') or l.get('dosagens') or l.get('problemas') or l.get('dosagens_da_visita')]
        if registros:
            story.append(p('Ocorrências e dosagens registradas', 'h2'))
            doses_vistas = set()
            for l in registros:
                story.append(p(f'{l["data"]} | {l["piscina"]}', 'h2'))
                for campo in ['observacao','problemas']:
                    story.extend(textos(l.get(campo)))
                if l.get('dosagens_da_visita') and l['data'] not in doses_vistas:
                    doses_vistas.add(l['data'])
                    story.append(p('Dosagens da visita sem identificação de piscina: ' + str(l['dosagens_da_visita']), 'small'))
                for d in l.get('dosagens') or []:
                    story.append(p(' | '.join(str(d.get(k) or '') for k in ['produto','quantidade','unidade','finalidade']), 'small'))
        story.append(KeepTogether([p('Encerramento do Período','h2'),
                p(f'O presente dashboard consolida os registros operacionais informados pela Bem Star Piscinas Ltda '
                  f'para {cliente}, em {nome_mes} / {ano}. As imagens documentam aspectos visuais e não substituem análises.'),
                p(f'_______________________________________________\n{responsavel}\nSócio-Administrador / Responsável pela operação — Bem Star Piscinas Ltda\n'
                  f'Operador responsável: {ctx.get("operador") or "Não informado"}\nUberlândia/MG, {emissao}')]))
        fotos = [Path(f) for f in (ctx.get('fotos') or []) if Path(f).is_file()]
        if fotos:
            story.extend([PageBreak(), p('Registro fotográfico', 'h1'), p('Imagens das visitas e anexos selecionados para este cliente e mês.', 'small')])
            meta = ctx.get('fotos_metadados') or {}
            for fp in fotos:
                m = meta.get(str(fp), {})
                cap = ' | '.join(str(m.get(k)) for k in ['piscina','data','legenda'] if m.get(k)) or fp.stem.replace('_',' ').replace('-', ' ').title()[:150]
                img = Image(str(fp), width=164*mm, height=78*mm, kind='proportional')
                story.append(KeepTogether([Spacer(1,5*mm), img, p(cap, 'small')]))
        story.extend([PageBreak(), p('Mais apoio para cuidar\ndo patrimônio coletivo', 'h1'),
                p('Responsabilidade Técnica | serviço distinto do registro operacional', 'h2'),
                p('A rotina de manutenção mantém o histórico do trabalho realizado. O acompanhamento por profissional da Química habilitado '
                  'pode ampliar a interpretação desse histórico e apoiar decisões do condomínio, dentro de um escopo definido.'),
                caixa('O que o condomínio pode ganhar',
                  '• Interpretação das tendências da água e investigação de desvios recorrentes.\n'
                  '• Orientação à equipe e revisão das rotinas de medição e controle.\n'
                  '• Apoio ao planejamento de análises complementares e intervenções.\n'
                  '• Documentação técnica e comunicação mais clara com a administração.'),
                Spacer(1,5*mm),p('Conheça as possibilidades de acompanhamento', 'h2'),
                p('Converse com nossa equipe sobre o serviço de RT da Aqua Gestão. Podemos avaliar as necessidades do condomínio '
                  'e apresentar uma proposta com escopo, frequência e documentação próprios. A apresentação é informativa e não condiciona '
                  'a continuidade dos serviços de manutenção à contratação desse serviço.'),
                p('Referência e alcance deste documento', 'h2'), p(REFERENCIA),
                p('A resolução trata da ART para o tratamento químico e controle da qualidade da água de piscinas de uso público e coletivo. '
                  'O art. 7º prevê fiscalização preventiva e orientativa dos condomínios residenciais. '
                  'O enquadramento e as exigências aplicáveis devem ser avaliados conforme o estabelecimento, o serviço e a legislação pertinente. '
                  'Este relatório não comprova a assunção de RT nem a emissão de ART; sua identificação como operacional não dispensa exigências legais eventualmente aplicáveis.'),
                p('Fonte: texto consolidado publicado pelo CRQ-MG — www.crqmg.org.br/legislacao/read/350. Consulta: 31/10/2026.', 'small'),
                Spacer(1,6*mm),caixa('Transparência para a gestão', AVISO)])
        doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=19*mm, rightMargin=19*mm,
                                topMargin=22*mm, bottomMargin=21*mm,
                                title=f'Relatório operacional Bem Star — {cliente} — {mes:02d}/{ano}', author='Bem Star Piscinas Ltda')
        doc.build(story, onFirstPage=capa, onLaterPages=rodape)
        return True, ''
    except Exception as exc:
        return False, str(exc)
