# Correções de emissão — 01/11/2026

Base remota: main f69c3503e54280994d9dbc89a12a9c44d3911fa7 (árvore a7e7f76bc16ec6374a0656e50f17b343e2c021f2).

## Comportamento final

- Visita do operador ACQUA e supervisão do RT priorizam `acqua_gestao_logo_rt.png`; o PDF de visita Bem Star incorpora `bem_star_logo.png` na capa, preservando as marcas separadas.
- Corrigido TypeError real no PDF Bem Star: ParagraphStyle recebia fontName duas vezes. Os três geradores foram executados com dados fictícios e os PDFs foram renderizados.
- CFQ 332/2025 referenciada com alteração 345/2026. NBR 10339 identificada como projeto, execução e manutenção; NBR 11238 cancelada não é mais o título do checklist.
- Removida a troca incorreta de NBR 10818 por NBR 10339 no normalizador do DOCX. Corrigidos também textos do modelo legado, referência indevida à acreditação pela ANVISA e títulos de faixas no PDF de supervisão.
- Faixas numéricas permanecem operacionais, sujeitas à revisão do RT, sem declaração de certificação à nova edição da norma de qualidade da água. O catálogo registra NBR 10818:2016 cancelada com revisão em 21/09/2026. Não foi adquirido/consultado o texto integral da revisão; não se inventaram novos limites ou exigências municipais.
- Novos contextos de EPI começam como Não informado; checklist operacional começa como Não verificado e admite N/A. Valores existentes dos rascunhos são preservados. Chaves legadas mantidas para compatibilidade.
- CA vazio é Não informado, nunca Conforme. Checklist NR-26 não infere OK da ausência de observação. Não informado e N/A não são convertidos automaticamente em Não no quadro de EPI.
- Nenhuma migração, exclusão de cadastro, alteração de dosagem ou alteração de abas de Sheets. Sem desenvolvimento da assinatura nesta etapa.

## Validação

- 145 testes locais aprovados: segurança de acesso, formulário do operador, contexto/mensal Bem Star, fotos, cadastro de piscinas, regras RT e três geradores reais de PDF.
- 1 teste adicional de regressão DOCX legado aprovado, juntamente com os três testes PDF: referências atualizadas e ausência de alegações antigas.
- 6 AppTests selecionados aprovados, zero falhas/erros/skips: `TestRelatorioRTProfissional` (2), `CadastroUI` (2) e `WidgetsMensal` (2). Dados fictícios, planilhas falsas e rede bloqueada. JSON de execução: /tmp/acqua_ui_selecionados.json (evidência efêmera, resumo persistido aqui).
- Total de cenários distintos nesta rodada: 152. Suíte completa de 306 não repetida nesta etapa.
- Sintaxe e git diff --check aprovados. Servidor Streamlit local iniciado na porta 8501.
- Logos nos três PDFs conferidas por pixels das imagens incorporadas. Mensal fictício gerado com seis páginas e DOCX; capa, tabelas, gráficos e foto inspecionados. Revisão DOCX adicional cobre o texto final do normalizador.
- Esta rodada não realiza gravações fictícias nem testes de permissões com PINs reais em produção. A comprovação funcional de escrita e isolamento usa AppTest; não é garantia de ausência absoluta de defeitos.

## Uso

RT: PIN pessoal RT no modo campo para registrar visita técnica; botão Abrir relatório mensal RT para revisar e emitir o mensal. Operadores usam PIN individual e só seus condomínios autorizados. Revisar cada checklist antes de emitir. Situação da ART continua selecionável; assinatura digital fica para etapa posterior.

## Fontes oficiais consultadas

- CRQ-MG, Resolução CFQ 332/2025: https://www.crqmg.org.br/legislacao/read/350
- CRQ-MG, Resolução CFQ 345/2026: https://www.crqmg.org.br/legislacao/read/363
- ABNT Catálogo: https://www.abntcatalogo.com.br/ — NBR 10339:2018 Versão Corrigida:2019 (em vigor); NBR 11238:1990 (cancelada em 19/09/2018); NBR 10818:2016 (cancelada com revisão em 21/09/2026).
- MTE, NR-6 atualizada 2025 e NR-26, portal de normas regulamentadoras vigentes. Aplicação conforme atividade e riscos, sem presumir documentação ou inspeção.
