# ACQUA GESTÃO — Correção dos defeitos da homologação do PR #14 (01/11/2026)

Branch `fix/rt-abnt-10818-2026`, a partir de `f0851ba95608d1415977f09b5b6802dea6a3f159`.
PR #14 permanece em rascunho. Sem merge, sem alteração de produção, Secrets,
credenciais, dados reais ou lógica de dosagem; sem assinatura digital.
Todos os testes usam dados fictícios, Google falso (escopo `drive.file`) e rede bloqueada.

**A homologação NÃO está aprovada**: depende de nova verificação independente no
Streamlit Cloud.

## Reprodução e causa de cada defeito

Os testes novos (`tests/test_homologacao_pr14.py` e
`tests/test_homologacao_pr14_apptest.py`) foram executados também sobre o código
original `f0851ba`: os 21 falham lá e passam com a correção.

| # | Defeito | Causa encontrada | Correção |
|---|---|---|---|
| 1 | PDF individual sem a foto | `gerar_pdf_relatorio_visita_bem_star` gravava cada foto num arquivo temporário e chamava `os.unlink` logo após criar o `Image`; o ReportLab só lê o arquivo em `doc.build()` → `FileNotFoundError` engolido por `except: pass`. Regerado a partir do Sheets, o PDF Bem Star também não buscava as fotos no Drive. | Foto mantida em memória (`BytesIO`) no próprio flowable até o fim do build, com orientação EXIF e proporção. O PDF Bem Star usa a mesma resolução de fotos do PDF ACQUA (base64 → Drive → pasta local), apenas quando o lançamento declara fotos. |
| 2 | Visita duplicada no PDF mensal | O modo campo grava a cópia local **antes** de o Sheets gerar o ID; a cópia local ficava sem `id_visita`. O relatório Bem Star unia local + Sheets comparando o JSON inteiro, então as duas cópias da V00005 entravam. | `sheets_salvar_lancamento_campo` devolve o ID ao lançamento (cópia local passa a tê-lo). Nova `deduplicar_lancamentos`: identidade por ID; para cópias antigas sem ID, data + `salvo_em` + operador. Campos são combinados (Sheets prevalece; base64 local é mantido). IDs diferentes nunca são unidos. |
| 3 | "Piscina adulta" além de "Piscina Adulto" | A tela do relatório mensal Bem Star sugeria e pré-selecionava o nome fixo "Piscina adulta"; agenda, união e PDF comparavam nomes exatos, e o PDF criava seção para toda piscina selecionada. | `chave_piscina` (caixa, acentos, espaços, adulta/adulto, infantis/infantil) e `canonizar_nomes_piscinas`, com prioridade para o nome do cadastro. "Piscina Adulto 2", "Piscina Infantil", "Piscina aquecida" continuam distintas. Sugestões genéricas só sem piscinas conhecidas. |
| 4 | Gama na empresa errada | `_carregar_dados_cliente_local`/`_enriquecer_cliente_com_dados_locais`: qualquer `dados_condominio.json` local (criado por visita, rascunho ou relatório) sem `empresa` assumia "Aqua Gestão" e **sobrescrevia** a coluna M do Sheets. Na planilha de homologação isso movia Gama **e Épsilon** para a Aqua e esvaziava a lista Bem Star. A lista do operador também reinseria a pasta local depois do filtro. | Coluna M do Sheets é a fonte principal; JSON local só vale quando declara empresa/serviços; sem nada, padrão legado inalterado. Pastas locais de clientes da planilha não são reinseridas na lista do operador. Painel de vencimentos, histórico e Central de Envio (Aqua) filtram pastas por empresa; pastas sem definição continuam visíveis (nada é escondido nem apagado). |
| 5 | Restauração do rascunho RT | No harness (planilha falsa, outra sessão, outro condomínio/período) a restauração em si concluiu; o defeito reproduzido foi a tela: o seletor "Carregar dados de cliente cadastrado" continuava no cliente anterior, e um "Carregar" seguinte misturava a tabela restaurada com outro cliente. | O seletor passa a acompanhar o condomínio restaurado. Teste ponta a ponta: rascunho Alfa 09/2026 restaurado sobre Beta 11/2026 → dados, período, identidade e seletor de Alfa; a edição de Beta é preservada como outra versão. **Pendência**: a mensagem exata de bloqueio vista no Cloud não foi reproduzida; ver abaixo. |
| 6 | Erro `csr_sel` e status inconsistente | (a) O botão "📄 Gerar relatório Bem Star (PDF)" estava fora do bloco que define `csr_sel`, e o tratamento avaliava `csr_sel`/`csr_mes`/`csr_ano` como padrão de `.get()` → `NameError: csr_sel` quando a lista Bem Star vinha vazia (situação causada pelo defeito 4). (b) "Status geral da água": o seletor tem chave própria, o Streamlit ignora `index=` e o valor antigo sobrescrevia, no fim da execução, o status calculado pelo parecer automático. | (a) Botão só com cliente listado e usando os dados do contexto gerado. (b) `rel_status_agua` é a fonte única, sincronizada antes do seletor; o parecer automático roda como callback (status aparece na mesma execução); escolha manual continua valendo. Painel "lançamento(s) encontrado(s)" passa a contar 1 para a V00005. |

Ajustes visuais encontrados na inspeção dos PDFs fictícios (PDF de visita Bem
Star): bloco de assinaturas sem quebras de linha (`\n` em `Paragraph`) e título
"PARECER OPERACIONAL" isolado no fim da página.

## Dados da homologação

Nenhuma rotina apaga ou reescreve visitas, fotos ou pastas: a V00005 e sua foto
no Drive continuam como estão; a cópia local antiga sem ID é reconhecida pela
assinatura de gravação e não precisa de migração.

## Pendências

- Defeito 5: confirmar no Streamlit Cloud com a mensagem/captura original. Se o
  bloqueio persistir, registrar o texto exibido (ex.: "Troca bloqueada…",
  "Rascunho de outra empresa…", "Fotografias pertencem a outro…").
- Nova verificação independente no Streamlit Cloud (homologação não aprovada).

---

# Bloqueios finais (após `1a2eb61`)

A homologação real confirmou as correções acima e apontou dois defeitos restantes.
Os testes novos (`tests/test_homologacao_pr14_bloqueios.py` e
`tests/test_homologacao_pr14_bloqueios_apptest.py`) foram executados sobre
`1a2eb61`: os que reproduzem os defeitos falham lá e passam com a correção.

## 1. Restauração/troca do rascunho RT bloqueada

**Reprodução** (app real, planilha falsa): restaurar o rascunho de Alfa, trocar
novembro → outubro com a planilha respondendo erro (cota/rede, como acontece no
Sheets), restaurar o rascunho de Beta. Resultado em `1a2eb61`: "Troca bloqueada:
não foi possível preservar o rascunho…", o mês volta para novembro (reversão da
transação) e, a partir daí, todo salvamento automático falha com
`ConflitoRascunho`.

**Causa**: `_guardar_rascunho_rt_antes_da_troca` só deixava a troca seguir se o
envio à planilha fosse confirmado na hora. Uma troca de período ou condomínio
só tinha alternativa na restauração explícita, e mesmo essa dependia do
estado da chave da sessão. O conteúdo, porém, já estava gravado de forma
atômica no registro local (`utils/rascunho_sync`), que é a cópia autoritativa
reenviada depois.

**Correção**: nova `_relatorio_rt_preservar_localmente`. Se o envio não for
confirmado, a troca segue somente quando o conteúdo atual já está num registro
local não descartado (mesmo hash), ou quando foi gravado agora numa chave nova
e conferido por releitura. Nenhuma cópia existente é alterada. A pendência é
reenviada automaticamente quando a planilha volta. **Falha da gravação local
continua bloqueando** (proteção contra perda de dados mantida, coberta por teste).

## 2. Fotografia ausente no PDF mensal Bem Star

**Cadeia verificada**: o Sheets guarda no payload da V00005 o nome do arquivo
(`fotos_antes`) e o ID do Drive (`fotos_antes_ids`), sem base64. A importação e a
deduplicação preservam os dois campos. O PDF individual resolve as fotos pelo
Drive. O mensal (`_coletar_fotos_bem_star_preview`) procurava apenas o arquivo
na pasta local `fotos_campo`, que não existe num servidor do Streamlit Cloud
reiniciado. Resultado: zero fotos, a página "Registro fotográfico" omitida (PDF
de 5 páginas) e nenhuma mensagem.

**Correção**: visitas sem arquivo local passam a usar o Drive (bytes originais)
e, na falta deste, a cópia base64 do lançamento local. Cada imagem é gravada em
JPEG numa pasta persistente do cliente (`_fotos_relatorio_mensal_bs/<sha256>.jpg`),
disponível até o fim da montagem, com deduplicação por conteúdo e legenda
"Visita V00005 — Antes do tratamento". Foto irrecuperável gera aviso na tela
("Registro fotográfico incompleto — V00005: …"), em vez de silêncio. Nenhuma
visita, linha do Sheets ou arquivo do Drive é alterado ou duplicado.
