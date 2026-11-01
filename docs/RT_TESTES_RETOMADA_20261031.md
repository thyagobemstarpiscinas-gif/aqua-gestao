# Retomada dos testes RT — 31/10/2026

## Execução local independente

Código: `5eaa7304872113a25330f00343c8fe7d658c3834` (PR #11).

Comando: `python tests/rodar_isolado.py --modules test_rt_relatorio test_rt_relatorio_apptest`

Resultado: **17 testes aprovados**, 172,460 segundos, saída 0. Inclui 15 testes de regras e 2 AppTests. Sem alteração de código. A suíte completa de 306 testes não foi repetida nesta retomada; o resultado anterior pertence à execução documentada no PR.

## Navegador — homologação

Destino: https://aqua-gestao-homologacao-rt.streamlit.app/

Aviso HOMOLOGAÇÃO visível; apenas clientes fictícios utilizados. Login administrativo concluído com entrada do PIN pelo formulário seguro, sem exposição do valor ao agente. O SHA implantado não foi identificado pela interface: as observações abaixo não comprovam que a implantação corresponde exatamente ao SHA local.

Verificações concluídas:

- Restauração do rascunho mais recente de Condomínio Alfa Fictício, outubro/2026: data 06/10/2026, pH 7,4, cloro livre 1,8 e uma fotografia preservada.
- Prévia confirmou uma foto proveniente dos anexos preservados do próprio contexto.
- Geração final mensal: interface confirmou registro com sucesso para Alfa, com botões de download PDF e DOCX.
- Consolidados do primeiro e quarto trimestres de 2026: interface disponibilizou PDF, informando zero visitas e zero piscinas identificadas. Não equivale a validar gráficos históricos com leituras na nuvem.
- ART do consolidado marcada Emitida sem número: geração bloqueada com mensagem explícita exigindo número.
- Restauração de Beta: nome Beta, pH 7,88 e cloro livre vazio. Retorno ao rascunho Alfa: nome Alfa, pH 7,4, cloro livre 1,8 e uma foto preservada. Verificação limitada a esses campos e à presença das fotos; não houve comparação binária das imagens no navegador.

## Pendências para aprovação final

- Confirmar SHA efetivamente implantado na homologação.
- Exercitar cadastro/persistência de tipos e volumes, upload novo com legenda, períodos trimestral/semestral/anual com visitas fictícias registradas e gráficos preenchidos, além da recuperação após reinício do servidor.
- Baixar e inspecionar visualmente os documentos gerados nesta execução Cloud; os botões e mensagens foram verificados, não o conteúdo binário dos arquivos.
- A imagem da logo ainda contém AQUA; a grafia ACQUA no título não corrige a arte.
- Sem merge e sem publicação em produção nesta retomada. Não foram apagados rascunhos, inicializadas abas de CRM/operadores, enviados e-mails ou alterados cadastros reais.


## Correções finais validadas — 31/10/2026

Código desta atualização: arte ACQUA versionada, cabeçalho DOCX com a mesma arte e títulos do consolidado unidos ao próximo elemento na paginação.

- Nova execução dos mesmos módulos isolados na versão final: **17 testes aprovados, zero falhas, zero erros, zero skips**, 216,975 segundos, saída 0.
- `python -m py_compile app.py utils/pdf_rt_consolidado.py` e `git diff --check`: aprovados.
- Trimestral, semestral e anual regenerados localmente com quatro datas de visitas fictícias e três piscinas: seis páginas cada; nome da piscina junto ao gráfico nas páginas 4, 5 e 6. Capa e página de gráfico anual renderizadas e conferidas.
- Imagem dos cabeçalhos internos DOCX conferida por SHA-256 contra a nova arte; imagens existentes no corpo preservadas.
- A suíte completa de 306 testes é evidência anterior e não foi repetida nesta alteração visual.
- Painel Streamlit confirmou homologação vinculada a `fix/rt-f123-persistencia-cloud-20261029` e produção a `main`. A branch de homologação será avançada sem force para a versão validada.
- A validação autenticada desta implantação na nuvem ainda está pendente; a nova sessão está na tela de PIN administrativo. Nenhuma alteração em produção nesta etapa.

Resultado estruturado: `rt_testes_finais_20261031.json`. Os dados usados nos testes e documentos locais são sintéticos; não há migração ou exclusão de dados reais.
