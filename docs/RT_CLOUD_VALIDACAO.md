# Preparação para persistência no Cloud — 29/10/2026

Esta versão acrescenta persistência dos últimos valores via Google Sheets. Não foi publicada.

## Alterações sobre o ZIP Windows
- app.py: os dois pontos de leitura/publicação passam pelo adaptador remoto.
- utils/rt_contexto.py: extração do construtor validado do registro; API local preservada.
- utils/ultimos_sheets.py: novo adaptador sem conexão na importação.
- tests/test_ultimos_sheets.py: 12 testes novos com Sheets em memória e rede bloqueada.
- tests/test_rt_ultimos_regressoes.py: fixture explicita modo local de desenvolvimento e conector fictício; expectativas mantidas.

## Comportamento
Em produção, usa a conexão conectar_sheets existente. Acrescenta a aba _RT_Ultimos_Valores se ela estiver comprovadamente ausente, somente ao publicar um relatório final concluído. Não renomeia nem altera outras abas. Colunas: Identidade, Gerado em, Dados JSON, SHA256. Um registro por geração final; nunca por autosave ou prévia. Registros anteriores não são sobrescritos. Não há limpeza automática. Consultar todo o histórico pode custar mais com o crescimento da aba.

A identidade depende de empresa + ID estável e independe do nome/período. O último registro é escolhido pelo instante UTC da geração; empate usa a posição na planilha. Datas antigas das análises não são transferidas. Leituras em produção não usam fallback para disco; falha, ausência ou corrupção produz aviso e preserva campos atuais. Não migra arquivos legados/local automaticamente. JSON acima de 45.000 caracteres é recusado, preservando o registro anterior. Conteúdo usa RAW e hash verificado; hash detecta alteração, não é assinatura criptográfica.

Após append, uma leitura confirma o registro completo. Se a resposta se perder, a escrita pode ter ocorrido: o aviso não afirma sucesso; uma consulta posterior pode encontrar o registro. Sem retentativa automática de publicação. O instante vem do servidor da aplicação; não há transação distribuída nem garantia de ordenação global entre servidores com relógios divergentes.

CODESPACES=true mantém o modo local já adotado pelo projeto e não abre a conexão. Outros ambientes usam Sheets. Nenhuma planilha real foi acessada nesta preparação.

## Rascunhos
A sincronização já existente dos rascunhos foi preservada. O aviso atual já informa que edições NÃO sincronizadas dependem da conservação do disco. Não há garantia de recuperar essas edições após perda do disco do Cloud. Não foi acrescentada persistência remota para uma edição cujo envio falhou.

## Validação desta versão
24 testes direcionados passaram: 12 novos + 12 regressões. Sintaxe dos cinco arquivos alterados/adicionados aprovada. Fixtures fictícias, armazenamento temporário e rede bloqueada. Não foi usada credencial real.

Os 213 testes em 25 processos do pacote anterior são evidência histórica anterior a esta mudança. Não foram repetidos; não se deve apresentar esse resultado como aprovação integral desta versão. Streamlit e gspread não estão instalados neste ambiente; AppTest, navegador, credenciais/permissões reais e Streamlit Cloud não foram verificados.

## Integração e próxima etapa
Preservar no GitHub modelos DOCX, assinaturas e demais arquivos omitidos do backup. Não substituir todo o repositório por este ZIP. O patch incluído representa somente esta etapa sobre o ZIP Windows, não o delta completo contra o PR #5. Na preparação original do ZIP não houve push, commit, merge ou deploy. Posteriormente o código foi enviado ao PR #6, que permanece em rascunho, sem merge ou deploy.

Antes da publicação, revisar a integração com a aba nova e verificar o fluxo remoto em homologação com uma planilha de teste. Não usar produção para essa verificação sem autorização.

Documentação consultada: https://docs.gspread.org/en/latest/api/models/worksheet.html ; https://docs.streamlit.io/develop/concepts/connections/connecting-to-data .


## Revisão e fluxo simulado adicional — 29/10/2026

Base revisada: PR #6, commit a3536a93df0b7e0a737aace472a936d1d7a0aef5. A revisão verificou o adaptador Sheets, os pontos de chamada do renderizador e recuperadores e a preservação dos demais arquivos na integração. Não havia comentários ou revisões externas no PR na consulta realizada. Nenhuma alteração na aplicação foi necessária nesta etapa.

Foram acrescentados quatro testes em tests/test_rt_cloud_fluxo.py. Todos passaram, sem erros ou ignorados:
- Renderizador final → publicação na planilha em memória → perda integral do disco → nova sessão em outro período → recuperação pelos dois aplicadores reais, com datas preservadas.
- Prévia, falha de DOCX/PDF/manifesto e contexto inválido não alteram os registros remotos anteriores.
- Falha de Sheets mantém os documentos simulados já gerados e exibe aviso de últimos valores não atualizados.
- Recuperação após geração final recusa outra empresa ou outro ID e preserva os campos ativos.

As funções reais do app foram extraídas por AST. Documentos, validação do contexto e serviços foram simulados; os testes não são AppTest nem inspecionam DOCX/PDF reais. A rede foi bloqueada pelos fixtures e o disco era temporário. Comando: python -m unittest discover -s tests -p test_rt_cloud_fluxo.py -v. Resultado: 4/4, zero falhas, erros ou ignorados. Sintaxe do novo teste aprovada.

Esses resultados complementam os 24 testes direcionados anteriores; a suíte de 213 testes não foi repetida. Nenhum serviço ou credencial de produção foi usado, nenhuma aba real foi criada e não houve merge/deploy.

Próxima verificação externa, ainda não executada: ambiente de homologação com planilha separada de dados fictícios e credenciais próprias. Conferir criação/permissões da aba nova, gerar um relatório final, recuperar em outra sessão/mês e verificar o resultado após reiniciar a instância de teste. Não reutilizar a planilha de produção. A publicação desse ambiente requer autorização separada de deploy.


## Preparação de homologação isolada — 29/10/2026

A configuração de destinos Google foi extraída para utils/ambiente.py. A seção opcional aqua_app nos secrets permite declarar homologacao e dois IDs exclusivos de teste. No modo explícito de homologação, IDs ausentes, inválidos ou iguais aos de produção interrompem o app. Sem a seção, os destinos existentes permanecem iguais. Foi acrescentado aviso visual no modo de homologação. Não houve alteração de autenticação, dosagem, nomes de abas ou regras de empresas.

Oito testes novos em tests/test_ambiente.py passaram, sem falhas, erros ou ignorados, incluindo execução do bloco real de interrupção por configuração insegura. Sintaxe de app.py, helper e testes verificada. O fluxo remoto simulado de quatro testes foi repetido após essa mudança e passou. Os testes de configuração usam mapeamentos fictícios; não iniciam Streamlit nem abrem conexão. Nenhuma credencial real foi usada e nenhum conteúdo da planilha de produção foi lido.

Foi feita apenas busca de metadados no Drive: identificadas a planilha de produção e uma cópia antiga, sem evidência de que a cópia contenha apenas dados fictícios. Nenhuma foi selecionada para homologação. A criação/publicação do ambiente externo permanece pendente. Roteiro completo: HOMOLOGACAO_RT.md.
