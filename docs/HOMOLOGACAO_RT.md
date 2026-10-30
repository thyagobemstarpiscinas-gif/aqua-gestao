# Homologação final RT — roteiro e resultados

## Situação atual — 30/10/2026

Ambiente separado criado e publicado com autorização do responsável: https://aqua-gestao-homologacao-rt.streamlit.app/ . PR #6 permanece em rascunho, dependente do #5, sem merge e sem publicação em produção.

Referência da branch consultada na rodada de verificação: `3e3c2aa5b6841f3f2e40b8e8d92da9ea221aa284`. O SHA efetivamente executado pelo Cloud não foi confirmado por logs ou metadados de deploy; a referência da branch não comprova o commit instalado.

### Evidências da rodada anterior, conforme registro da sessão

- Leitura da planilha de teste e recuperação dos valores de Alfa pelos dois botões, com origem 09/2026 e datas de outubro preservadas.
- Recuperação em nova sessão após despertar do Cloud. Isso não comprova reinício do processo nem descarte físico do disco.
- Beta não herdou os dados de Alfa; seus próprios marcadores foram preservados ao negar recuperação sem histórico.
- Rascunho inválido recusado com “Total de análises inválido”, preservando pH, produto, diagnóstico e foto de Beta após rerun.
- Rascunho válido restaurado com a foto de Beta; ao restaurar Alfa, a foto de Beta não apareceu.
- Nomes colidentes de Residencial Sol bloquearam importação e geração por ambiguidade.
- PDF final de Beta com quatro páginas, marcador e foto próprios, sem marcadores de Alfa. Prévias não alteraram os registros finais na planilha.

### Verificação realizada na continuação

- Login administrativo com empresa Bem Star confirmado pela tela: “Empresa logada: Bem Star Piscinas”.
- Carteira e cadastro mostram um cliente: Condomínio Gama Fictício.
- Seletor do relatório Bem Star aberto: uma única opção, Condomínio Gama Fictício; nenhum cliente Aqua Gestão listado.
- Campos de operador, observações gerais e parecer técnico-operacional inicialmente vazios, sem marcadores RT.
- Nenhum relatório Bem Star foi gerado nesta verificação; nenhum cadastro foi alterado.

### Limites e pendências

- Separação do painel Bem Star confirmada no cenário acima; não equivale a auditoria de todos os caminhos de autorização.
- Na rodada anterior, a imagem da prévia ficou indisponível no Cloud por dependência ausente; geração DOCX/PDF funcionou, conforme registro da sessão. Diagnóstico e correção dessa dependência ainda precisam de verificação técnica.
- Falta comprovar o SHA instalado e repetir a recuperação após reinício controlado do processo de homologação.
- Pendências locais de rascunhos e concorrência entre escritores continuam com as limitações descritas em RT_CLOUD_VALIDACAO.md.
- Merge e publicação em produção continuam fora da autorização desta etapa.

---

## Roteiro original de preparação (histórico)

As instruções abaixo registram o planejamento anterior. Afirmações de ambiente ainda não criado ou deploy não autorizado referem-se à preparação, superada pela situação atual acima.

## Situação

PR #6 em rascunho, dependente do PR #5. As correções estão enviadas para uma branch separada. Nenhum merge ou deploy foi autorizado nesta etapa. Os testes locais não substituem a verificação externa abaixo.

## Aplicativo separado

Criar um aplicativo de homologação no Streamlit Cloud, usando:
- Repositório: thyagobemstarpiscinas-gif/aqua-gestao
- Branch: fix/rt-f123-persistencia-cloud-20261029
- Entrada: app.py
- Python: 3.12, como a validação local anterior
- Acesso privado/restrito durante a verificação

Não alterar a branch, os secrets ou a configuração do aplicativo de produção.

## Dados e credenciais

Usar planilha nova e pasta Drive nova, ambas exclusivamente de teste; não copiar dados reais de clientes nem usar a cópia antiga encontrada no Drive como se fosse um conjunto fictício. Usar conta de serviço dedicada, com acesso somente aos dois recursos de teste. A conta conectada ao ChatGPT não substitui a conta de serviço que o aplicativo utiliza.

Compartilhar esses recursos com o e-mail da conta de serviço de teste. Configurar a chave apenas no painel privado de secrets do aplicativo de homologação; não enviar chaves neste chat nem incluir no repositório.

Adicionar aos secrets do aplicativo de teste:

```toml
[aqua_app]
ambiente = "homologacao"
sheets_id = "SUBSTITUIR_PELO_ID_DA_PLANILHA_NOVA"
drive_fotos_folder_id = "SUBSTITUIR_PELO_ID_DA_PASTA_NOVA"
```

Manter as credenciais de teste na seção `gcp_service_account`, conforme o formato já utilizado pelo app. As configurações de autenticação devem usar usuários/PINs de teste; não copiar secrets de produção integralmente. Não configurar CODESPACES=true no aplicativo de homologação, pois isso seleciona persistência local para os últimos valores e rascunhos.

Sem a seção aqua_app, o app mantém os destinos existentes de produção por compatibilidade. Por isso, o aplicativo de homologação NÃO deve ser iniciado antes de configurar aqua_app. Quando modo homologacao está explícito, IDs ausentes, inválidos ou iguais aos destinos de produção interrompem a inicialização; não há fallback para produção. A tela mostra aviso de homologação.

Não existe prova por formato de ID de que uma planilha seja fictícia: é necessário conferir os recursos selecionados e restringir as permissões da conta de serviço. Nenhuma planilha real foi criada, modificada ou compartilhada nesta preparação.

## Roteiro curto com dados fictícios

1. Cadastrar dois clientes fictícios em Aqua Gestão e um em Bem Star; usar IDs estáveis e nomes colidentes para conferir isolamento.
2. Gerar relatório final de Alfa com marcadores de teste. Conferir os arquivos e a criação da aba `_RT_Ultimos_Valores` na planilha de teste.
3. Abrir outra sessão, selecionar outro mês e recuperar pelos dois botões. Conferir valores, mês de origem e datas atuais preservadas.
4. Selecionar Beta/Bem Star. Confirmar que não recebem valores, fotos ou diagnósticos de Alfa.
5. Conferir que prévias não acrescentam registro final. Reiniciar somente a instância de teste e repetir a recuperação para conferir independência do disco.
6. Conferir restauração válida de rascunho e recusa de rascunho inválido sem apagar campos. Se houver pendência, conferir o aviso de que ainda não está salva na planilha.

Encerrar após esse roteiro; não repetir a suíte longa sem uma falha ou mudança que justifique. Registrar resultados e qualquer impedimento real. Se algo falhar, corrigir e retestar apenas a parte afetada antes de decidir a publicação.

## Critério para liberar

A etapa simulada está documentada em RT_CLOUD_VALIDACAO.md. A homologação externa só estará concluída após os passos acima, com o commit efetivamente instalado registrado. Merge e publicação em produção exigem decisão separada do responsável. O PR #6 tem base na branch do PR #5; não deve ser tratado como PR independente contra main.
