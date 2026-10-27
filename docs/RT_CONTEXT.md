# Proteção integrada RT — mapa anterior à implementação

Referência: f128aafa178260b9089bc1b8e10a842bf52dda80, PR #5 Draft.

* Identidade: empresa + ID do cadastro (ou nome avulso validado) + mês/ano.
* Campos: todos os `rel_` escalares são conteúdo do contexto, incluindo
  identificação, ART, diagnóstico/parecer, análises, dosagens, observações,
  recomendações, NBR/EPI, legendas e contadores. Uploads e seletores de importação
  são transitórios; metadados de origem e importação pertencem ao contexto.
* Entrada: Carregar cliente, edição direta de nome/período, cadastro e restauração.
* Persistência: `_relatorio_rt_salvar_rascunho` escreve JSON local e, no Cloud,
  sincroniza `_Rascunhos_RT`. Retorno False não comprova persistência remota.
  A rotação da chave de sessão só pode ocorrer após sucesso.
* D1 atual guarda apenas `_valores_importaveis_rt`; diagnóstico, ART e demais
  campos continuam globais. Fotos são objetos UploadedFile transitórios.
* Prévia: execução automática antes da conciliação, via
  `gerar_previa_exata_relatorio` → `renderizar_relatorio_oficial` →
  `_renderizar_relatorio_rt`. Final chega ao mesmo renderizador com preview=False.
* Saídas: DOCX por `preencher_relatorio_mensal_docx`; PDF por ReportLab,
  conversão DOCX como fallback. A validação comum deve preceder escrita de arquivos.
* Fotos: resolvedor atual busca upload, Drive e pasta por nome sem vínculo
  comprovado de identidade; será substituído no RT por anexos vinculados ao contexto.
* Segurança: não migrar dados reais, não consultar Sheets/Drive nos testes.
  Testes de falha, roundtrip e conteúdo DOCX/PDF devem usar simulações/temporários.

ART: padrão em processo de formalização, sem emissão/número/vigência presumidos.
Cadastro futuro explícito de ART emitida continua suportado.

## Solução integrada

* `utils/rt_contexto.py` define a seleção dos campos, limpeza, captura serializada
  e verificação de hash das fotos, e materialização em pasta de identidade.
* `guardar_tabela_rt` guarda todos os campos RT e só limpa/rotaciona a sessão
  após retorno positivo da persistência; a falha restaura o contexto de origem.
* `preparar_contexto_rt` é comum ao escritório e à área RT do operador. Reafirma
  o estado antes de reruns, inclusive na troca cancelada por falha de gravação.
* `validar_contexto_rt` é a porta de entrada do renderizador RT, para preview e
  final. O modo RT do operador utiliza a mesma validação de identidade.
* A execução da prévia acontece após importação, conciliação e widgets RT.
* Campos de cadastro só podem iniciar outro período do mesmo ID/empresa; ART,
  textos técnicos, medições e fotografias são próprios de cada período.
* Fotos são bytes + nome + SHA-256 no campo JSON `rel_fotos_preservadas` do próprio
  rascunho. A galeria/relatório pode recuperar essas fotos com o uploader vazio.
* Arquivos RT ficam em `generated/_rt_contextos/<hash da identidade>/`, evitando
  sobreposição por nome entre empresas, clientes e períodos.
* O rascunho RT não coleta/restaura campos do relatório operacional `csr_` nem
  chaves de autenticação. JSON importado não pode alterar autenticação/empresa.
* A barra antiga por pasta/nome foi substituída pela persistência comum. Não há
  restauração automática de arquivo legado sem identidade nem busca genérica de
  fotos locais/Drive por nome para completar o relatório RT.

Nenhum dado antigo é migrado ou excluído. Uma futura associação de fotos legadas
sem origem comprovada depende de fluxo explícito; não é feita por inferência.
