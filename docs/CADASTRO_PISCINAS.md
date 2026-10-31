# Etapa 1 — tipos e volumes das piscinas

Base: `180628089fe1ac9b76044c3f729b97359fb45471`.

O cadastro mantém Adulto, Infantil, Family e duas extras identificadas por nome.
Cada posição possui um tipo explícito; o padrão legado é Não informado.
Volumes vazios continuam sem atribuição; valores negativos, infinitos e texto
inválido bloqueiam o salvamento. Extras com volume exigem nome.

## Persistência e compatibilidade

* A:O da aba Clientes mantém seu mapeamento; J: L guarda os volumes padrão.
* P guarda JSON com schema `acqua.piscinas.v1`, tipos e extras com seus volumes.
  A coluna é expandida somente ao salvar, se necessário. Não se renomeiam abas.
* A atualização dos dados e metadados ocorre na mesma chamada de escrita da linha.
  Conteúdo desconhecido em P bloqueia a edição sem sobrescrevê-lo; os outros
  clientes continuam listados. Não há migração em lote nem consulta de produção.
* O volume total passa a incluir extras quando informadas. Um total legado sem
  distribuição por piscina é preservado quando se preenchem somente os tipos,
  sem atribuí-lo automaticamente à piscina Adulto.
* Sem metadados remotos, o cadastro usa o JSON local legado para as extras.
  Após salvar, a informação remota prevalece sobre a cópia local das piscinas.
* Chamadores antigos sem o novo argumento continuam atualizando apenas A:O.
* Carregar dados é obrigatório após selecionar outro cliente ou painel antes de
  salvar uma edição. A falha remota impede atualizar a cópia local do cadastro.
* Rascunhos, ART, fotografias RT, fórmulas de dosagem e renderizadores não mudam.

## Validação

Suítes executadas com planilha fictícia, disco temporário e rede bloqueada:

```sh
python tests/rodar_isolado.py --modules test_cadastro_piscinas test_cadastro_piscinas_ui
python tests/rodar_isolado.py --modules test_bem_star_contexto test_rt_protecao_integrada
python -m py_compile app.py utils/cadastro_piscinas.py
git diff --check
```

32 testes únicos aprovados entre as execuções: 13 de cadastro/persistência,
2 AppTests do formulário (Aqua Gestão e Bem Star), 8 de contexto Bem Star e
9 de proteção RT (incluindo modo operador e documentos).
O teste adicional do total legado foi executado após o ajuste correspondente.

Não foi feita homologação em navegador, gravação no Sheets real ou implantação.
A integração visual desses dados no relatório RT é a próxima etapa.
