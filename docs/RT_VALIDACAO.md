# Proteção integrada RT — validação para revisão independente

Base: `f128aafa178260b9089bc1b8e10a842bf52dda80` (PR #5 Draft).

## Implementado

| Defeito | Proteção | Evidência |
|---|---|---|
| D1-A | Persistência deve retornar sucesso antes de limpar/rotacionar o contexto. Falha bloqueia troca e reafirma widgets antes de rerun. | AppTest do escritório e do RT com falha de disco; última edição recuperada em nova sessão depois da recuperação do armazenamento; parecer e foto mantidos no RT. |
| D1-B | Snapshot completo dos campos RT por empresa + cliente + período, com identificação própria. | Roundtrip de diagnóstico, parecer, ART fictícia, análises/dosagens/observações/EPI; DOCX e PDF de Beta sem diagnóstico/ART/análise/foto de Alfa. |
| D1-C | Validação comum na entrada do renderizador e prévia executada depois da conciliação/importação e dos widgets. | Edição direta de nome, inspeção do XML e das imagens do DOCX e do texto/imagens do PDF da prévia. |
| D1-D | Bytes e SHA-256 das fotos serializados no rascunho próprio; materialização em pasta de identidade. | Foto restaurada após A→B→A e período→outro→período; foto recuperada encontrada nos documentos de Alfa e ausente nos de Beta. |
| D1-E | Falha de consulta/identificação impede tanto prévia quanto final antes da escrita. | Nenhum novo arquivo nem alteração de horário de prévia durante indisponibilidade simulada. |

ART começa em **Em processo de formalização**. Documento declara **ART não emitida**.
Não há número/data/vigência inventados. Cadastro futuro de ART emitida permanece disponível;
números positivos usados nos testes são exclusivamente fictícios.

## Resultado

**176 testes únicos aprovados**, 0 falhas pendentes, 0 erros e 0 ignorados.

A execução foi por módulos em processos separados. O módulo do formulário teve dois
casos reprovados durante o desenvolvimento: aviso da troca de período não reapresentado
e perda do widget no rerun de troca RT cancelada. Ambos foram corrigidos e reexecutados.
A rodada de 11 retestes passou, assim como os 2 retestes finais do seletor/falha RT.
Repetições não foram somadas à contagem de testes únicos.

* Funções, segurança e rascunhos: 141.
* Formulário real com AppTest: 26 (24 passaram na rodada do módulo, 2 nos retestes).
* Novas regressões integradas: 9.

Detalhes e hashes dos quatro arquivos Beta finais/prévias inspecionados:
[rt_validacao_resultados.json](rt_validacao_resultados.json).

Também foram gerados/abertos DOCX e PDF de Alfa, comprovando a recuperação da foto
sintética 13×17, do diagnóstico e da ART fictícia no contexto correto. A imagem é
verificada por dimensões após decodificação, inclusive quando reencodificada no PDF.
A página inicial de um PDF Beta foi renderizada e inspecionada visualmente; isso não
constitui revisão editorial completa de todas as páginas.

## Como reproduzir

Instalar `requirements.txt` e `pypdf` para inspeção dos PDFs. Executar:

```sh
python tests/rodar_isolado.py
python tests/rodar_isolado.py --modules test_rt_protecao_integrada
```

O executor bloqueia sockets e a fábrica Drive, configura armazenamento de teste e usa
os mocks/cópias temporárias da suíte. Não fornecer Secrets reais aos testes.

Ambiente validado: Python 3.12.14, Streamlit 1.61.0, Starlette 1.3.1,
python-docx 1.2.0, ReportLab 4.4.9 e pypdf 6.10.0.

## Limites e riscos restantes

* Não foi feito merge, deploy, acesso a formulários reais ou gravação em Sheets/Drive.
* AppTest não comprova apresentação/upload real em Edge/Chrome ou funcionamento no Cloud.
* Fotos antigas encontradas só por nome/pasta não são associadas automaticamente; não houve migração.
* Rascunhos com fotos usam o limite já existente de 40 partes de 40.000 caracteres.
  Se a persistência remota exceder capacidade ou falhar, a troca deve permanecer bloqueada.
  Não foi implementado armazenamento remoto de fotos fora do rascunho nesta alteração.
* A consulta continua sujeita ao cache existente; falhas são detectadas quando a consulta é refeita.
* Casos de arquivo de imagem ilegível/HEIC e outras pendências fora de D1 não foram ampliados.
* A escrita remota pelo GitHub foi recusada com **403 Resource not accessible by integration**
  pelas APIs Git Data e Contents. Não há credencial Git local configurada. Portanto,
  o PR remoto permanece em Draft no SHA de referência; o novo commit é local até existir
  uma conexão com escrita autorizada para publicá-lo.

Ver [RT_CONTEXT.md](RT_CONTEXT.md) para o mapa e a solução centralizada.
