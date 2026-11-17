# Pontos de atenção no relatório mensal RT

O mensal do escritório identificava apenas violações de limites aceitáveis. Leituras
fora das faixas ideais e CC no limite aceitável podiam resultar em parecer sem ressalvas.
Além disso, apenas as duas primeiras ocorrências de ORP entravam nas observações.

- Os limites existentes da NBR 10818:2026 permanecem inalterados.
- Faixa ideal e meta interna não se tornam limites de não conformidade.
- Meta interna de CC: menor que 0,50 mg/L. De 0,50 até 0,60: atenção; acima de 0,60:
  não conformidade. Exatamente 0,60 recebe o texto “no limite aceitável”.
- ORP abaixo de 650 mV continua como atenção, com a ressalva “quando aplicável”.
- Leituras ausentes, CT inferior ao CL, TDS informativo e prioridade da não
  conformidade conservam suas regras.
- Alertas agrupam todas as datas, piscinas e valores. O formulário conserva cinco
  campos e reúne o excedente no último, sem descartar ocorrências/recomendações.
- O status interno CONFORME permanece compatível com rascunhos; a apresentação na
  tela, PDF e DOCX inclui a ressalva de atenção quando houver alertas.
- Tabela e gráficos mensais distinguem atenção em âmbar e limite excedido em vermelho.
  Os gráficos distinguem faixa aceitável, ideal e meta interna de CC.
- Ao gerar um rascunho antigo, os textos editados pelo RT são preservados e os
  alertas atuais são acrescentados aos dados do documento, sem regravar o rascunho.

Validação: 46 testes de regras/documentos e três cenários AppTest (sem rede),
incluindo limites exatos, várias piscinas, relatório sem dados, PDF e DOCX reais,
rascunho com texto manual e todas as ocorrências do caso de referência. Inspeção
visual das tabelas, gráficos, interpretação e parecer do PDF de teste.

Nenhuma medição de produção é alterada; documentos já emitidos exigem nova geração.
