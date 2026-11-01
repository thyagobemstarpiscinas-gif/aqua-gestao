# ACQUA GESTÃO — Finalização técnica do PR #14 (01/11/2026)

Branch `fix/rt-abnt-10818-2026`, a partir de `e0d28878536cc031ebab53967acf17f9eea21b0e`.
PR #14 permanece em rascunho. Sem merge, sem alteração de produção, sem
assinatura digital, sem migração de planilhas e sem alteração da lógica de
dosagem. Nenhum dado real foi lido ou gravado: todos os testes e documentos
usaram dados fictícios, planilhas falsas e rede bloqueada.

## 1. Critérios ABNT NBR 10818:2026 centralizados

Fonte única: `utils/normas_piscinas.py` (Tabela 3, piscina coletiva,
tratamento clorado simples, conforme `docs/RT_NORMAS_20261101.md`).

| Parâmetro | Limite aceitável | Faixa ideal |
|---|---|---|
| pH | 7,0–7,8 | 7,4–7,6 |
| CRL (mg/L) | 1,0–5,0 (0,5 só com tratamento complementar cadastrado) | 2,0–4,0 |
| Cloro combinado (mg/L) | até 0,6 | — |
| Alcalinidade (mg/L CaCO3) | 60–180 | 80–120 |
| Dureza cálcica (mg/L CaCO3) | 100–400 | 180–250 |
| CYA (mg/L) | até 100 | 30–50 |
| Turbidez (NTU) | < 0,5 | — |
| ORP (mV) | mín. 650, quando aplicável | — |

Antes desta etapa, somente CC e turbidez usavam a norma; pH, CRL, alcalinidade,
dureza e CYA ainda eram avaliados por faixas antigas (pH 7,2–7,8; CRL 0,5–3,0;
dureza 150–300; alcalinidade 80–120; CYA 30–50) e rotulados como "faixa ideal".
Consequência prática corrigida: CRL 0,8 mg/L não era sinalizado, e CRL 4,0,
pH 7,1, dureza 120 etc. eram apontados como desvio.

Geradores e telas alinhados:

- PDF mensal RT em uso (`build_monthly_rt_pdf_bytes` do `app.py`): desvios,
  linha de limites de referência, conclusão parcial, referência à NBR 10818:2026.
- `rt_monthly_report.py` (módulo; a função de PDF é sobrescrita pelo `app.py`).
- PDF premium + DOCX mensal (`avaliar_conformidade_analises`, textos
  automáticos, tabela 2.1 com limite aceitável e faixa ideal).
- DOCX legado `relatorio_mensal.docx`: tabela de faixas reescrita em tempo de
  geração; textos de referência atualizados (o arquivo modelo não foi alterado).
- Relatório consolidado (tabela "% no limite", gráficos com faixa da norma).
- PDFs de visita: operador ACQUA, supervisão RT e Bem Star; prévia HTML.
- Alertas do formulário do operador e expander "Limites NBR 10818:2026".
- Contrato (`template_rt_aqua_v2_relatorio.docx`) citava "NBR 10818:2025",
  edição inexistente; corrigido para 2026 no momento da geração.

Regras aplicadas: CT nunca é classificado isoladamente ("Avaliar CC" /
"Registrado"); CC só quando CT ≥ CRL; CT < CRL vira "Rever"/"Inconsistente" e
recomendação de repetir a medição (antes três pontos convertiam em CC = 0);
TDS e temperatura não são classificados; "CONFORME" automático continua
significando apenas ausência de leituras fora do limite e os textos dizem que
a avaliação é parcial.

## 2. Defeitos corrigidos fora da norma

- PDF mensal RT: `_to_float` removia todo ponto decimal. Valores do Sheets
  como "7.4" viravam 74 (falso desvio de pH) e CC 0,3 aparecia como "3".
- PDF mensal RT: bloco de assinatura imprimia `<br/>` literal.
- Recomendação automática de pH nunca era gerada (comparação "ph=" × "pH=").
- PDF de visita do operador: símbolos ⚠ ✓ 🏊 apareciam como quadrados (sem
  glifo na fonte incorporada).
- PDF de visita Bem Star: fonte dependia do sistema (Liberation × Vera);
  agora sempre Vera incorporada. Faixa da capa cruzava o subtítulo.
- Segurança: painel admin exibia os 2 primeiros dígitos do PIN do operador
  (metade de um PIN de 4 dígitos), também no CSV exportado; mensagens de "PIN
  em uso" ecoavam o PIN digitado. Agora a máscara é fixa (`••••`).
- `scripts/healthcheck.py` abria sempre a planilha de produção; agora usa a
  mesma regra do app (`utils.ambiente.destinos_google`) e recusa produção em
  homologação.
- `tests/test_normas_10818_2026.py` estava em estilo pytest e não era executado
  por `tests/rodar_isolado.py` (unittest); agora roda.

## Testes executados

`python tests/rodar_isolado.py` (rede bloqueada, planilhas falsas, dados fictícios):

- Base `e0d2887`: 318 testes, 317 aprovados, 1 falha (fonte do PDF de visita
  Bem Star dependente do sistema — corrigida nesta etapa).
- Branch (`4d1a6bf` + ajuste de teste): 334 testes, 333 aprovados e 1 falha
  esperada (`test_consolidacao_por_piscina_e_periodo` assumia a faixa antiga
  pH 7,2–7,8). O teste foi atualizado para os limites 2026 e passa
  (`test_rt_relatorio`: 15/15). A suíte integral não foi repetida após esse ajuste.

## 3. Acessos (RT, administrador, operadores)

Revisado sem alteração de regra: PINs especiais só via Secrets (`[rt] pin`,
`[admin] pin`, `[operador] pin_mestre`); comparação em tempo constante; PIN vazio
nunca autentica; operador inativo negado; lista vazia não dá acesso total;
PINs reservados não podem ser atribuídos a operadores; campos de PIN do tipo
senha. Recuperação administrativa: o administrador redefine o PIN do operador
(o PIN atual não é exibido). Não há recuperação de PIN de RT/admin pelo app —
é feita no painel de Secrets, por decisão de segurança.

## 4. Homologação Streamlit

`utils/ambiente.py` mantém: sem `[aqua_app]` → produção; `ambiente =
"homologacao"` exige IDs próprios de planilha e pasta, recusa IDs de produção e
interrompe a inicialização. O healthcheck passou a seguir a mesma regra. Os IDs
de produção continuam no código (já eram públicos) apenas como referência de
bloqueio. Não foi feito acesso ao app de homologação nem à produção nesta
etapa; PINs/credenciais não foram copiados.

## 5. Documentos fictícios para inspeção

Gerados pelos geradores reais com o script
`tests/gerar_documentos_ficticios.py` (sem rede):

```
python tests/gerar_documentos_ficticios.py generated/inspecao_ficticia
```

Saída: visita Bem Star, visita operador ACQUA, supervisão RT, mensal RT,
mensal RT premium + DOCX, consolidado trimestral, mensal Bem Star, DOCX legado.
Conferidos visualmente: logos corretos por marca, capas, tabelas, gráficos,
foto fictícia, fontes incorporadas, ausência de quadrados/`<br/>`.

## 6. Pendências e bloqueios

- O texto integral das normas não está no repositório (direitos autorais); as
  faixas seguem o resumo de `RT_NORMAS_20261101.md`. Conferência final pelo RT.
- Tratamento complementar (CRL mín. 0,5), TDS basal e rotina de 3 medições/dia
  ainda não têm cadastro; o sistema não os infere.
- Dosagem: as sugestões automáticas continuam com as metas internas anteriores
  (não alteradas por regra do projeto). A mensagem "dentro da faixa ideal" da
  sugestão de dosagem refere-se a essas metas internas.
- Operador ACQUA PDF não calcula CC quando o lançamento não traz "cloraminas".
- Homologação no Streamlit Cloud com o SHA final ainda precisa ser executada
  pelo responsável (roteiro em `HOMOLOGACAO_RT.md`).
- `pypdf` é usado por um teste e não consta em `requirements.txt` (dependência
  só de teste).

## 7. Identidade visual ACQUA GESTÃO (correção posterior)

Causa: `build_monthly_rt_pdf_bytes` (app.py) tinha lista própria de logos
iniciando por `aqua_gestao_logo.png` (arte antiga "AQUA GESTÃO") e nunca
consultava `acqua_gestao_logo_rt.png`. Havia ainda artes antigas embutidas em
base64 (`LOGO_AQUA_OFICIAL_B64`, `LOGO_AQUA_APROVADA_B64`,
`LOGO_DOSSIE_JPG_B64`, logo do dossiê FDS/GHS) usadas como alternativa, e todos
os modelos DOCX ACQUA (`relatorio_mensal.docx`, `aditivo.docx`,
`template_rt_aqua*.docx`) traziam a arte antiga embutida.

Correção:
- `utils/marca.py`: arte aprovada única (`acqua_gestao_logo_rt.png`), sem
  alternativa antiga; sem a arte aprovada o documento sai com marca textual.
- `encontrar_logo()` e `garantir_logo_aqua_oficial()` retornam só a arte
  aprovada; varredura genérica de arquivos "logo" removida.
- Base64 antigos removidos do app.py; dossiê FDS/GHS usa o arquivo aprovado.
- DOCX: na geração, a arte antiga (identificada por SHA-256) é substituída
  pela aprovada, sem distorção, em corpo e cabeçalhos; relações órfãs
  removidas. Modelos Bem Star não são alterados. Os arquivos-modelo e as
  artes antigas permanecem no repositório (histórico), sem uso.
- Nome-marca em cabeçalhos/títulos: "ACQUA GESTÃO" (PDFs, prévias HTML,
  propostas/contratos CRM, DOCX). A razão social registrada "AQUA GESTÃO
  CONTROLE TÉCNICO LTDA", ao lado do CNPJ, não foi alterada.
- Prévias HTML ACQUA passam a exibir a logo aprovada.
- Alerta da prévia HTML de visita passa a avaliar cada piscina (antes dizia
  "dentro dos limites" com CRL fora do limite em uma piscina).

Verificação: `tests/test_marca_acqua.py` gera 28 documentos (23 ACQUA e
5 Bem Star) e classifica as imagens incorporadas por impressão digital
(`utils/verificar_marca.py`): ACQUA só com a arte aprovada, Bem Star só com a
própria marca, nenhum com a arte antiga ou com o nome-marca antigo.

Pendências de marca:
- Razão social "AQUA GESTÃO CONTROLE TÉCNICO LTDA" e "nome fantasia RT AQUA
  GESTÃO" no texto do contrato: decisão jurídica/cadastral do responsável.
- Textos em caixa normal "Aqua Gestão" (rodapés, e-mails, telas) não foram
  alterados nesta etapa.
- Não existe gerador de "relatório de vistoria inicial" no código atual.
- `logo_url` das assinaturas de e-mail vem dos Secrets e não foi verificado.
- No DOCX mensal, a página 2 mostra a logo no cabeçalho e no corpo (leiaute
  herdado do modelo; agora ambas com a arte aprovada).

Suíte integral após a correção de marca (`python tests/rodar_isolado.py`,
rede bloqueada, dados fictícios): 342 testes, 342 aprovados, 0 falhas.
