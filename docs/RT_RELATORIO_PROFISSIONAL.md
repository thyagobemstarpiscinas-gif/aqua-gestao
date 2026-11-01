# Relatório RT ACQUA GESTÃO — capa, piscinas, fotos, gráficos, ART e consolidados

Branch: `feat/acqua-rt-relatorio-profissional-20261031`, criada sobre o PR #10
(`feat/cadastro-tipos-piscinas`, `fefc3ef`), que por sua vez parte de `main` (`1806280`).
Este PR **inclui o commit do PR #10**; revisar/integrar um no lugar do outro, não ambos.

Escopo: somente o relatório RT da ACQUA GESTÃO (escritório e, onde indicado, área RT do
operador). O relatório mensal Bem Star (`utils/*bem_star*`) não foi alterado.

## Situação encontrada (31/10/2026)

* PRs #1–#9 integrados em `main`. Os registros de `HOMOLOGACAO_RT.md` e
  `RT_CLOUD_VALIDACAO.md` que citam #5/#6 "em rascunho" são históricos (ambos foram
  integrados em 30/10/2026).
* PR #10 aberto em rascunho; a apresentação das piscinas no RT estava pendente.
* Branch `feat/acqua-rt-profissional-20261031` existia sem commits próprios (igual a `main`).
* Fotos do RT já eram preservadas no rascunho (bytes + SHA-256 por contexto); faltavam
  legenda, identificação da piscina e remoção explícita.
* PDF RT: sem capa própria, logo pequena, cabeçalhos de tabela invisíveis (texto azul sobre
  fundo azul) e gráfico apenas no DOCX, sem separação por piscina.

## Revisão e correções do PR #10

* `value_input_option` voltou a `USER_ENTERED` na inclusão/edição da linha de Clientes.
  O PR passara para `RAW`, o que gravaria volumes (D, J–L) e a data (H) como texto em
  linhas novas/editadas, diferente das linhas antigas. A coluna P (JSON começando com `{`)
  continua texto com `USER_ENTERED`.
* Cliente com nome já existente: mensagem específica ("Já existe um cliente com esse nome")
  em vez do aviso genérico de falha remota.
* Demais pontos do PR (schema `acqua.piscinas.v1`, bloqueio de conteúdo desconhecido na
  coluna P, total legado preservado, carregamento obrigatório antes de editar) mantidos.

## Implementado

| Item | Onde | Comportamento |
|---|---|---|
| Piscinas no RT | `_piscinas_relatorio_rt`, `utils/rt_relatorio.piscinas_do_cliente` | Lidas do cadastro pelo **ID estável** do contexto (`aqua_gestao\|id:Cn\|mm/aaaa`). Identificação, tipo e volume de cada piscina; tipo "Não informado" em cadastros antigos; volume total sem distribuição é informado à parte e nunca atribuído a uma piscina. Relatório avulso (sem ID): nenhuma inferência por nome. |
| Análises por piscina | coluna **Piscina** (`rel_analise_piscina_N`) | Seletor com as piscinas do cadastro (texto livre no avulso). A importação das visitas cria **uma linha por piscina** de cada visita (antes só a primeira piscina era usada). Linhas com dados e linhas extras pedidas não são mais escondidas pela regra "1 por semana". |
| Capa | `utils/capa_acqua.py` (PDF), `_docx_inserir_capa_rt` (DOCX) | Página inteira azul-noite/dourado, logo com ~92 mm, título **ACQUA GESTÃO**, cliente, período, piscinas, RT, situação da ART e data. No DOCX a capa não recebe o cabeçalho/rodapé do modelo. Rodapé do PDF e campo "Empresa" do RT: "ACQUA GESTÃO – Controle Técnico de Piscinas". |
| Fotos | `utils/rt_contexto.py` | Legenda (até 200 caracteres) e piscina por foto, gravadas **no mesmo JSON `rel_fotos_preservadas`** do contexto (identidade + SHA-256). Reenvio da mesma imagem não apaga a legenda. Remoção explícita por botão. PDF/DOCX: "Foto N — Piscina — Legenda"; sem legenda, mantém o texto padrão anterior. |
| Gráficos | `utils/rt_relatorio.gerar_grafico` (PIL, sem nova dependência) | Um gráfico por piscina no PDF e no DOCX (pH, CRL, CT, alcalinidade, dureza, CYA — só os parâmetros com leitura), com faixa de referência e pontos fora da faixa em vermelho. Com "antes / após" usa a leitura inicial. |
| ART | `utils/rt_relatorio` + seletores | Situações: **Não emitida**, **Em processo de emissão** (padrão), **Emitida**. Valores antigos ("Em processo de formalização", "Em tramitação") são lidos como "Em processo de emissão"; nenhum vira "Emitida". Emitida exige número; vigência é opcional, mas se informada exige início e fim válidos e fim ≥ início. O documento mostra número/vigência só quando informados ("Vigência não informada" quando ausente). A área RT do operador usa as mesmas situações e a mesma validação. |
| Consolidado | `gerar_relatorio_consolidado_rt`, `utils/pdf_rt_consolidado.py` | Trimestral (1–4), semestral (1–2) ou anual, a partir das visitas **identificadas** do cliente carregado. Por piscina: leituras, mínimo, média, máximo, % na faixa e médias mensais; gráficos históricos; capa própria. ART do consolidado é escolhida no próprio painel (não herda a ART mensal). Somente leitura: não altera rascunhos, análises mensais nem cadastro. PDF em `generated/_rt_consolidados/<sha256(empresa\|id\|tipo\|índice\|ano)>/`. |
| Correção visual | PDF RT | Cabeçalhos das tabelas agora em branco sobre azul; título da piscina não se separa da sua tabela. Parágrafos vazios no fim do modelo DOCX não geram mais página em branco. |

## Proteção dos dados

* Nenhuma migração, exclusão ou escrita em dados reais. Nenhuma alteração de abas.
* Isolamento: piscinas pelo ID estável; fotos/legendas no JSON com identidade validada
  (`ler_fotos` recusa outra identidade); consolidado exige cliente com ID e consulta com
  status `ok` (falha, bloqueio ou ambiguidade suspendem a geração).
* Rascunhos: legenda/piscina das fotos e piscina das análises são campos `rel_` do próprio
  contexto e seguem a mesma persistência transacional (falha de gravação continua
  bloqueando a troca). Widgets de legenda são recarregados do JSON após restauração ou
  troca de contexto, evitando sobrescrever o valor restaurado.
* Validação comum: `validar_relatorio_mensal` (usada por prévia e final) passou a usar
  `validar_art`.

## Testes

Ambiente: Python 3.13.18, Streamlit 1.61.0, ReportLab 5.0.1, python-docx, pypdf, PyMuPDF.
Rede bloqueada, planilha falsa em memória (`tests/e2e/fakes`), disco temporário.

```sh
python tests/rodar_isolado.py --modules test_rt_relatorio test_rt_relatorio_apptest
python tests/rodar_isolado.py
```

* `tests/test_rt_relatorio.py` (15): ART (legado, texto, validação), piscinas (novo, legado,
  total sem distribuição, JSON corrompido), leitura/agrupamento, uma linha por piscina,
  períodos, consolidação, gráficos, legendas no rascunho/reenvio/remoção/isolamento e PDF
  consolidado.
* `tests/test_rt_relatorio_apptest.py` (2, app real): importação por piscina, foto com
  legenda e piscina, ART emitida sem vigência, geração final (PDF/DOCX com capa, piscinas,
  volumes, legenda, gráficos por piscina), troca Alfa→Beta→Alfa e restauração em nova
  sessão preservando foto/legenda, consolidado do 3º trimestre e bloqueio de ART emitida
  sem número.
* Testes existentes ajustados à nova nomenclatura da ART (2 linhas): rótulo do seletor do
  operador e o texto "formalização" → "processo de emissão".
* Resultado da suíte completa: ver descrição do PR.

Conferência visual: PDF mensal (6 páginas), DOCX convertido por LibreOffice (8 páginas) e
PDF consolidado (3 páginas) gerados pelo teste AppTest foram renderizados e inspecionados.

## Limitações

* A arte corrigida `acqua_gestao_logo_rt.png` tem prioridade nos geradores RT PDF/DOCX e consolidado. A imagem foi corrigida para **ACQUA GESTÃO**, preservando o desenho existente; a imagem incorporada anterior permanece apenas como fallback em instalações incompletas.
* `EMPRESA_RT` (contratos, aditivos, e-mails) não foi alterado; a nova grafia vale para os
  documentos do relatório RT.
* O consolidado usa as visitas registradas (aba de visitas), não as edições manuais feitas
  nos rascunhos mensais nem relatórios avulsos. A ART do consolidado não é persistida.
* O gerador RT da área do operador (`build_monthly_rt_pdf_bytes`) recebeu somente a nova
  situação da ART; capa, piscinas, legendas e gráficos estão no gerador do escritório.
* Fotos continuam limitadas ao tamanho do rascunho (40 partes × 40.000 caracteres).
* O estado inicial descrito acima é histórico. A retomada e os limites atuais de homologação estão em `RT_TESTES_RETOMADA_20261031.md`.

## Correções finais da retomada (31/10/2026)

* A arte versionada `acqua_gestao_logo_rt.png` tem prioridade na capa mensal PDF/DOCX e nos consolidados. Os cabeçalhos internos do modelo DOCX também recebem a arte corrigida; imagens no corpo são preservadas.
* Títulos do consolidado usam `keepWithNext`, mantendo o nome da piscina junto ao gráfico na quebra de página.
* Consolidados regenerados localmente com dados sintéticos: trimestral (4º), semestral (2º) e anual de 2026; quatro datas de visita e três piscinas, seis páginas por PDF. Os gráficos das três piscinas ficaram nas páginas 4, 5 e 6, cada qual com seu nome.
* Verificação independente do cabeçalho DOCX: bytes da imagem relacionados ao cabeçalho correspondem ao SHA-256 da arte corrigida; quantidade de imagens preexistentes no corpo preservada.
* Validação isolada de regras, geração e restauração: ver resultado final da retomada. A suíte completa anterior de 306 testes não foi repetida para esta correção visual.
