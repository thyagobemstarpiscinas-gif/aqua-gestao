# Piscinas do cadastro no modo operador — 01/11/2026

Defeito: uma piscina adulto já cadastrada remotamente era exibida como não configurada no formulário do operador.

Causa: a tela dependia de uma lista em JSON local e interpretava exatamente [Piscina Adulto] como configuração ausente. O cadastro remoto precisa alimentar a tela também quando o arquivo local não existe após reinício do Cloud.

Correção: resolver_piscinas_operador usa as piscinas identificadas no cadastro remoto (volumes individuais, tipos e extras). A lista local explícita permanece como compatibilidade. Uma única piscina adulto é configuração válida. Apenas volume total sem distribuição não identifica uma piscina.

Não grava cadastro nem visita ao carregar. Não altera cálculos ou recomendações de dosagem, permissões, PINs, rascunhos ou assinatura.

Validação: 72 testes de regras/cadastro/formulário/segurança aprovados, incluindo seis regressões do resolvedor; dois AppTests com planilha fictícia aprovados, cobrindo Aqua Gestão e Bem Star sem pasta local. Sintaxe e git diff --check aprovados. Todos os valores e credenciais dos testes são fictícios.

A verificação na produção identificou o perfil autenticado, a restrição ao cliente autorizado e a disponibilidade do histórico, sem acesso aos módulos exclusivos de RT. Nenhum nome, identificador ou histórico de cliente/operador de produção é reproduzido nesta nota. Não foram criadas visitas ou medições de teste em produção.

A consulta integral da revisão NBR 10818 continua aguardando a conclusão do cadastro do titular no portal ABNT–CFQ. Esta correção não modifica limites numéricos.
