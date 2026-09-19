# 🧭 Decisões técnicas e compromissos

## 1. Regras operacionais no SQL Server

As [procedures de operação](../database/migracoes/003_Operacoes.sql) concentram validação e persistência. Isso permite confirmar ou desfazer conjuntamente venda, pagamentos e movimentos, enquanto os formulários cuidam da interação. O compromisso é acoplar parte relevante do domínio ao SQL Server; mudanças de comportamento exigem evolução coordenada do VB6 e do T-SQL.

## 2. Atomicidade, repetição e concorrência são problemas diferentes

`usp_FinalizarVenda` abre uma transação com `XACT_ABORT` e tratamento TRY/CATCH. Uma falha de regra ou persistência desfaz os efeitos transacionais. Isso não garante que uma resposta chegue ao cliente após o commit.

Para uma tentativa cujo resultado ficou incerto, a interface conserva a chave de operação. A procedure usa `sp_getapplock` associado à transação e procura a venda existente antes de inserir. A chave deve ser reutilizada para a mesma tentativa; não é uma autorização para editar o carrinho e reenviar outra venda com o mesmo identificador.

Bloqueios de atualização e retenção protegem verificações de saldo. Os produtos são obtidos em ordem crescente de código para reduzir o risco de deadlocks, não para afirmar que todo deadlock é impossível. Isso é separado da concorrência otimista dos cadastros, que usa `rowversion` para rejeitar edições baseadas em uma versão antiga.

## 3. Preservar fatos comerciais

Itens guardam descrição, preço e custo da venda. Alterar o cadastro posteriormente não reescreve o histórico. Serviços integram o total, mas não geram baixa de estoque. Cancelamentos mudam o estado da venda e registram movimentos inversos; a justificativa permanece auditável.

Dinheiro, PIX e cartões são discriminados para que pagamentos eletrônicos não aumentem o dinheiro físico esperado no caixa. O troco é permitido apenas em dinheiro. O registro de estorno no sistema não realiza devolução bancária automática.

## 4. Relatórios com cardinalidade explícita

Uma venda pode ter vários itens e vários pagamentos. Somar depois de juntar essas duas coleções diretamente pode multiplicar valores. A [consulta por período](../database/consultas/001_VendasPorPeriodo.sql) mantém consultas separadas para totais de vendas e totais de itens, sem juntar pagamentos aos itens. O exemplo demonstra a decisão de modelagem da consulta, sem alegar medição de desempenho.

## 5. Acesso e limites de segurança

[modSeguranca](../src/modSeguranca.bas) usa PBKDF2-HMAC-SHA256, salt aleatório e 600 mil iterações. A senha é codificada em UTF-16LE para essa implementação. As procedures validam sessão e perfil; ocultar ou desabilitar um botão não é a única barreira de autorização.

A conexão usa a identidade Windows. Uma pessoa com privilégios administrativos no SQL Server pode acessar os dados diretamente, independentemente do perfil do aplicativo. `Trust Server Certificate=True` facilita a conexão local, mas não valida a cadeia do certificado do servidor. Uma implantação compartilhada exige desenho de permissões mínimas, certificados e distribuição do aplicativo; não é apresentada como pronta para esse cenário.

## 6. Evoluir a interface preservando os eventos

[modInterface](../src/modInterface.bas) aplica tema, tipografia e posicionamento nos formulários existentes. Painel, PDV e cadastros têm layouts próprios; telas auxiliares usam uma organização comum. Os atalhos chamam os eventos já existentes e respeitam controles habilitados. O uso de controles nativos evita adicionar uma biblioteca visual paga, mas mantém limitações de aparência e de DPI do VB6.

## 7. Reprodutibilidade com escopo explícito

O instalador recusa um destino existente. O runner utiliza um banco temporário com identificador único e dados fictícios. O projeto versiona definições de schema e de módulos, não um snapshot de dados reais. Os testes são manuais e locais; não há pipeline CI, benchmark ou validação de recuperação após queda de energia.
