# Validação — 18/09/2026

## Ambiente e método

Visual Basic 6, referência ADO 2.8, driver MSOLEDBSQL19 de 32 bits e SQL Server Express 17. Testes executados em SEV_QA_20260918; restauração ensaiada em SEV_RESTAURACAO_QA_20260918. Nenhum cadastro fictício foi gravado no SEV_DB principal.

## Resultados confirmados

- Compilação da versão final sem erros.
- Derivação PBKDF2 produzida no VB6 comparada com hashlib do Python: resultado idêntico.
- Login e sessão via ADO 32 bits; parâmetros opcionais e recordset desconectado.
- Cadastro e edição pela tela, incluindo rowversion.
- Carga das telas de clientes, fornecedores, produtos, serviços, usuários, configurações, estoque, consultas e caixa.
- Pesquisa de produto por código de barras na tela do PDV.
- Venda mista real pela rotina de eventos VB6; XML parametrizado transmitido como NVARCHAR(MAX).
- Produto R$ 20 + serviço R$ 15: total R$ 35; somente o produto baixa estoque.
- Pagamento em dinheiro + PIX e troco calculados corretamente.
- Repetição da chave da venda não duplica registros.
- Venda somente de serviço mantém estoque inalterado.
- Rejeição de estoque insuficiente, desconto não permitido, soma de pagamentos incorreta e venda com caixa fechado.
- Cancelamento repõe produtos e estorna valores; repetição é recusada.
- Suprimento, sangria e recusa de retirada acima do dinheiro disponível.
- Fechamento calculou diferença negativa de R$ 1 em cenário proposital e diferença zero pelo fluxo VB6.
- Edição com rowversion antiga recusada; auditoria registra valores anteriores.
- Vendedor impedido de alterar cadastros e estoque; sessão encerrada não permite consulta autenticada.
- Duas vendas concorrentes da última unidade, em caixas diferentes: uma aceita, uma recusada, saldo final zero.
- Backup criado pelo VB6, verificado e restaurado em banco separado; DBCC CHECKDB concluído sem erro.
- Revisão visual do menu/painel e PDV em execução dentro do VB6.

## Limitações da validação

Não houve operação real de loja, integração bancária, periféricos ou emissão fiscal — essas integrações estão fora do escopo. Não foi feito teste de carga prolongado, queda de energia ou revisão independente de segurança. As rotinas críticas foram exercitadas com dados fictícios; isso não substitui a conferência operacional inicial pelo responsável da loja.

O Smart App Control bloqueou uma versão do executável de QA. A execução do fonte dentro do ambiente VB6 permitiu continuar a validação, sem mudar proteções do Windows. A liberação de um EXE novo depende das políticas de confiança/assinatura da máquina.

## Verificação após implantação

O SEV_DB principal passou por DBCC CHECKDB sem erros: 17 tabelas, 45 procedures e 5 triggers. Usuários e vendas permaneceram vazios, e a procedure de primeiro acesso retornou 1. O executável final instalado em `bin\SistemaEstoqueVendas.exe` foi iniciado com sucesso e apresentou a janela **SEV - Acesso**. O bloqueio do executável instrumentado de QA não ocorreu na versão final. Nenhuma proteção do Windows foi alterada.
