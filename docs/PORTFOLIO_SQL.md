# Roteiro técnico para apresentar o SEV

Use este roteiro para explicar decisões e demonstrar o código, sem afirmar experiência de operação em produção que não ocorreu.

1. **Problema de negócio:** produto e serviço participam da mesma venda, mas só produto possui estoque. Mostre as constraints de ItensVenda e o histórico de preço/custo.
2. **Atomicidade:** em `usp_FinalizarVenda`, explique BEGIN TRANSACTION, TRY/CATCH, XACT_ABORT e o rollback de todos os efeitos quando uma regra falha.
3. **Concorrência:** mostre o bloqueio dos produtos em ordem de código e a verificação de saldo dentro da transação. Diferencie isso do `rowversion`, usado na edição dos cadastros.
4. **Repetição de requisição:** mostre ChaveOperacao, índice único e `sp_getapplock`. Uma falha de conexão não deve provocar uma segunda venda ao repetir a mesma tentativa.
5. **Integridade:** explique PK, FK, CHECK, índices filtrados e inativação de cadastros. Diferencie uma constraint estrutural de uma regra operacional validada por procedure.
6. **Cancelamento:** a venda não é apagada. Estoque e caixa recebem movimentos inversos, e a auditoria mantém a justificativa.
7. **Financeiro:** DECIMAL/Currency evitam uso de ponto flutuante binário para dinheiro. O troco não aumenta a receita. Pagamentos eletrônicos não entram no dinheiro físico do fechamento.
8. **Auditoria:** triggers registram antes/depois dos cadastros na mesma transação; procedures registram as operações sensíveis.
9. **Consultas:** mostre JOINs e agregações de vendas por período/vendedor. Não junte itens e pagamentos antes de somar, pois isso pode multiplicar valores.
10. **Testes:** execute o runner num banco descartável e explique uma rejeição esperada, como estoque insuficiente ou edição com rowversion antiga.

O [inventário](INVENTARIO_SQL.md) contém links para os arquivos exatos. Não é necessário exportar dados reais para demonstrar nenhuma dessas decisões.

## Limitações a explicar

Os perfis internos não restringem um administrador do SQL Server. O projeto não integra emissão fiscal, pagamentos bancários ou ordem de serviço. As quantidades são inteiras, as listagens possuem limites documentados e a interface modernizada mantém controles nativos do VB6. Esses limites fazem parte do escopo atual.
