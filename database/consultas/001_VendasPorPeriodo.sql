-- Exemplo de consulta analítica, somente leitura. Não é um dump de dados.
-- No aplicativo, consultas equivalentes são feitas por procedures autenticadas.
DECLARE @Inicio DATE=CONVERT(DATE,SYSDATETIME()),@Fim DATE=CONVERT(DATE,SYSDATETIME());
SELECT u.UsuarioID,u.Nome AS Vendedor,COUNT(*) AS QuantidadeVendas,
       SUM(v.SubTotal) AS ValorBruto,SUM(v.Desconto) AS Descontos,
       SUM(v.TotalVenda) AS ValorLiquido,AVG(v.TotalVenda) AS TicketMedio
FROM dbo.Vendas v JOIN dbo.Usuarios u ON u.UsuarioID=v.UsuarioID
WHERE v.Status='FINALIZADA' AND v.DataVenda>=@Inicio
  AND v.DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim))
GROUP BY u.UsuarioID,u.Nome ORDER BY ValorLiquido DESC;

-- Agrega itens sem juntar pagamentos: evita multiplicar os valores por parcela.
SELECT i.TipoItem,SUM(CONVERT(BIGINT,i.Quantidade)) AS Quantidade,
       SUM(i.SubTotal) AS ValorBrutoDosItens
FROM dbo.ItensVenda i JOIN dbo.Vendas v ON v.VendaID=i.VendaID
WHERE v.Status='FINALIZADA' AND v.DataVenda>=@Inicio
  AND v.DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim))
GROUP BY i.TipoItem;
