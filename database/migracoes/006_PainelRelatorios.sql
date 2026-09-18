USE SEV_DB;
GO
CREATE OR ALTER VIEW dbo.vw_VendasFinalizadas
AS
SELECT v.VendaID,v.NumeroVenda,v.DataVenda,v.ClienteID,v.UsuarioID,v.CaixaID,v.SubTotal,v.Desconto,v.TotalVenda
FROM dbo.Vendas v WHERE v.Status='FINALIZADA';
GO
CREATE OR ALTER PROCEDURE dbo.usp_PainelResumo @Token UNIQUEIDENTIFIER
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT;
 SELECT (SELECT NomeEstabelecimento FROM dbo.Configuracoes WHERE ConfiguracaoID=1) AS Estabelecimento,
 (SELECT COUNT(*) FROM dbo.vw_VendasFinalizadas WHERE DataVenda>=CONVERT(DATE,SYSDATETIME())) AS VendasHoje,
 (SELECT COALESCE(SUM(TotalVenda),0) FROM dbo.vw_VendasFinalizadas WHERE DataVenda>=CONVERT(DATE,SYSDATETIME())) AS TotalHoje,
 (SELECT COUNT(*) FROM dbo.Produtos WHERE Ativo=1 AND EstoqueAtual<=EstoqueMinimo) AS ProdutosNoMinimo,
 (SELECT COUNT(*) FROM dbo.Caixas WHERE Status='ABERTO') AS CaixasAbertos;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_RelatorioItens @Token UNIQUEIDENTIFIER,@Inicio DATE,@Fim DATE
AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@U OUTPUT;
 IF @Inicio IS NULL OR @Fim IS NULL OR @Fim<@Inicio THROW 50600,N'Período inválido.',1;
 SELECT i.TipoItem,COALESCE(i.ProdutoID,i.ServicoID) AS Codigo,i.DescricaoItem,
 SUM(CONVERT(BIGINT,i.Quantidade)) AS Quantidade,SUM(i.SubTotal) AS TotalBruto,
 SUM(CONVERT(DECIMAL(18,2),i.Quantidade)*COALESCE(i.CustoUnitario,0)) AS CustoHistorico
 FROM dbo.ItensVenda i JOIN dbo.vw_VendasFinalizadas v ON v.VendaID=i.VendaID
 WHERE v.DataVenda>=@Inicio AND v.DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim))
 GROUP BY i.TipoItem,i.ProdutoID,i.ServicoID,i.DescricaoItem ORDER BY i.TipoItem,i.DescricaoItem;
END;
GO
