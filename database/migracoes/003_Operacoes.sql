USE SEV_DB;
GO
IF OBJECT_ID('dbo.SessoesAplicacao','U') IS NULL
CREATE TABLE dbo.SessoesAplicacao(Token UNIQUEIDENTIFIER NOT NULL PRIMARY KEY,UsuarioID INT NOT NULL REFERENCES dbo.Usuarios(UsuarioID),ExpiraEm DATETIME2 NOT NULL);
IF COL_LENGTH('dbo.Usuarios','TentativasLogin') IS NULL ALTER TABLE dbo.Usuarios ADD TentativasLogin INT NOT NULL DEFAULT(0),BloqueadoAte DATETIME2 NULL;
IF COL_LENGTH('dbo.Vendas','ChaveOperacao') IS NULL ALTER TABLE dbo.Vendas ADD ChaveOperacao UNIQUEIDENTIFIER NULL;
GO
IF NOT EXISTS(SELECT 1 FROM sys.indexes WHERE name='UX_Vendas_ChaveOperacao' AND object_id=OBJECT_ID('dbo.Vendas')) CREATE UNIQUE INDEX UX_Vendas_ChaveOperacao ON dbo.Vendas(ChaveOperacao) WHERE ChaveOperacao IS NOT NULL;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ValidarSessao @Token UNIQUEIDENTIFIER,@Modo VARCHAR(20)='LEITURA',@UsuarioID INT OUTPUT
AS
BEGIN
 SET NOCOUNT ON;
 SET @UsuarioID=NULL;
 DECLARE @Perfil NVARCHAR(50),@Cancela BIT,@Ajusta BIT;
 SELECT @UsuarioID=u.UsuarioID,@Perfil=p.Nome,@Cancela=p.PodeCancelarVenda,@Ajusta=p.PodeAjustarEstoque
 FROM dbo.SessoesAplicacao s JOIN dbo.Usuarios u ON u.UsuarioID=s.UsuarioID JOIN dbo.Perfis p ON p.PerfilID=u.PerfilID
 WHERE s.Token=@Token AND s.ExpiraEm>SYSDATETIME() AND u.Ativo=1 AND p.Ativo=1;
 IF @UsuarioID IS NULL THROW 50200,N'Sessão inválida ou expirada. Entre novamente.',1;
 DECLARE @Contexto VARBINARY(128)=CONVERT(BINARY(4),@UsuarioID);
 SET CONTEXT_INFO @Contexto;
 IF @Modo='ADMIN' AND @Perfil<>N'Administrador' THROW 50201,N'Operação exclusiva do administrador.',1;
 IF @Modo='CADASTRO' AND @Perfil NOT IN(N'Administrador',N'Gerente') THROW 50201,N'Seu perfil não pode alterar cadastros.',1;
 IF @Modo='ESTOQUE' AND @Ajusta=0 THROW 50201,N'Seu perfil não pode ajustar estoque.',1;
 IF @Modo='CANCELAR' AND @Cancela=0 THROW 50201,N'Seu perfil não pode cancelar vendas.',1;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_LoginPreparar @Usuario NVARCHAR(50)
AS
BEGIN
 SET NOCOUNT ON;
 SELECT CASE WHEN EXISTS(SELECT 1 FROM dbo.Usuarios) THEN 0 ELSE 1 END AS PrimeiroAcesso;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_LoginSalt @Usuario NVARCHAR(50)
AS
BEGIN
 SET NOCOUNT ON;
 SELECT SUBSTRING(SenhaHashCodificada,1,32) AS Salt FROM dbo.Usuarios WHERE Usuario=@Usuario AND Ativo=1
 UNION ALL SELECT CONVERT(VARCHAR(32),CRYPT_GEN_RANDOM(16),2) WHERE NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE Usuario=@Usuario AND Ativo=1);
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_PrimeiroAdministrador @Nome NVARCHAR(100),@Usuario NVARCHAR(50),@Credencial VARCHAR(512)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
 BEGIN TRANSACTION;
 IF EXISTS(SELECT 1 FROM dbo.Usuarios WITH(UPDLOCK,HOLDLOCK)) THROW 50202,N'O administrador inicial já existe.',1;
 IF @Credencial IS NULL OR LEN(@Credencial)<>97 OR SUBSTRING(@Credencial,33,1)<>':' OR TRY_CONVERT(VARBINARY(16),LEFT(@Credencial,32),2) IS NULL OR TRY_CONVERT(VARBINARY(32),RIGHT(@Credencial,64),2) IS NULL THROW 50203,N'Credencial inválida.',1;
 IF NULLIF(LTRIM(RTRIM(@Nome)),N'') IS NULL OR NULLIF(LTRIM(RTRIM(@Usuario)),N'') IS NULL THROW 50203,N'Informe nome e usuário.',1;
 INSERT dbo.Usuarios(Nome,Usuario,SenhaHashCodificada,PerfilID) SELECT @Nome,@Usuario,@Credencial,PerfilID FROM dbo.Perfis WHERE Nome=N'Administrador';
 DECLARE @ID INT=CONVERT(INT,SCOPE_IDENTITY());
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID) VALUES(@ID,N'Administrador inicial criado',N'Usuarios',@ID);
 COMMIT;
 SELECT @ID AS ID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_AutenticarUsuario @Usuario NVARCHAR(50),@Prova VARCHAR(64)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @ID INT,@Hash VARCHAR(512),@Bloqueado DATETIME2,@Token UNIQUEIDENTIFIER=NEWID(),@Perfil NVARCHAR(50);
 BEGIN TRANSACTION;
 SELECT @ID=u.UsuarioID,@Hash=u.SenhaHashCodificada,@Bloqueado=u.BloqueadoAte,@Perfil=p.Nome
 FROM dbo.Usuarios u WITH(UPDLOCK,HOLDLOCK) JOIN dbo.Perfis p ON p.PerfilID=u.PerfilID
 WHERE u.Usuario=@Usuario AND u.Ativo=1 AND p.Ativo=1;
 IF @ID IS NULL OR @Bloqueado>SYSDATETIME() OR RIGHT(@Hash,64) COLLATE Latin1_General_100_BIN2<>@Prova COLLATE Latin1_General_100_BIN2 OR @Prova IS NULL
 BEGIN
  IF @ID IS NOT NULL UPDATE dbo.Usuarios SET TentativasLogin=TentativasLogin+1,BloqueadoAte=CASE WHEN TentativasLogin+1>=5 THEN DATEADD(MINUTE,5,SYSDATETIME()) ELSE BloqueadoAte END WHERE UsuarioID=@ID;
  INSERT dbo.Auditoria(UsuarioID,Acao) VALUES(@ID,N'Falha de autenticação');
  COMMIT;
  THROW 50204,N'Usuário ou senha inválidos, ou acesso temporariamente bloqueado.',1;
 END;
 UPDATE dbo.Usuarios SET TentativasLogin=0,BloqueadoAte=NULL WHERE UsuarioID=@ID;
 DELETE dbo.SessoesAplicacao WHERE ExpiraEm<SYSDATETIME();
 INSERT dbo.SessoesAplicacao VALUES(@Token,@ID,DATEADD(HOUR,12,SYSDATETIME()));
 INSERT dbo.Auditoria(UsuarioID,Acao) VALUES(@ID,N'Login');
 COMMIT;
 SELECT CONVERT(VARCHAR(36),@Token) AS Token,@ID AS UsuarioID,@Perfil AS Perfil;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_Sair @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DELETE dbo.SessoesAplicacao WHERE Token=@Token; SELECT 1 AS OK; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_UsuariosListar @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'ADMIN',@U OUTPUT;
 SELECT u.UsuarioID,u.Nome,u.Usuario,u.PerfilID,p.Nome AS Perfil,u.Ativo FROM dbo.Usuarios u JOIN dbo.Perfis p ON p.PerfilID=u.PerfilID ORDER BY u.Nome; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_PerfisListar @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT; SELECT * FROM dbo.Perfis ORDER BY PerfilID; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_UsuariosSalvar @Token UNIQUEIDENTIFIER,@ID INT,@Nome NVARCHAR(100),@Usuario NVARCHAR(50),@PerfilID INT,@Ativo BIT,@Credencial VARCHAR(512)=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'ADMIN',@U OUTPUT;
 IF @ID=@U AND (@Ativo=0 OR @PerfilID<>(SELECT PerfilID FROM dbo.Perfis WHERE Nome=N'Administrador')) THROW 50205,N'Não remova seu próprio acesso administrativo.',1;
 IF NULLIF(LTRIM(RTRIM(@Nome)),N'') IS NULL OR NULLIF(LTRIM(RTRIM(@Usuario)),N'') IS NULL THROW 50203,N'Informe nome e usuário.',1;
 IF @ID=0 AND @Credencial IS NULL THROW 50203,N'Informe a senha do novo usuário.',1;
 IF @Credencial IS NOT NULL AND (LEN(@Credencial)<>97 OR SUBSTRING(@Credencial,33,1)<>':' OR TRY_CONVERT(VARBINARY(16),LEFT(@Credencial,32),2) IS NULL OR TRY_CONVERT(VARBINARY(32),RIGHT(@Credencial,64),2) IS NULL) THROW 50203,N'Credencial inválida.',1;
 BEGIN TRY BEGIN TRANSACTION;
 IF @ID=0 BEGIN INSERT dbo.Usuarios(Nome,Usuario,PerfilID,Ativo,SenhaHashCodificada) VALUES(@Nome,@Usuario,@PerfilID,@Ativo,@Credencial); SET @ID=CONVERT(INT,SCOPE_IDENTITY()); END
 ELSE BEGIN
 UPDATE dbo.Usuarios SET Nome=@Nome,Usuario=@Usuario,PerfilID=@PerfilID,Ativo=@Ativo,SenhaHashCodificada=COALESCE(@Credencial,SenhaHashCodificada),TentativasLogin=0,BloqueadoAte=NULL WHERE UsuarioID=@ID;
 IF @@ROWCOUNT=0 THROW 50206,N'Usuário não encontrado.',1;
 DELETE dbo.SessoesAplicacao WHERE UsuarioID=@ID AND Token<>@Token;
 END;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID) VALUES(@U,N'Usuário salvo',N'Usuarios',@ID);
 COMMIT; SELECT @ID AS ID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_EstoqueMovimentar @Token UNIQUEIDENTIFIER,@ProdutoID INT,@Tipo CHAR(1),@Quantidade INT,@Motivo NVARCHAR(250)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @U INT,@Saldo INT,@Permite BIT; EXEC dbo.usp_ValidarSessao @Token,'ESTOQUE',@U OUTPUT;
 IF @Tipo NOT IN('E','S') OR @Quantidade IS NULL OR @Quantidade<=0 OR NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL THROW 50300,N'Informe tipo, quantidade positiva e motivo.',1;
 BEGIN TRY BEGIN TRANSACTION;
 SELECT @Permite=PermitirEstoqueNegativo FROM dbo.Configuracoes WITH(HOLDLOCK) WHERE ConfiguracaoID=1;
 SELECT @Saldo=EstoqueAtual FROM dbo.Produtos WITH(UPDLOCK,HOLDLOCK) WHERE ProdutoID=@ProdutoID AND Ativo=1;
 IF @Saldo IS NULL THROW 50301,N'Produto inativo ou inexistente.',1;
 IF @Tipo='S' AND @Saldo<@Quantidade AND ISNULL(@Permite,0)=0 THROW 50302,N'Estoque insuficiente.',1;
 UPDATE dbo.Produtos SET EstoqueAtual=EstoqueAtual+CASE WHEN @Tipo='E' THEN @Quantidade ELSE -@Quantidade END WHERE ProdutoID=@ProdutoID;
 INSERT dbo.MovimentacoesEstoque(ProdutoID,UsuarioID,TipoMovimento,Quantidade,Motivo) VALUES(@ProdutoID,@U,@Tipo,@Quantidade,@Motivo);
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes) VALUES(@U,N'Ajuste de estoque',N'Produtos',@ProdutoID,@Motivo);
 COMMIT; SELECT EstoqueAtual FROM dbo.Produtos WHERE ProdutoID=@ProdutoID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_TerminaisListar @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT; SELECT * FROM dbo.TerminaisCaixa WHERE Ativo=1 ORDER BY Nome; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CaixasListar @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT;
 SELECT c.*,t.Nome AS Terminal,u.Nome AS Operador FROM dbo.Caixas c JOIN dbo.TerminaisCaixa t ON t.TerminalCaixaID=c.TerminalCaixaID JOIN dbo.Usuarios u ON u.UsuarioID=c.UsuarioAberturaID WHERE c.Status='ABERTO' ORDER BY c.CaixaID; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CaixaResumo @Token UNIQUEIDENTIFIER,@CaixaID INT
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT;
 SELECT c.CaixaID,c.Status,c.SaldoInicial,c.SaldoInicial+COALESCE(SUM(CASE WHEN m.FormaPagamento='DINHEIRO' THEN CASE WHEN m.TipoMovimento IN('VENDA','SUPRIMENTO') THEN m.Valor ELSE -m.Valor END ELSE 0 END),0) AS DinheiroEsperado,
 COALESCE(SUM(CASE WHEN m.FormaPagamento<>'DINHEIRO' THEN CASE WHEN m.TipoMovimento='VENDA' THEN m.Valor WHEN m.TipoMovimento='ESTORNO' THEN -m.Valor ELSE 0 END ELSE 0 END),0) AS EletronicosLiquidos
 FROM dbo.Caixas c LEFT JOIN dbo.MovimentosCaixa m ON m.CaixaID=c.CaixaID WHERE c.CaixaID=@CaixaID GROUP BY c.CaixaID,c.Status,c.SaldoInicial; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_AbrirCaixa @Token UNIQUEIDENTIFIER,@TerminalCaixaID INT,@SaldoInicial DECIMAL(12,2)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON; DECLARE @U INT,@ID INT; EXEC dbo.usp_ValidarSessao @Token,'CAIXA',@U OUTPUT;
 IF @SaldoInicial IS NULL OR @SaldoInicial<0 THROW 50400,N'Saldo inicial inválido.',1;
 BEGIN TRY BEGIN TRANSACTION;
 IF NOT EXISTS(SELECT 1 FROM dbo.TerminaisCaixa WITH(UPDLOCK,HOLDLOCK) WHERE TerminalCaixaID=@TerminalCaixaID AND Ativo=1) THROW 50400,N'Terminal inválido.',1;
 IF EXISTS(SELECT 1 FROM dbo.Caixas WHERE TerminalCaixaID=@TerminalCaixaID AND Status='ABERTO') THROW 50401,N'O terminal já possui caixa aberto.',1;
 INSERT dbo.Caixas(TerminalCaixaID,UsuarioAberturaID,SaldoInicial) VALUES(@TerminalCaixaID,@U,@SaldoInicial); SET @ID=CONVERT(INT,SCOPE_IDENTITY());
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID) VALUES(@U,N'Abertura de caixa',N'Caixas',@ID);
 COMMIT; SELECT @ID AS ID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_MovimentarCaixa @Token UNIQUEIDENTIFIER,@CaixaID INT,@Tipo VARCHAR(12),@Valor DECIMAL(12,2),@Motivo NVARCHAR(250)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON; DECLARE @U INT,@Saldo DECIMAL(12,2),@Operador INT; EXEC dbo.usp_ValidarSessao @Token,'CAIXA',@U OUTPUT;
 IF @Tipo NOT IN('SUPRIMENTO','SANGRIA') OR @Valor IS NULL OR @Valor<=0 OR NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL THROW 50402,N'Informe tipo, valor positivo e motivo.',1;
 BEGIN TRY BEGIN TRANSACTION;
 SELECT @Saldo=SaldoInicial,@Operador=UsuarioAberturaID FROM dbo.Caixas WITH(UPDLOCK,HOLDLOCK) WHERE CaixaID=@CaixaID AND Status='ABERTO';
 IF @Saldo IS NULL THROW 50403,N'Caixa não está aberto.',1;
 IF @Operador<>@U EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@U OUTPUT;
 SELECT @Saldo=@Saldo+COALESCE(SUM(CASE WHEN TipoMovimento IN('VENDA','SUPRIMENTO') THEN Valor ELSE -Valor END),0) FROM dbo.MovimentosCaixa WHERE CaixaID=@CaixaID AND FormaPagamento='DINHEIRO';
 IF @Tipo='SANGRIA' AND @Valor>@Saldo THROW 50404,N'Dinheiro insuficiente para a retirada.',1;
 INSERT dbo.MovimentosCaixa(CaixaID,UsuarioID,TipoMovimento,FormaPagamento,Valor,Descricao) VALUES(@CaixaID,@U,@Tipo,'DINHEIRO',@Valor,@Motivo);
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes) VALUES(@U,@Tipo,N'Caixas',@CaixaID,@Motivo);
 COMMIT; SELECT @CaixaID AS ID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FecharCaixa @Token UNIQUEIDENTIFIER,@CaixaID INT,@Conferido DECIMAL(12,2)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON; DECLARE @U INT,@Saldo DECIMAL(12,2),@Operador INT; EXEC dbo.usp_ValidarSessao @Token,'CAIXA',@U OUTPUT;
 IF @Conferido IS NULL OR @Conferido<0 THROW 50405,N'Valor conferido inválido.',1;
 BEGIN TRY BEGIN TRANSACTION;
 SELECT @Saldo=SaldoInicial,@Operador=UsuarioAberturaID FROM dbo.Caixas WITH(UPDLOCK,HOLDLOCK) WHERE CaixaID=@CaixaID AND Status='ABERTO';
 IF @Saldo IS NULL THROW 50403,N'Caixa não está aberto.',1;
 IF @Operador<>@U EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@U OUTPUT;
 SELECT @Saldo=@Saldo+COALESCE(SUM(CASE WHEN TipoMovimento IN('VENDA','SUPRIMENTO') THEN Valor ELSE -Valor END),0) FROM dbo.MovimentosCaixa WHERE CaixaID=@CaixaID AND FormaPagamento='DINHEIRO';
 UPDATE dbo.Caixas SET Status='FECHADO',DataFechamento=SYSDATETIME(),UsuarioFechamentoID=@U,SaldoEsperadoFechamento=@Saldo,SaldoConferidoFechamento=@Conferido WHERE CaixaID=@CaixaID;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID) VALUES(@U,N'Fechamento de caixa',N'Caixas',@CaixaID);
 COMMIT; SELECT @Saldo AS Esperado,@Conferido AS Conferido,@Conferido-@Saldo AS Diferenca;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FinalizarVenda @Token UNIQUEIDENTIFIER,@CaixaID INT,@ClienteID INT=NULL,@Desconto DECIMAL(12,2)=0,@Chave UNIQUEIDENTIFIER,@Itens NVARCHAR(MAX),@Pagamentos NVARCHAR(MAX)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @U INT,@VendaID INT,@Subtotal DECIMAL(12,2),@Total DECIMAL(12,2),@Limite DECIMAL(5,2),@Permite BIT,@Operador INT,@Produto INT,@Necessario INT,@Saldo INT,@ResultadoBloqueio INT;
 EXEC dbo.usp_ValidarSessao @Token,'VENDA',@U OUTPUT;
 DECLARE @XmlItens XML=TRY_CONVERT(XML,@Itens),@XmlPagamentos XML=TRY_CONVERT(XML,@Pagamentos);
 IF @Chave IS NULL OR @XmlItens IS NULL OR @XmlPagamentos IS NULL THROW 50500,N'Venda incompleta.',1;
 DECLARE @I TABLE(Sequencia INT IDENTITY PRIMARY KEY,Tipo CHAR(1),ID INT,Qtd INT,Descricao NVARCHAR(150),Preco DECIMAL(12,2),Custo DECIMAL(12,2));
 INSERT @I(Tipo,ID,Qtd) SELECT n.value('@tipo','CHAR(1)'),n.value('@id','INT'),n.value('@qtd','INT') FROM @XmlItens.nodes('/itens/item') t(n);
 DECLARE @P TABLE(Forma VARCHAR(10),Valor DECIMAL(12,2),Recebido DECIMAL(12,2));
 INSERT @P SELECT n.value('@forma','VARCHAR(10)'),n.value('@valor','DECIMAL(12,2)'),n.value('@recebido','DECIMAL(12,2)') FROM @XmlPagamentos.nodes('/pagamentos/pagamento') t(n);
 IF NOT EXISTS(SELECT 1 FROM @I) OR EXISTS(SELECT 1 FROM @I WHERE Tipo NOT IN('P','S') OR ID<=0 OR Qtd<=0) THROW 50501,N'Itens inválidos.',1;
 IF EXISTS(SELECT 1 FROM @P WHERE Forma NOT IN('DINHEIRO','PIX','DEBITO','CREDITO') OR Valor<=0 OR Recebido<Valor OR (Forma<>'DINHEIRO' AND Recebido<>Valor)) THROW 50502,N'Pagamentos inválidos.',1;
 BEGIN TRY BEGIN TRANSACTION;
 -- Serializa a chave de operação, tornando a repetição da mesma requisição segura.
 DECLARE @Recurso NVARCHAR(255)=N'SEV_VENDA_'+CONVERT(NVARCHAR(36),@Chave);
 EXEC @ResultadoBloqueio=sys.sp_getapplock @Resource=@Recurso,@LockMode='Exclusive',@LockOwner='Transaction',@LockTimeout=15000;
 IF @ResultadoBloqueio<0 THROW 50503,N'Venda em processamento. Aguarde e consulte o histórico.',1;
 SELECT @VendaID=VendaID FROM dbo.Vendas WHERE ChaveOperacao=@Chave;
 IF @VendaID IS NOT NULL BEGIN COMMIT; SELECT @VendaID AS VendaID,NumeroVenda,TotalVenda FROM dbo.Vendas WHERE VendaID=@VendaID; RETURN; END;
 SELECT @Operador=UsuarioAberturaID FROM dbo.Caixas WITH(UPDLOCK,HOLDLOCK) WHERE CaixaID=@CaixaID AND Status='ABERTO';
 IF @Operador IS NULL THROW 50504,N'Abra o caixa antes de vender.',1;
 IF @Operador<>@U EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@U OUTPUT;
 IF @ClienteID IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Clientes WITH(HOLDLOCK) WHERE ClienteID=@ClienteID AND Ativo=1) THROW 50505,N'Cliente inativo ou inexistente.',1;
 SELECT @Permite=PermitirEstoqueNegativo FROM dbo.Configuracoes WITH(HOLDLOCK) WHERE ConfiguracaoID=1;
 -- Bloqueia produtos sempre em ordem crescente para reduzir deadlocks.
 DECLARE produtos CURSOR LOCAL FAST_FORWARD FOR SELECT ID,SUM(Qtd) FROM @I WHERE Tipo='P' GROUP BY ID ORDER BY ID;
 OPEN produtos; FETCH NEXT FROM produtos INTO @Produto,@Necessario;
 WHILE @@FETCH_STATUS=0
 BEGIN
  SET @Saldo=NULL;
  SELECT @Saldo=EstoqueAtual FROM dbo.Produtos WITH(UPDLOCK,HOLDLOCK) WHERE ProdutoID=@Produto AND Ativo=1;
  IF @Saldo IS NULL THROW 50506,N'Produto inativo ou inexistente.',1;
  IF @Saldo<@Necessario AND ISNULL(@Permite,0)=0 THROW 50507,N'Estoque insuficiente. A venda não foi gravada.',1;
  FETCH NEXT FROM produtos INTO @Produto,@Necessario;
 END;
 CLOSE produtos; DEALLOCATE produtos;
 UPDATE i SET Descricao=p.Nome,Preco=p.PrecoVenda,Custo=p.PrecoCusto FROM @I i JOIN dbo.Produtos p ON p.ProdutoID=i.ID WHERE i.Tipo='P';
 UPDATE i SET Descricao=s.Nome,Preco=s.Preco FROM @I i JOIN dbo.Servicos s WITH(HOLDLOCK) ON s.ServicoID=i.ID AND s.Ativo=1 WHERE i.Tipo='S';
 IF EXISTS(SELECT 1 FROM @I WHERE Preco IS NULL) THROW 50508,N'Serviço inativo ou inexistente.',1;
 SELECT @Subtotal=SUM(CONVERT(DECIMAL(18,2),Qtd)*Preco) FROM @I;
 SELECT @Limite=p.PercentualMaximoDesconto FROM dbo.Usuarios u JOIN dbo.Perfis p ON p.PerfilID=u.PerfilID WHERE u.UsuarioID=@U;
 IF @Desconto IS NULL OR @Desconto<0 OR @Desconto>@Subtotal OR @Desconto*100>@Subtotal*@Limite THROW 50509,N'Desconto inválido ou acima da permissão do perfil.',1;
 SET @Total=@Subtotal-@Desconto;
 IF COALESCE((SELECT SUM(Valor) FROM @P),0)<>@Total THROW 50510,N'Pagamentos não conferem com o total. Confira os preços atuais.',1;
 INSERT dbo.Vendas(NumeroVenda,ClienteID,UsuarioID,CaixaID,SubTotal,Desconto,Status,ChaveOperacao)
 VALUES(LEFT(REPLACE(CONVERT(VARCHAR(36),@Chave),'-',''),20),@ClienteID,@U,@CaixaID,@Subtotal,@Desconto,'ABERTA',@Chave);
 SET @VendaID=CONVERT(INT,SCOPE_IDENTITY());
 UPDATE dbo.Vendas SET NumeroVenda='V'+CONVERT(VARCHAR(19),@VendaID) WHERE VendaID=@VendaID;
 INSERT dbo.ItensVenda(VendaID,TipoItem,ProdutoID,ServicoID,DescricaoItem,Quantidade,PrecoUnitario,CustoUnitario)
 SELECT @VendaID,Tipo,CASE WHEN Tipo='P' THEN ID END,CASE WHEN Tipo='S' THEN ID END,Descricao,Qtd,Preco,Custo FROM @I;
 INSERT dbo.PagamentosVenda(VendaID,FormaPagamento,Valor,ValorRecebido) SELECT @VendaID,Forma,Valor,Recebido FROM @P;
 UPDATE p SET EstoqueAtual=p.EstoqueAtual-i.Qtd FROM dbo.Produtos p JOIN(SELECT ID,SUM(Qtd) Qtd FROM @I WHERE Tipo='P' GROUP BY ID)i ON i.ID=p.ProdutoID;
 INSERT dbo.MovimentacoesEstoque(ProdutoID,VendaID,UsuarioID,TipoMovimento,Quantidade,Motivo) SELECT ID,@VendaID,@U,'S',SUM(Qtd),N'Venda' FROM @I WHERE Tipo='P' GROUP BY ID;
 INSERT dbo.MovimentosCaixa(CaixaID,VendaID,UsuarioID,TipoMovimento,FormaPagamento,Valor,Descricao) SELECT @CaixaID,@VendaID,@U,'VENDA',Forma,SUM(Valor),N'Finalização da venda' FROM @P GROUP BY Forma;
 UPDATE dbo.Vendas SET Status='FINALIZADA' WHERE VendaID=@VendaID;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID) VALUES(@U,N'Venda finalizada',N'Vendas',@VendaID);
 COMMIT;
 SELECT VendaID,NumeroVenda,TotalVenda FROM dbo.Vendas WHERE VendaID=@VendaID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CancelarVenda @Token UNIQUEIDENTIFIER,@VendaID INT,@CaixaEstornoID INT,@Motivo NVARCHAR(250)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 DECLARE @U INT,@Estado VARCHAR(12),@Caixa INT; EXEC dbo.usp_ValidarSessao @Token,'CANCELAR',@U OUTPUT;
 IF NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL THROW 50520,N'Informe o motivo do cancelamento.',1;
 BEGIN TRY BEGIN TRANSACTION;
 SELECT @Caixa=CaixaID FROM dbo.Caixas WITH(UPDLOCK,HOLDLOCK) WHERE CaixaID=@CaixaEstornoID AND Status='ABERTO';
 IF @Caixa IS NULL THROW 50521,N'Selecione um caixa aberto para registrar a devolução.',1;
 SELECT @Estado=Status FROM dbo.Vendas WITH(UPDLOCK,HOLDLOCK) WHERE VendaID=@VendaID;
 IF @Estado IS NULL OR @Estado<>'FINALIZADA' THROW 50522,N'A venda não está finalizada ou já foi cancelada.',1;
 DECLARE @Dinheiro DECIMAL(12,2),@Devolver DECIMAL(12,2);
 SELECT @Dinheiro=SaldoInicial FROM dbo.Caixas WHERE CaixaID=@CaixaEstornoID;
 SELECT @Dinheiro=@Dinheiro+COALESCE(SUM(CASE WHEN TipoMovimento IN('VENDA','SUPRIMENTO') THEN Valor ELSE -Valor END),0) FROM dbo.MovimentosCaixa WHERE CaixaID=@CaixaEstornoID AND FormaPagamento='DINHEIRO';
 SELECT @Devolver=COALESCE(SUM(Valor),0) FROM dbo.PagamentosVenda WHERE VendaID=@VendaID AND FormaPagamento='DINHEIRO';
 IF @Devolver>@Dinheiro THROW 50523,N'Dinheiro insuficiente para devolver. Registre um suprimento ou selecione outro caixa.',1;

 UPDATE p SET EstoqueAtual=p.EstoqueAtual+i.Qtd FROM dbo.Produtos p JOIN(SELECT ProdutoID,SUM(Quantidade) Qtd FROM dbo.ItensVenda WHERE VendaID=@VendaID AND TipoItem='P' GROUP BY ProdutoID)i ON i.ProdutoID=p.ProdutoID;
 INSERT dbo.MovimentacoesEstoque(ProdutoID,VendaID,UsuarioID,TipoMovimento,Quantidade,Motivo) SELECT ProdutoID,@VendaID,@U,'E',SUM(Quantidade),@Motivo FROM dbo.ItensVenda WHERE VendaID=@VendaID AND TipoItem='P' GROUP BY ProdutoID;
 INSERT dbo.MovimentosCaixa(CaixaID,VendaID,UsuarioID,TipoMovimento,FormaPagamento,Valor,Descricao) SELECT @CaixaEstornoID,@VendaID,@U,'ESTORNO',FormaPagamento,SUM(Valor),@Motivo FROM dbo.PagamentosVenda WHERE VendaID=@VendaID GROUP BY FormaPagamento;
 UPDATE dbo.Vendas SET Status='CANCELADA',UsuarioCancelamentoID=@U,DataCancelamento=SYSDATETIME(),MotivoCancelamento=@Motivo WHERE VendaID=@VendaID;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes) VALUES(@U,N'Venda cancelada',N'Vendas',@VendaID,@Motivo);
 COMMIT; SELECT @VendaID AS ID;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ConsultaOperacional @Token UNIQUEIDENTIFIER,@Tipo VARCHAR(30),@Inicio DATE,@Fim DATE,@ID INT=0
AS
BEGIN
 SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@U OUTPUT;
 IF @Inicio IS NULL OR @Fim IS NULL OR @Fim<@Inicio THROW 50600,N'Período inválido.',1;
 IF @Tipo='VENDAS' SELECT TOP(1000) v.VendaID,v.NumeroVenda,v.DataVenda,u.Nome AS Vendedor,v.TotalVenda,v.Status FROM dbo.Vendas v JOIN dbo.Usuarios u ON u.UsuarioID=v.UsuarioID WHERE v.DataVenda>=@Inicio AND v.DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim)) ORDER BY v.VendaID DESC;
 ELSE IF @Tipo='ITENS' SELECT TipoItem,DescricaoItem,Quantidade,PrecoUnitario,SubTotal FROM dbo.ItensVenda WHERE VendaID=@ID;
 ELSE IF @Tipo='PAGAMENTOS' SELECT FormaPagamento,Valor,ValorRecebido,Troco FROM dbo.PagamentosVenda WHERE VendaID=@ID;
 ELSE IF @Tipo='ESTOQUE' SELECT TOP(1000) m.MovimentacaoID,p.Nome,m.TipoMovimento,m.Quantidade,m.Motivo,m.DataMovimento FROM dbo.MovimentacoesEstoque m JOIN dbo.Produtos p ON p.ProdutoID=m.ProdutoID WHERE m.DataMovimento>=@Inicio AND m.DataMovimento<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim)) ORDER BY m.MovimentacaoID DESC;
 ELSE IF @Tipo='MINIMO' SELECT ProdutoID,Nome,EstoqueAtual,EstoqueMinimo FROM dbo.Produtos WHERE Ativo=1 AND EstoqueAtual<=EstoqueMinimo ORDER BY Nome;
 ELSE IF @Tipo='CAIXA' SELECT TOP(1000) CaixaID,TipoMovimento,FormaPagamento,Valor,Descricao,DataMovimento FROM dbo.MovimentosCaixa WHERE DataMovimento>=@Inicio AND DataMovimento<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim)) ORDER BY MovimentoID DESC;
 ELSE IF @Tipo='AUDITORIA' BEGIN EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@U OUTPUT; SELECT TOP(1000) AuditoriaID,UsuarioID,Acao,Entidade,RegistroID,Detalhes,DataAcao FROM dbo.Auditoria WHERE DataAcao>=@Inicio AND DataAcao<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim)) ORDER BY AuditoriaID DESC; END
 ELSE IF @Tipo='VENDEDOR' SELECT u.Nome,COUNT(*) AS Vendas,SUM(v.TotalVenda) AS Total FROM dbo.Vendas v JOIN dbo.Usuarios u ON u.UsuarioID=v.UsuarioID WHERE v.Status='FINALIZADA' AND v.DataVenda>=@Inicio AND v.DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim)) GROUP BY u.Nome,u.UsuarioID;
 ELSE IF @Tipo='RESUMO' SELECT COUNT(*) AS VendasFinalizadas,COALESCE(SUM(TotalVenda),0) AS Total,COALESCE(SUM(Desconto),0) AS Descontos FROM dbo.Vendas WHERE Status='FINALIZADA' AND DataVenda>=@Inicio AND DataVenda<DATEADD(DAY,1,CONVERT(DATETIME2,@Fim));
 ELSE THROW 50601,N'Consulta desconhecida.',1;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ConfiguracoesObter @Token UNIQUEIDENTIFIER
AS BEGIN SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'ADMIN',@U OUTPUT; SELECT * FROM dbo.Configuracoes WHERE ConfiguracaoID=1; END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ConfiguracoesSalvar @Token UNIQUEIDENTIFIER,@Nome NVARCHAR(150),@PermitirNegativo BIT,@LimiteAdministrador DECIMAL(5,2),@LimiteGerente DECIMAL(5,2),@LimiteVendedor DECIMAL(5,2)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'ADMIN',@U OUTPUT;
 IF NULLIF(LTRIM(RTRIM(@Nome)),N'') IS NULL OR @PermitirNegativo IS NULL OR @LimiteAdministrador IS NULL OR @LimiteGerente IS NULL OR @LimiteVendedor IS NULL OR @LimiteAdministrador NOT BETWEEN 0 AND 100 OR @LimiteGerente NOT BETWEEN 0 AND 100 OR @LimiteVendedor NOT BETWEEN 0 AND 100 THROW 50602,N'Configuração inválida.',1;
 BEGIN TRY BEGIN TRANSACTION;
 UPDATE dbo.Configuracoes SET NomeEstabelecimento=@Nome,PermitirEstoqueNegativo=@PermitirNegativo WHERE ConfiguracaoID=1;
 UPDATE dbo.Perfis SET PercentualMaximoDesconto=CASE Nome WHEN N'Administrador' THEN @LimiteAdministrador WHEN N'Gerente' THEN @LimiteGerente WHEN N'Vendedor' THEN @LimiteVendedor END WHERE Nome IN(N'Administrador',N'Gerente',N'Vendedor');
 INSERT dbo.Auditoria(UsuarioID,Acao) VALUES(@U,N'Configurações alteradas'); COMMIT; SELECT 1 AS OK;
 END TRY BEGIN CATCH IF XACT_STATE()<>0 ROLLBACK; THROW; END CATCH;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_BackupCriar @Token UNIQUEIDENTIFIER,@ArquivoBackup NVARCHAR(1000)=NULL OUTPUT
AS
BEGIN
 SET NOCOUNT ON; DECLARE @U INT; EXEC dbo.usp_ValidarSessao @Token,'ADMIN',@U OUTPUT;
 DECLARE @Pasta NVARCHAR(500)=CONVERT(NVARCHAR(500),SERVERPROPERTY('InstanceDefaultBackupPath'));
 IF @Pasta IS NULL THROW 50610,N'Pasta padrão de backup não encontrada.',1;
 DECLARE @Arquivo NVARCHAR(1000)=@Pasta+CASE WHEN RIGHT(@Pasta,1)='\' THEN '' ELSE '\' END+N'SEV_DB_'+REPLACE(CONVERT(NVARCHAR(36),NEWID()),'-','')+N'.bak';
 BACKUP DATABASE SEV_DB TO DISK=@Arquivo WITH COPY_ONLY,CHECKSUM;
 RESTORE VERIFYONLY FROM DISK=@Arquivo WITH CHECKSUM;
 INSERT dbo.Auditoria(UsuarioID,Acao,Detalhes) VALUES(@U,N'Backup criado e verificado',@Arquivo);
 SET @ArquivoBackup=@Arquivo;
 SELECT @Arquivo AS Arquivo;
END;
GO
