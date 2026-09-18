-- Exclusivo para o banco descartável criado por Executar-Testes.ps1.
-- Não contém USE SEV_DB e não exporta nem altera dados de produção.
SET NOCOUNT ON;
IF DB_NAME() NOT LIKE N'SEV_TESTES[_]%' THROW 51080,N'Execute pelo runner em um banco SEV_TESTES_*.',1;
DECLARE @UsuarioID INT=(SELECT UsuarioID FROM dbo.Usuarios WHERE Usuario=N'qa_admin');
IF @UsuarioID IS NULL THROW 51081,N'Usuário fictício não foi preparado pelo runner.',1;
DECLARE @Casos TABLE(CasoID INT,NomeTeste NVARCHAR(100),Nome NVARCHAR(150),Preco DECIMAL(12,2),Duracao INT,ErroEsperado INT);
INSERT @Casos VALUES
 (1,N'Nome vazio',N'   ',15,10,50010),
 (2,N'Preço negativo',N'Serviço fictício',-1,10,50011),
 (3,N'Duração inválida',N'Serviço fictício',15,0,50012);
DECLARE @Caso INT=1,@NomeTeste NVARCHAR(100),@Nome NVARCHAR(150),@Preco DECIMAL(12,2),@Duracao INT,@Esperado INT,@Obtido INT,@ServicoID INT,@Token UNIQUEIDENTIFIER;
WHILE @Caso<=3
BEGIN
 SELECT @NomeTeste=NomeTeste,@Nome=Nome,@Preco=Preco,@Duracao=Duracao,@Esperado=ErroEsperado FROM @Casos WHERE CasoID=@Caso;
 SET @Obtido=0; SET @Token=NEWID();
 BEGIN TRY
  BEGIN TRANSACTION;
  INSERT dbo.SessoesAplicacao(Token,UsuarioID,ExpiraEm) VALUES(@Token,@UsuarioID,DATEADD(MINUTE,5,SYSDATETIME()));
  EXEC dbo.usp_ServicoInserir @Nome=@Nome,@Descricao=NULL,@Preco=@Preco,@DuracaoEstimadaMinutos=@Duracao,@ServicoID=@ServicoID OUTPUT,@Token=@Token;
  ROLLBACK TRANSACTION;
 END TRY
 BEGIN CATCH
  SET @Obtido=ERROR_NUMBER();
  IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
 END CATCH;
 IF @Obtido<>@Esperado THROW 51082,N'Validação de serviço retornou resultado diferente do esperado.',1;
 PRINT N'PASSOU: '+@NomeTeste;
 SET @Caso+=1;
END;
