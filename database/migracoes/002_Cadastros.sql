USE SEV_DB;
GO
SET XACT_ABORT ON;

IF COL_LENGTH('dbo.Categorias','Versao') IS NULL ALTER TABLE dbo.Categorias ADD Versao ROWVERSION;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CategoriasListar
    @Busca NVARCHAR(150)=N'', @IncluirInativos BIT=0
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT TOP (500) * FROM dbo.Categorias
 WHERE (@IncluirInativos=1 OR Ativo=1)
 AND (@Busca=N'' OR CHARINDEX(@Busca,Nome)>0 )
 ORDER BY Nome,CategoriaID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CategoriasObter @ID INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT * FROM dbo.Categorias WHERE CategoriaID=@ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CategoriasSituacao @ID INT,@Ativo BIT,@Versao BINARY(8)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 UPDATE dbo.Categorias SET Ativo=@Ativo WHERE CategoriaID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_CategoriasSalvar
 @ID INT=0, @Versao BINARY(8)=NULL,
 @Nome NVARCHAR(100)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
  SET @Nome=NULLIF(LTRIM(RTRIM(@Nome)),N'');
  IF @Nome IS NULL THROW 50101,N'Informe o nome.',1;
 IF @ID=0
 BEGIN
  INSERT dbo.Categorias (Nome) VALUES (@Nome);
  SET @ID=CONVERT(INT,SCOPE_IDENTITY());
 END
 ELSE
 BEGIN
  UPDATE dbo.Categorias SET Nome=@Nome WHERE CategoriaID=@ID AND Versao=@Versao;
  IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 END
 SELECT @ID AS ID;
END;
GO
IF COL_LENGTH('dbo.Clientes','Versao') IS NULL ALTER TABLE dbo.Clientes ADD Versao ROWVERSION;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ClientesListar
    @Busca NVARCHAR(150)=N'', @IncluirInativos BIT=0
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT TOP (500) * FROM dbo.Clientes
 WHERE (@IncluirInativos=1 OR Ativo=1)
 AND (@Busca=N'' OR CHARINDEX(@Busca,Nome)>0 )
 ORDER BY Nome,ClienteID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ClientesObter @ID INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT * FROM dbo.Clientes WHERE ClienteID=@ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ClientesSituacao @ID INT,@Ativo BIT,@Versao BINARY(8)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 UPDATE dbo.Clientes SET Ativo=@Ativo WHERE ClienteID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ClientesSalvar
 @ID INT=0, @Versao BINARY(8)=NULL,
 @Nome NVARCHAR(100),
@CPF_CNPJ VARCHAR(20),
@Telefone VARCHAR(20),
@Email NVARCHAR(254),
@Endereco NVARCHAR(250)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
  SET @Nome=NULLIF(LTRIM(RTRIM(@Nome)),N'');
 SET @CPF_CNPJ=NULLIF(LTRIM(RTRIM(@CPF_CNPJ)),N'');
 SET @Telefone=NULLIF(LTRIM(RTRIM(@Telefone)),N'');
 SET @Email=NULLIF(LTRIM(RTRIM(@Email)),N'');
 SET @Endereco=NULLIF(LTRIM(RTRIM(@Endereco)),N'');
  IF @Nome IS NULL THROW 50101,N'Informe o nome.',1;
 IF @ID=0
 BEGIN
  INSERT dbo.Clientes (Nome,CPF_CNPJ,Telefone,Email,Endereco) VALUES (@Nome,@CPF_CNPJ,@Telefone,@Email,@Endereco);
  SET @ID=CONVERT(INT,SCOPE_IDENTITY());
 END
 ELSE
 BEGIN
  UPDATE dbo.Clientes SET Nome=@Nome,CPF_CNPJ=@CPF_CNPJ,Telefone=@Telefone,Email=@Email,Endereco=@Endereco WHERE ClienteID=@ID AND Versao=@Versao;
  IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 END
 SELECT @ID AS ID;
END;
GO
IF COL_LENGTH('dbo.Fornecedores','Versao') IS NULL ALTER TABLE dbo.Fornecedores ADD Versao ROWVERSION;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FornecedoresListar
    @Busca NVARCHAR(150)=N'', @IncluirInativos BIT=0
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT TOP (500) * FROM dbo.Fornecedores
 WHERE (@IncluirInativos=1 OR Ativo=1)
 AND (@Busca=N'' OR CHARINDEX(@Busca,NomeRazao)>0 )
 ORDER BY NomeRazao,FornecedorID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FornecedoresObter @ID INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT * FROM dbo.Fornecedores WHERE FornecedorID=@ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FornecedoresSituacao @ID INT,@Ativo BIT,@Versao BINARY(8)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 UPDATE dbo.Fornecedores SET Ativo=@Ativo WHERE FornecedorID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_FornecedoresSalvar
 @ID INT=0, @Versao BINARY(8)=NULL,
 @NomeRazao NVARCHAR(150),
@CNPJ VARCHAR(20),
@Telefone VARCHAR(20),
@Email NVARCHAR(254)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
  SET @NomeRazao=NULLIF(LTRIM(RTRIM(@NomeRazao)),N'');
 SET @CNPJ=NULLIF(LTRIM(RTRIM(@CNPJ)),N'');
 SET @Telefone=NULLIF(LTRIM(RTRIM(@Telefone)),N'');
 SET @Email=NULLIF(LTRIM(RTRIM(@Email)),N'');
  IF @NomeRazao IS NULL THROW 50101,N'Informe o nome.',1;
 IF @ID=0
 BEGIN
  INSERT dbo.Fornecedores (NomeRazao,CNPJ,Telefone,Email) VALUES (@NomeRazao,@CNPJ,@Telefone,@Email);
  SET @ID=CONVERT(INT,SCOPE_IDENTITY());
 END
 ELSE
 BEGIN
  UPDATE dbo.Fornecedores SET NomeRazao=@NomeRazao,CNPJ=@CNPJ,Telefone=@Telefone,Email=@Email WHERE FornecedorID=@ID AND Versao=@Versao;
  IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 END
 SELECT @ID AS ID;
END;
GO
IF COL_LENGTH('dbo.Produtos','Versao') IS NULL ALTER TABLE dbo.Produtos ADD Versao ROWVERSION;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ProdutosListar
    @Busca NVARCHAR(150)=N'', @IncluirInativos BIT=0
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT TOP (500) * FROM dbo.Produtos
 WHERE (@IncluirInativos=1 OR Ativo=1)
 AND (@Busca=N'' OR CHARINDEX(@Busca,Nome)>0 OR CodigoBarras=@Busca)
 ORDER BY Nome,ProdutoID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ProdutosObter @ID INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT * FROM dbo.Produtos WHERE ProdutoID=@ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ProdutosSituacao @ID INT,@Ativo BIT,@Versao BINARY(8)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 UPDATE dbo.Produtos SET Ativo=@Ativo WHERE ProdutoID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ProdutosSalvar
 @ID INT=0, @Versao BINARY(8)=NULL,
 @Nome NVARCHAR(150),
@CodigoBarras VARCHAR(50),
@CategoriaID INT,
@FornecedorID INT,
@PrecoCusto DECIMAL(12,2),
@PrecoVenda DECIMAL(12,2),
@EstoqueMinimo INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
  SET @Nome=NULLIF(LTRIM(RTRIM(@Nome)),N'');
 SET @CodigoBarras=NULLIF(LTRIM(RTRIM(@CodigoBarras)),N'');
  IF @Nome IS NULL THROW 50101,N'Informe o nome.',1;
 IF @PrecoCusto IS NULL OR @PrecoVenda IS NULL OR @EstoqueMinimo IS NULL OR @PrecoCusto<0 OR @PrecoVenda<0 OR @EstoqueMinimo<0 THROW 50102,N'Preços e estoque mínimo inválidos.',1;
 IF @CategoriaID IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Categorias WHERE CategoriaID=@CategoriaID AND Ativo=1) THROW 50103,N'Categoria inativa ou inexistente.',1;
 IF @FornecedorID IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Fornecedores WHERE FornecedorID=@FornecedorID AND Ativo=1) THROW 50104,N'Fornecedor inativo ou inexistente.',1;

 IF @ID=0
 BEGIN
  INSERT dbo.Produtos (Nome,CodigoBarras,CategoriaID,FornecedorID,PrecoCusto,PrecoVenda,EstoqueMinimo) VALUES (@Nome,@CodigoBarras,@CategoriaID,@FornecedorID,@PrecoCusto,@PrecoVenda,@EstoqueMinimo);
  SET @ID=CONVERT(INT,SCOPE_IDENTITY());
 END
 ELSE
 BEGIN
  UPDATE dbo.Produtos SET Nome=@Nome,CodigoBarras=@CodigoBarras,CategoriaID=@CategoriaID,FornecedorID=@FornecedorID,PrecoCusto=@PrecoCusto,PrecoVenda=@PrecoVenda,EstoqueMinimo=@EstoqueMinimo WHERE ProdutoID=@ID AND Versao=@Versao;
  IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 END
 SELECT @ID AS ID;
END;
GO
IF COL_LENGTH('dbo.Servicos','Versao') IS NULL ALTER TABLE dbo.Servicos ADD Versao ROWVERSION;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ServicosListar
    @Busca NVARCHAR(150)=N'', @IncluirInativos BIT=0
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT TOP (500) * FROM dbo.Servicos
 WHERE (@IncluirInativos=1 OR Ativo=1)
 AND (@Busca=N'' OR CHARINDEX(@Busca,Nome)>0 )
 ORDER BY Nome,ServicoID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ServicosObter @ID INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'LEITURA',@UsuarioSessao OUTPUT;
 SELECT * FROM dbo.Servicos WHERE ServicoID=@ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ServicosSituacao @ID INT,@Ativo BIT,@Versao BINARY(8)
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 UPDATE dbo.Servicos SET Ativo=@Ativo WHERE ServicoID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_ServicosAtualizar
 @ID INT,@Versao BINARY(8),@Nome NVARCHAR(150),@Descricao NVARCHAR(500),@Preco DECIMAL(12,2),@DuracaoEstimadaMinutos INT
, @Token UNIQUEIDENTIFIER=NULL
AS
BEGIN
 SET NOCOUNT ON; DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT; SET XACT_ABORT ON;
 SET @Nome=NULLIF(LTRIM(RTRIM(@Nome)),N'');
 IF @Nome IS NULL OR @Preco IS NULL OR @Preco<0 OR (@DuracaoEstimadaMinutos IS NOT NULL AND @DuracaoEstimadaMinutos<=0)
 THROW 50101,N'Informe nome, preço não negativo e duração positiva ou vazia.',1;
 UPDATE dbo.Servicos SET Nome=@Nome,Descricao=NULLIF(LTRIM(RTRIM(@Descricao)),N''),Preco=@Preco,DuracaoEstimadaMinutos=@DuracaoEstimadaMinutos
 WHERE ServicoID=@ID AND Versao=@Versao;
 IF @@ROWCOUNT=0 THROW 50100,N'Registro alterado por outro usuário. Pesquise novamente.',1;
 SELECT @ID AS ID;
END;
GO