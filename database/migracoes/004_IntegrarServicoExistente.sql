USE SEV_DB;
GO

CREATE OR ALTER PROCEDURE dbo.usp_ServicoInserir
    @Nome NVARCHAR(150),
    @Descricao NVARCHAR(500),
    @Preco DECIMAL(12,2),
    @DuracaoEstimadaMinutos INT,
    @ServicoID INT OUTPUT,
    @Token UNIQUEIDENTIFIER = NULL
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @UsuarioSessao INT; EXEC dbo.usp_ValidarSessao @Token,'CADASTRO',@UsuarioSessao OUTPUT;
    SET XACT_ABORT ON;

    -- Limpa o resultado antes de tentar cadastrar.
    SET @ServicoID = NULL;

    -- Remove espaços nas extremidades.
    SET @Nome = LTRIM(RTRIM(@Nome));

    -- Uma descrição vazia será armazenada como NULL.
    SET @Descricao =
        NULLIF(LTRIM(RTRIM(@Descricao)), N'');

    IF @Nome IS NULL OR LEN(@Nome) = 0
    BEGIN
        ;THROW 50010, N'Informe o nome do serviço.', 1;
    END;

    IF @Preco IS NULL OR @Preco < 0
    BEGIN
        ;THROW 50011,
            N'O preço deve ser maior ou igual a zero.',
            1;
    END;

    IF @DuracaoEstimadaMinutos IS NOT NULL
       AND @DuracaoEstimadaMinutos <= 0
    BEGIN
        ;THROW 50012,
            N'A duração deve ser positiva ou não informada.',
            1;
    END;

    INSERT INTO dbo.Servicos
    (
        Nome,
        Descricao,
        Preco,
        DuracaoEstimadaMinutos
    )
    VALUES
    (
        @Nome,
        @Descricao,
        @Preco,
        @DuracaoEstimadaMinutos
    );

    -- Devolve o código gerado para o novo serviço.
    SET @ServicoID = CONVERT(INT, SCOPE_IDENTITY());
END;
GO