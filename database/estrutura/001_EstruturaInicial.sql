/* =========================================================
   BANCO: SEV_DB
   Aplicar somente em banco novo ou sem tabelas de usuário.
   Este script não exclui tabelas existentes.
   ========================================================= */

USE master;
GO

IF DB_ID(N'SEV_DB') IS NULL
BEGIN
    EXEC(N'CREATE DATABASE [SEV_DB];');
END;
GO

USE SEV_DB;
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;

-- Opções necessárias para os índices filtrados.
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_PADDING ON;
SET ANSI_WARNINGS ON;
SET CONCAT_NULL_YIELDS_NULL ON;
SET ARITHABORT ON;
SET NUMERIC_ROUNDABORT OFF;

-- Impede aplicar uma estrutura nova sobre tabelas existentes.
IF EXISTS (
    SELECT 1
    FROM sys.tables
    WHERE is_ms_shipped = 0
)
BEGIN
    ;THROW 50001,
        N'SEV_DB já possui tabelas. Não foi aplicada a estrutura. Use um script de migração.',
        1;
END;

BEGIN TRY
    BEGIN TRANSACTION;

    /* 1. PERFIS */

    CREATE TABLE dbo.Perfis
    (
        PerfilID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Perfis PRIMARY KEY,

        Nome NVARCHAR(50) NOT NULL
            CONSTRAINT UQ_Perfis_Nome UNIQUE,

        PercentualMaximoDesconto DECIMAL(5,2) NOT NULL
            CONSTRAINT DF_Perfis_Desconto DEFAULT (0),

        PodeCancelarVenda BIT NOT NULL
            CONSTRAINT DF_Perfis_Cancelamento DEFAULT (0),

        PodeAjustarEstoque BIT NOT NULL
            CONSTRAINT DF_Perfis_Ajuste DEFAULT (0),

        Ativo BIT NOT NULL
            CONSTRAINT DF_Perfis_Ativo DEFAULT (1),

        CONSTRAINT CK_Perfis_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0),

        CONSTRAINT CK_Perfis_Desconto
            CHECK (PercentualMaximoDesconto BETWEEN 0 AND 100)
    );

    /* 2. USUÁRIOS */

    CREATE TABLE dbo.Usuarios
    (
        UsuarioID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Usuarios PRIMARY KEY,

        Nome NVARCHAR(100) NOT NULL,

        Usuario NVARCHAR(50) NOT NULL
            CONSTRAINT UQ_Usuarios_Usuario UNIQUE,

        /*
          Guarda uma representação produzida por biblioteca
          de hash de senha: algoritmo, parâmetros, salt e hash.
          Não colocar a senha original neste campo.
        */
        SenhaHashCodificada VARCHAR(512) NOT NULL,

        PerfilID INT NOT NULL,

        Ativo BIT NOT NULL
            CONSTRAINT DF_Usuarios_Ativo DEFAULT (1),

        DataCadastro DATETIME2(0) NOT NULL
            CONSTRAINT DF_Usuarios_Data DEFAULT (SYSDATETIME()),

        CONSTRAINT FK_Usuarios_Perfis
            FOREIGN KEY (PerfilID)
            REFERENCES dbo.Perfis(PerfilID),

        CONSTRAINT CK_Usuarios_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0),

        CONSTRAINT CK_Usuarios_Usuario
            CHECK (LEN(LTRIM(RTRIM(Usuario))) > 0),

        CONSTRAINT CK_Usuarios_Hash
            CHECK (LEN(LTRIM(RTRIM(SenhaHashCodificada))) > 0)
    );

    /* 3. CLIENTES */

    CREATE TABLE dbo.Clientes
    (
        ClienteID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Clientes PRIMARY KEY,

        Nome NVARCHAR(100) NOT NULL,
        CPF_CNPJ VARCHAR(20) NULL,
        Telefone VARCHAR(20) NULL,
        Email NVARCHAR(254) NULL,
        Endereco NVARCHAR(250) NULL,

        DataCadastro DATETIME2(0) NOT NULL
            CONSTRAINT DF_Clientes_Data DEFAULT (SYSDATETIME()),

        Ativo BIT NOT NULL
            CONSTRAINT DF_Clientes_Ativo DEFAULT (1),

        CONSTRAINT CK_Clientes_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0)
    );

    /* 4. CATEGORIAS */

    CREATE TABLE dbo.Categorias
    (
        CategoriaID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Categorias PRIMARY KEY,

        Nome NVARCHAR(100) NOT NULL
            CONSTRAINT UQ_Categorias_Nome UNIQUE,

        Ativo BIT NOT NULL
            CONSTRAINT DF_Categorias_Ativo DEFAULT (1),

        CONSTRAINT CK_Categorias_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0)
    );

    /* 5. FORNECEDORES */

    CREATE TABLE dbo.Fornecedores
    (
        FornecedorID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Fornecedores PRIMARY KEY,

        NomeRazao NVARCHAR(150) NOT NULL,
        CNPJ VARCHAR(20) NULL,
        Telefone VARCHAR(20) NULL,
        Email NVARCHAR(254) NULL,

        Ativo BIT NOT NULL
            CONSTRAINT DF_Fornecedores_Ativo DEFAULT (1),

        CONSTRAINT CK_Fornecedores_Nome
            CHECK (LEN(LTRIM(RTRIM(NomeRazao))) > 0)
    );

    /* 6. PRODUTOS */

    CREATE TABLE dbo.Produtos
    (
        ProdutoID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Produtos PRIMARY KEY,

        CodigoBarras VARCHAR(50) NULL,
        Nome NVARCHAR(150) NOT NULL,
        CategoriaID INT NULL,

        -- Fornecedor principal do cadastro.
        FornecedorID INT NULL,

        PrecoCusto DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Produtos_Custo DEFAULT (0),

        PrecoVenda DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Produtos_Preco DEFAULT (0),

        EstoqueAtual INT NOT NULL
            CONSTRAINT DF_Produtos_Estoque DEFAULT (0),

        EstoqueMinimo INT NOT NULL
            CONSTRAINT DF_Produtos_Minimo DEFAULT (0),

        Ativo BIT NOT NULL
            CONSTRAINT DF_Produtos_Ativo DEFAULT (1),

        CONSTRAINT FK_Produtos_Categorias
            FOREIGN KEY (CategoriaID)
            REFERENCES dbo.Categorias(CategoriaID),

        CONSTRAINT FK_Produtos_Fornecedores
            FOREIGN KEY (FornecedorID)
            REFERENCES dbo.Fornecedores(FornecedorID),

        CONSTRAINT CK_Produtos_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0),

        CONSTRAINT CK_Produtos_CodigoBarras
            CHECK (
                CodigoBarras IS NULL
                OR LEN(LTRIM(RTRIM(CodigoBarras))) > 0
            ),

        CONSTRAINT CK_Produtos_Precos
            CHECK (PrecoCusto >= 0 AND PrecoVenda >= 0),

        CONSTRAINT CK_Produtos_Minimo
            CHECK (EstoqueMinimo >= 0)
    );

    CREATE UNIQUE INDEX UX_Produtos_CodigoBarras
        ON dbo.Produtos(CodigoBarras)
        WHERE CodigoBarras IS NOT NULL;

    /*
      EstoqueAtual não recebe CHECK >= 0 porque o escopo prevê
      uma configuração opcional para permitir saldo negativo.
      As procedures deverão aplicar essa regra e controlar
      todas as alterações de saldo.
    */

    /* 7. SERVIÇOS */

    CREATE TABLE dbo.Servicos
    (
        ServicoID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Servicos PRIMARY KEY,

        Nome NVARCHAR(150) NOT NULL,
        Descricao NVARCHAR(500) NULL,

        Preco DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Servicos_Preco DEFAULT (0),

        DuracaoEstimadaMinutos INT NULL,

        Ativo BIT NOT NULL
            CONSTRAINT DF_Servicos_Ativo DEFAULT (1),

        DataCadastro DATETIME2(0) NOT NULL
            CONSTRAINT DF_Servicos_Data DEFAULT (SYSDATETIME()),

        CONSTRAINT CK_Servicos_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0),

        CONSTRAINT CK_Servicos_Preco
            CHECK (Preco >= 0),

        CONSTRAINT CK_Servicos_Duracao
            CHECK (
                DuracaoEstimadaMinutos IS NULL
                OR DuracaoEstimadaMinutos > 0
            )
    );

    /* 8. TERMINAIS DE CAIXA */

    CREATE TABLE dbo.TerminaisCaixa
    (
        TerminalCaixaID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_TerminaisCaixa PRIMARY KEY,

        Nome NVARCHAR(100) NOT NULL
            CONSTRAINT UQ_TerminaisCaixa_Nome UNIQUE,

        Ativo BIT NOT NULL
            CONSTRAINT DF_TerminaisCaixa_Ativo DEFAULT (1),

        CONSTRAINT CK_TerminaisCaixa_Nome
            CHECK (LEN(LTRIM(RTRIM(Nome))) > 0)
    );

    /* 9. SESSÕES DE CAIXA */

    CREATE TABLE dbo.Caixas
    (
        CaixaID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Caixas PRIMARY KEY,

        TerminalCaixaID INT NOT NULL,
        UsuarioAberturaID INT NOT NULL,
        UsuarioFechamentoID INT NULL,

        DataAbertura DATETIME2(0) NOT NULL
            CONSTRAINT DF_Caixas_Abertura DEFAULT (SYSDATETIME()),

        DataFechamento DATETIME2(0) NULL,

        -- Valores referentes ao dinheiro físico.
        SaldoInicial DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Caixas_SaldoInicial DEFAULT (0),

        SaldoEsperadoFechamento DECIMAL(12,2) NULL,
        SaldoConferidoFechamento DECIMAL(12,2) NULL,

        Status VARCHAR(10) NOT NULL
            CONSTRAINT DF_Caixas_Status DEFAULT ('ABERTO'),

        CONSTRAINT FK_Caixas_Terminais
            FOREIGN KEY (TerminalCaixaID)
            REFERENCES dbo.TerminaisCaixa(TerminalCaixaID),

        CONSTRAINT FK_Caixas_UsuarioAbertura
            FOREIGN KEY (UsuarioAberturaID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT FK_Caixas_UsuarioFechamento
            FOREIGN KEY (UsuarioFechamentoID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT CK_Caixas_Status
            CHECK (Status IN ('ABERTO', 'FECHADO')),

        CONSTRAINT CK_Caixas_Saldos
            CHECK (
                SaldoInicial >= 0
                AND (
                    SaldoConferidoFechamento IS NULL
                    OR SaldoConferidoFechamento >= 0
                )
            ),

        CONSTRAINT CK_Caixas_Fechamento
            CHECK (
                (
                    Status = 'ABERTO'
                    AND DataFechamento IS NULL
                    AND UsuarioFechamentoID IS NULL
                    AND SaldoEsperadoFechamento IS NULL
                    AND SaldoConferidoFechamento IS NULL
                )
                OR
                (
                    Status = 'FECHADO'
                    AND DataFechamento IS NOT NULL
                    AND UsuarioFechamentoID IS NOT NULL
                    AND SaldoEsperadoFechamento IS NOT NULL
                    AND SaldoConferidoFechamento IS NOT NULL
                    AND DataFechamento >= DataAbertura
                )
            )
    );

    CREATE UNIQUE INDEX UX_Caixas_TerminalAberto
        ON dbo.Caixas(TerminalCaixaID)
        WHERE Status = 'ABERTO';

    /* 10. VENDAS */

    CREATE TABLE dbo.Vendas
    (
        VendaID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Vendas PRIMARY KEY,

        NumeroVenda VARCHAR(20) NOT NULL
            CONSTRAINT UQ_Vendas_Numero UNIQUE,

        ClienteID INT NULL,
        UsuarioID INT NOT NULL,
        CaixaID INT NOT NULL,

        DataVenda DATETIME2(0) NOT NULL
            CONSTRAINT DF_Vendas_Data DEFAULT (SYSDATETIME()),

        SubTotal DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Vendas_SubTotal DEFAULT (0),

        Desconto DECIMAL(12,2) NOT NULL
            CONSTRAINT DF_Vendas_Desconto DEFAULT (0),

        -- O total não pode ser digitado de forma inconsistente.
        TotalVenda AS (
            CONVERT(DECIMAL(12,2), SubTotal - Desconto)
        ) PERSISTED,

        Status VARCHAR(12) NOT NULL
            CONSTRAINT DF_Vendas_Status DEFAULT ('ABERTA'),

        UsuarioCancelamentoID INT NULL,
        DataCancelamento DATETIME2(0) NULL,
        MotivoCancelamento NVARCHAR(250) NULL,

        CONSTRAINT FK_Vendas_Clientes
            FOREIGN KEY (ClienteID)
            REFERENCES dbo.Clientes(ClienteID),

        CONSTRAINT FK_Vendas_Usuarios
            FOREIGN KEY (UsuarioID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT FK_Vendas_Caixas
            FOREIGN KEY (CaixaID)
            REFERENCES dbo.Caixas(CaixaID),

        CONSTRAINT FK_Vendas_UsuarioCancelamento
            FOREIGN KEY (UsuarioCancelamentoID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT CK_Vendas_Numero
            CHECK (LEN(LTRIM(RTRIM(NumeroVenda))) > 0),

        CONSTRAINT CK_Vendas_Valores
            CHECK (
                SubTotal >= 0
                AND Desconto >= 0
                AND Desconto <= SubTotal
            ),

        CONSTRAINT CK_Vendas_Status
            CHECK (Status IN ('ABERTA', 'FINALIZADA', 'CANCELADA')),

        CONSTRAINT CK_Vendas_Cancelamento
            CHECK (
                (
                    Status <> 'CANCELADA'
                    AND UsuarioCancelamentoID IS NULL
                    AND DataCancelamento IS NULL
                    AND MotivoCancelamento IS NULL
                )
                OR
                (
                    Status = 'CANCELADA'
                    AND UsuarioCancelamentoID IS NOT NULL
                    AND DataCancelamento IS NOT NULL
                    AND DataCancelamento >= DataVenda
                    AND MotivoCancelamento IS NOT NULL
                    AND LEN(LTRIM(RTRIM(MotivoCancelamento))) > 0
                )
            )
    );

    /* 11. ITENS DA VENDA */

    CREATE TABLE dbo.ItensVenda
    (
        ItemVendaID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_ItensVenda PRIMARY KEY,

        VendaID INT NOT NULL,
        TipoItem CHAR(1) NOT NULL,
        ProdutoID INT NULL,
        ServicoID INT NULL,

        DescricaoItem NVARCHAR(150) NOT NULL,
        Quantidade INT NOT NULL,
        PrecoUnitario DECIMAL(12,2) NOT NULL,

        /*
          Produto: custo histórico obrigatório.
          Serviço: NULL significa custo não apurado.
          Não interpretar NULL como lucro integral.
        */
        CustoUnitario DECIMAL(12,2) NULL,

        SubTotal AS (
            CONVERT(DECIMAL(12,2), Quantidade * PrecoUnitario)
        ) PERSISTED,

        CONSTRAINT FK_ItensVenda_Vendas
            FOREIGN KEY (VendaID)
            REFERENCES dbo.Vendas(VendaID),

        CONSTRAINT FK_ItensVenda_Produtos
            FOREIGN KEY (ProdutoID)
            REFERENCES dbo.Produtos(ProdutoID),

        CONSTRAINT FK_ItensVenda_Servicos
            FOREIGN KEY (ServicoID)
            REFERENCES dbo.Servicos(ServicoID),

        CONSTRAINT CK_ItensVenda_Tipo
            CHECK (TipoItem IN ('P', 'S')),

        CONSTRAINT CK_ItensVenda_Referencia
            CHECK (
                (
                    TipoItem = 'P'
                    AND ProdutoID IS NOT NULL
                    AND ServicoID IS NULL
                )
                OR
                (
                    TipoItem = 'S'
                    AND ServicoID IS NOT NULL
                    AND ProdutoID IS NULL
                )
            ),

        CONSTRAINT CK_ItensVenda_Descricao
            CHECK (LEN(LTRIM(RTRIM(DescricaoItem))) > 0),

        CONSTRAINT CK_ItensVenda_Quantidade
            CHECK (Quantidade > 0),

        CONSTRAINT CK_ItensVenda_Preco
            CHECK (PrecoUnitario >= 0),

        CONSTRAINT CK_ItensVenda_Custo
            CHECK (
                (CustoUnitario IS NULL OR CustoUnitario >= 0)
                AND (
                    TipoItem <> 'P'
                    OR CustoUnitario IS NOT NULL
                )
            )
    );

    /* 12. PAGAMENTOS */

    CREATE TABLE dbo.PagamentosVenda
    (
        PagamentoID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_PagamentosVenda PRIMARY KEY,

        VendaID INT NOT NULL,
        FormaPagamento VARCHAR(10) NOT NULL,

        -- Parcela efetivamente usada para pagar a venda.
        Valor DECIMAL(12,2) NOT NULL,

        -- Dinheiro entregue ou valor da operação eletrônica.
        ValorRecebido DECIMAL(12,2) NOT NULL,

        Troco AS (
            CONVERT(DECIMAL(12,2), ValorRecebido - Valor)
        ) PERSISTED,

        CONSTRAINT FK_PagamentosVenda_Vendas
            FOREIGN KEY (VendaID)
            REFERENCES dbo.Vendas(VendaID),

        CONSTRAINT CK_PagamentosVenda_Forma
            CHECK (
                FormaPagamento IN
                ('DINHEIRO', 'PIX', 'DEBITO', 'CREDITO')
            ),

        CONSTRAINT CK_PagamentosVenda_Valores
            CHECK (
                Valor > 0
                AND ValorRecebido >= Valor
            ),

        CONSTRAINT CK_PagamentosVenda_Troco
            CHECK (
                FormaPagamento = 'DINHEIRO'
                OR ValorRecebido = Valor
            )
    );

    /* 13. MOVIMENTOS DE CAIXA */

    CREATE TABLE dbo.MovimentosCaixa
    (
        MovimentoID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_MovimentosCaixa PRIMARY KEY,

        CaixaID INT NOT NULL,
        VendaID INT NULL,
        UsuarioID INT NOT NULL,

        TipoMovimento VARCHAR(12) NOT NULL,
        FormaPagamento VARCHAR(10) NOT NULL,

        /*
          Valor sempre positivo.
          VENDA/SUPRIMENTO entram.
          SANGRIA/ESTORNO saem.
        */
        Valor DECIMAL(12,2) NOT NULL,

        Descricao NVARCHAR(250) NULL,

        DataMovimento DATETIME2(0) NOT NULL
            CONSTRAINT DF_MovimentosCaixa_Data
            DEFAULT (SYSDATETIME()),

        CONSTRAINT FK_MovimentosCaixa_Caixas
            FOREIGN KEY (CaixaID)
            REFERENCES dbo.Caixas(CaixaID),

        CONSTRAINT FK_MovimentosCaixa_Vendas
            FOREIGN KEY (VendaID)
            REFERENCES dbo.Vendas(VendaID),

        CONSTRAINT FK_MovimentosCaixa_Usuarios
            FOREIGN KEY (UsuarioID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT CK_MovimentosCaixa_Tipo
            CHECK (
                TipoMovimento IN
                ('VENDA', 'SUPRIMENTO', 'SANGRIA', 'ESTORNO')
            ),

        CONSTRAINT CK_MovimentosCaixa_Forma
            CHECK (
                FormaPagamento IN
                ('DINHEIRO', 'PIX', 'DEBITO', 'CREDITO')
            ),

        CONSTRAINT CK_MovimentosCaixa_Valor
            CHECK (Valor > 0),

        CONSTRAINT CK_MovimentosCaixa_Origem
            CHECK (
                (
                    TipoMovimento IN ('VENDA', 'ESTORNO')
                    AND VendaID IS NOT NULL
                )
                OR
                (
                    TipoMovimento IN ('SUPRIMENTO', 'SANGRIA')
                    AND VendaID IS NULL
                    AND FormaPagamento = 'DINHEIRO'
                )
            )
    );

    /* 14. MOVIMENTAÇÕES DE ESTOQUE */

    CREATE TABLE dbo.MovimentacoesEstoque
    (
        MovimentacaoID INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_MovimentacoesEstoque PRIMARY KEY,

        ProdutoID INT NOT NULL,
        VendaID INT NULL,
        UsuarioID INT NOT NULL,

        TipoMovimento CHAR(1) NOT NULL,
        Quantidade INT NOT NULL,
        Motivo NVARCHAR(250) NOT NULL,

        DataMovimento DATETIME2(0) NOT NULL
            CONSTRAINT DF_MovimentacoesEstoque_Data
            DEFAULT (SYSDATETIME()),

        CONSTRAINT FK_MovimentacoesEstoque_Produtos
            FOREIGN KEY (ProdutoID)
            REFERENCES dbo.Produtos(ProdutoID),

        CONSTRAINT FK_MovimentacoesEstoque_Vendas
            FOREIGN KEY (VendaID)
            REFERENCES dbo.Vendas(VendaID),

        CONSTRAINT FK_MovimentacoesEstoque_Usuarios
            FOREIGN KEY (UsuarioID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT CK_MovimentacoesEstoque_Tipo
            CHECK (TipoMovimento IN ('E', 'S')),

        CONSTRAINT CK_MovimentacoesEstoque_Quantidade
            CHECK (Quantidade > 0),

        CONSTRAINT CK_MovimentacoesEstoque_Motivo
            CHECK (LEN(LTRIM(RTRIM(Motivo))) > 0)
    );

    /* 15. AUDITORIA */

    CREATE TABLE dbo.Auditoria
    (
        AuditoriaID BIGINT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_Auditoria PRIMARY KEY,

        /*
          Pode ficar NULL para um evento do sistema ou
          tentativa de login sem usuário identificado.
        */
        UsuarioID INT NULL,

        Acao NVARCHAR(100) NOT NULL,
        Entidade NVARCHAR(100) NULL,
        RegistroID BIGINT NULL,
        Detalhes NVARCHAR(MAX) NULL,

        DataAcao DATETIME2(0) NOT NULL
            CONSTRAINT DF_Auditoria_Data DEFAULT (SYSDATETIME()),

        CONSTRAINT FK_Auditoria_Usuarios
            FOREIGN KEY (UsuarioID)
            REFERENCES dbo.Usuarios(UsuarioID),

        CONSTRAINT CK_Auditoria_Acao
            CHECK (LEN(LTRIM(RTRIM(Acao))) > 0)
    );

    /* 16. CONFIGURAÇÕES */

    CREATE TABLE dbo.Configuracoes
    (
        ConfiguracaoID INT NOT NULL
            CONSTRAINT PK_Configuracoes PRIMARY KEY,

        NomeEstabelecimento NVARCHAR(150) NOT NULL,

        PermitirEstoqueNegativo BIT NOT NULL
            CONSTRAINT DF_Configuracoes_Estoque DEFAULT (0),

        PastaBackup NVARCHAR(500) NULL,

        CONSTRAINT CK_Configuracoes_RegistroUnico
            CHECK (ConfiguracaoID = 1),

        CONSTRAINT CK_Configuracoes_Nome
            CHECK (
                LEN(LTRIM(RTRIM(NomeEstabelecimento))) > 0
            )
    );

    /* ÍNDICES INICIAIS */

    CREATE INDEX IX_ItensVenda_VendaID
        ON dbo.ItensVenda(VendaID);

    CREATE INDEX IX_PagamentosVenda_VendaID
        ON dbo.PagamentosVenda(VendaID);

    CREATE INDEX IX_Vendas_CaixaID
        ON dbo.Vendas(CaixaID);

    CREATE INDEX IX_Vendas_DataVenda
        ON dbo.Vendas(DataVenda);

    CREATE INDEX IX_MovimentosCaixa_CaixaData
        ON dbo.MovimentosCaixa(CaixaID, DataMovimento);

    CREATE INDEX IX_MovimentacoesEstoque_ProdutoData
        ON dbo.MovimentacoesEstoque(ProdutoID, DataMovimento);

    CREATE INDEX IX_Auditoria_DataAcao
        ON dbo.Auditoria(DataAcao);

    /* DADOS INICIAIS */

    /*
      Limites de desconto começam em zero até você definir
      a política comercial. Não é criada senha padrão.
    */
    INSERT INTO dbo.Perfis
    (
        Nome,
        PercentualMaximoDesconto,
        PodeCancelarVenda,
        PodeAjustarEstoque
    )
    VALUES
        (N'Administrador', 0, 1, 1),
        (N'Gerente',       0, 1, 1),
        (N'Vendedor',      0, 0, 0);

    INSERT INTO dbo.TerminaisCaixa (Nome)
    VALUES (N'Caixa principal');

    INSERT INTO dbo.Configuracoes
    (
        ConfiguracaoID,
        NomeEstabelecimento,
        PermitirEstoqueNegativo
    )
    VALUES
    (
        1,
        N'Meu estabelecimento',
        0
    );

    COMMIT TRANSACTION;

    PRINT N'Estrutura de SEV_DB criada com sucesso.';
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    THROW;
END CATCH;
GO