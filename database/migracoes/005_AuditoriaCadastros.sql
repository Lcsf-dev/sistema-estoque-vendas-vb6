USE SEV_DB;
GO
CREATE OR ALTER TRIGGER dbo.trg_Categorias_Auditoria ON dbo.Categorias
AFTER INSERT, UPDATE, DELETE AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Usuario INT=TRY_CONVERT(INT,SUBSTRING(CONTEXT_INFO(),1,4));
 IF NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE UsuarioID=@Usuario) SET @Usuario=NULL;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes)
 SELECT @Usuario,CASE WHEN d.CategoriaID IS NULL THEN N'Cadastro criado' WHEN i.CategoriaID IS NULL THEN N'Cadastro excluído' ELSE N'Cadastro alterado' END,N'Categorias',COALESCE(i.CategoriaID,d.CategoriaID),
 N'Antes: '+COALESCE((SELECT d.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')+N'; Depois: '+COALESCE((SELECT i.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')
 FROM inserted i FULL OUTER JOIN deleted d ON d.CategoriaID=i.CategoriaID;
END;
GO
CREATE OR ALTER TRIGGER dbo.trg_Clientes_Auditoria ON dbo.Clientes
AFTER INSERT, UPDATE, DELETE AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Usuario INT=TRY_CONVERT(INT,SUBSTRING(CONTEXT_INFO(),1,4));
 IF NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE UsuarioID=@Usuario) SET @Usuario=NULL;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes)
 SELECT @Usuario,CASE WHEN d.ClienteID IS NULL THEN N'Cadastro criado' WHEN i.ClienteID IS NULL THEN N'Cadastro excluído' ELSE N'Cadastro alterado' END,N'Clientes',COALESCE(i.ClienteID,d.ClienteID),
 N'Antes: '+COALESCE((SELECT d.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')+N'; Depois: '+COALESCE((SELECT i.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')
 FROM inserted i FULL OUTER JOIN deleted d ON d.ClienteID=i.ClienteID;
END;
GO
CREATE OR ALTER TRIGGER dbo.trg_Fornecedores_Auditoria ON dbo.Fornecedores
AFTER INSERT, UPDATE, DELETE AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Usuario INT=TRY_CONVERT(INT,SUBSTRING(CONTEXT_INFO(),1,4));
 IF NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE UsuarioID=@Usuario) SET @Usuario=NULL;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes)
 SELECT @Usuario,CASE WHEN d.FornecedorID IS NULL THEN N'Cadastro criado' WHEN i.FornecedorID IS NULL THEN N'Cadastro excluído' ELSE N'Cadastro alterado' END,N'Fornecedores',COALESCE(i.FornecedorID,d.FornecedorID),
 N'Antes: '+COALESCE((SELECT d.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')+N'; Depois: '+COALESCE((SELECT i.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')
 FROM inserted i FULL OUTER JOIN deleted d ON d.FornecedorID=i.FornecedorID;
END;
GO
CREATE OR ALTER TRIGGER dbo.trg_Produtos_Auditoria ON dbo.Produtos
AFTER INSERT, UPDATE, DELETE AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Usuario INT=TRY_CONVERT(INT,SUBSTRING(CONTEXT_INFO(),1,4));
 IF NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE UsuarioID=@Usuario) SET @Usuario=NULL;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes)
 SELECT @Usuario,CASE WHEN d.ProdutoID IS NULL THEN N'Cadastro criado' WHEN i.ProdutoID IS NULL THEN N'Cadastro excluído' ELSE N'Cadastro alterado' END,N'Produtos',COALESCE(i.ProdutoID,d.ProdutoID),
 N'Antes: '+COALESCE((SELECT d.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')+N'; Depois: '+COALESCE((SELECT i.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')
 FROM inserted i FULL OUTER JOIN deleted d ON d.ProdutoID=i.ProdutoID;
END;
GO
CREATE OR ALTER TRIGGER dbo.trg_Servicos_Auditoria ON dbo.Servicos
AFTER INSERT, UPDATE, DELETE AS
BEGIN
 SET NOCOUNT ON;
 DECLARE @Usuario INT=TRY_CONVERT(INT,SUBSTRING(CONTEXT_INFO(),1,4));
 IF NOT EXISTS(SELECT 1 FROM dbo.Usuarios WHERE UsuarioID=@Usuario) SET @Usuario=NULL;
 INSERT dbo.Auditoria(UsuarioID,Acao,Entidade,RegistroID,Detalhes)
 SELECT @Usuario,CASE WHEN d.ServicoID IS NULL THEN N'Cadastro criado' WHEN i.ServicoID IS NULL THEN N'Cadastro excluído' ELSE N'Cadastro alterado' END,N'Servicos',COALESCE(i.ServicoID,d.ServicoID),
 N'Antes: '+COALESCE((SELECT d.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')+N'; Depois: '+COALESCE((SELECT i.* FOR JSON PATH,WITHOUT_ARRAY_WRAPPER),N'null')
 FROM inserted i FULL OUTER JOIN deleted d ON d.ServicoID=i.ServicoID;
END;
GO
