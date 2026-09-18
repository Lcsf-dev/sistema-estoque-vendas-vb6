# Inventário SQL

Comparação das definições instaladas com os scripts versionados: **45 procedures, 5 triggers e 1 view**. Todas têm fonte correspondente. Não contém exportação de dados.

| Objeto | Tipo | Fonte |
|---|---|---|
| `dbo.usp_AbrirCaixa` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_AutenticarUsuario` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_BackupCriar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_CaixaResumo` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_CaixasListar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_CancelarVenda` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_CategoriasListar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_CategoriasObter` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_CategoriasSalvar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_CategoriasSituacao` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ClientesListar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ClientesObter` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ClientesSalvar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ClientesSituacao` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ConfiguracoesObter` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_ConfiguracoesSalvar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_ConsultaOperacional` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_EstoqueMovimentar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_FecharCaixa` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_FinalizarVenda` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_FornecedoresListar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_FornecedoresObter` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_FornecedoresSalvar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_FornecedoresSituacao` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_LoginPreparar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_LoginSalt` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_MovimentarCaixa` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_PainelResumo` | P | [006_PainelRelatorios.sql](../database/migracoes/006_PainelRelatorios.sql) |
| `dbo.usp_PerfisListar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_PrimeiroAdministrador` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_ProdutosListar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ProdutosObter` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ProdutosSalvar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ProdutosSituacao` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_RelatorioItens` | P | [006_PainelRelatorios.sql](../database/migracoes/006_PainelRelatorios.sql) |
| `dbo.usp_Sair` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_ServicoInserir` | P | [004_IntegrarServicoExistente.sql](../database/migracoes/004_IntegrarServicoExistente.sql) |
| `dbo.usp_ServicosAtualizar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ServicosListar` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ServicosObter` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_ServicosSituacao` | P | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) |
| `dbo.usp_TerminaisListar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_UsuariosListar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_UsuariosSalvar` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.usp_ValidarSessao` | P | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) |
| `dbo.trg_Categorias_Auditoria` | TR | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) |
| `dbo.trg_Clientes_Auditoria` | TR | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) |
| `dbo.trg_Fornecedores_Auditoria` | TR | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) |
| `dbo.trg_Produtos_Auditoria` | TR | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) |
| `dbo.trg_Servicos_Auditoria` | TR | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) |
| `dbo.vw_VendasFinalizadas` | V | [006_PainelRelatorios.sql](../database/migracoes/006_PainelRelatorios.sql) |
