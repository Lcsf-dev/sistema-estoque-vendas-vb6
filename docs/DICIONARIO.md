# Dicionário de dados

Gerado a partir dos metadados do banco de validação com todas as migrações aplicadas.

## Auditoria

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| AuditoriaID | bigint | Não | Identificador automático |
| UsuarioID | int | Sim |  |
| Acao | nvarchar(100) | Não |  |
| Entidade | nvarchar(100) | Sim |  |
| RegistroID | bigint | Sim |  |
| Detalhes | nvarchar(MAX) | Sim |  |
| DataAcao | datetime2 | Não |  |

## Caixas

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| CaixaID | int | Não | Identificador automático |
| TerminalCaixaID | int | Não |  |
| UsuarioAberturaID | int | Não |  |
| UsuarioFechamentoID | int | Sim |  |
| DataAbertura | datetime2 | Não |  |
| DataFechamento | datetime2 | Sim |  |
| SaldoInicial | decimal(12,2) | Não |  |
| SaldoEsperadoFechamento | decimal(12,2) | Sim |  |
| SaldoConferidoFechamento | decimal(12,2) | Sim |  |
| Status | varchar(10) | Não |  |

## Categorias

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| CategoriaID | int | Não | Identificador automático |
| Nome | nvarchar(100) | Não |  |
| Ativo | bit | Não |  |
| Versao | timestamp | Não | Controle de edição simultânea |

## Clientes

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| ClienteID | int | Não | Identificador automático |
| Nome | nvarchar(100) | Não |  |
| CPF_CNPJ | varchar(20) | Sim |  |
| Telefone | varchar(20) | Sim |  |
| Email | nvarchar(254) | Sim |  |
| Endereco | nvarchar(250) | Sim |  |
| DataCadastro | datetime2 | Não |  |
| Ativo | bit | Não |  |
| Versao | timestamp | Não | Controle de edição simultânea |

## Configuracoes

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| ConfiguracaoID | int | Não |  |
| NomeEstabelecimento | nvarchar(150) | Não |  |
| PermitirEstoqueNegativo | bit | Não |  |
| PastaBackup | nvarchar(500) | Sim |  |

## Fornecedores

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| FornecedorID | int | Não | Identificador automático |
| NomeRazao | nvarchar(150) | Não |  |
| CNPJ | varchar(20) | Sim |  |
| Telefone | varchar(20) | Sim |  |
| Email | nvarchar(254) | Sim |  |
| Ativo | bit | Não |  |
| Versao | timestamp | Não | Controle de edição simultânea |

## ItensVenda

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| ItemVendaID | int | Não | Identificador automático |
| VendaID | int | Não |  |
| TipoItem | char(1) | Não |  |
| ProdutoID | int | Sim |  |
| ServicoID | int | Sim |  |
| DescricaoItem | nvarchar(150) | Não |  |
| Quantidade | int | Não |  |
| PrecoUnitario | decimal(12,2) | Não |  |
| CustoUnitario | decimal(12,2) | Sim |  |
| SubTotal | decimal(12,2) | Sim | Calculado pelo banco |

## MovimentacoesEstoque

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| MovimentacaoID | int | Não | Identificador automático |
| ProdutoID | int | Não |  |
| VendaID | int | Sim |  |
| UsuarioID | int | Não |  |
| TipoMovimento | char(1) | Não |  |
| Quantidade | int | Não |  |
| Motivo | nvarchar(250) | Não |  |
| DataMovimento | datetime2 | Não |  |

## MovimentosCaixa

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| MovimentoID | int | Não | Identificador automático |
| CaixaID | int | Não |  |
| VendaID | int | Sim |  |
| UsuarioID | int | Não |  |
| TipoMovimento | varchar(12) | Não |  |
| FormaPagamento | varchar(10) | Não |  |
| Valor | decimal(12,2) | Não |  |
| Descricao | nvarchar(250) | Sim |  |
| DataMovimento | datetime2 | Não |  |

## PagamentosVenda

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| PagamentoID | int | Não | Identificador automático |
| VendaID | int | Não |  |
| FormaPagamento | varchar(10) | Não |  |
| Valor | decimal(12,2) | Não |  |
| ValorRecebido | decimal(12,2) | Não |  |
| Troco | decimal(12,2) | Sim | Calculado pelo banco |

## Perfis

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| PerfilID | int | Não | Identificador automático |
| Nome | nvarchar(50) | Não |  |
| PercentualMaximoDesconto | decimal(5,2) | Não |  |
| PodeCancelarVenda | bit | Não |  |
| PodeAjustarEstoque | bit | Não |  |
| Ativo | bit | Não |  |

## Produtos

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| ProdutoID | int | Não | Identificador automático |
| CodigoBarras | varchar(50) | Sim |  |
| Nome | nvarchar(150) | Não |  |
| CategoriaID | int | Sim |  |
| FornecedorID | int | Sim |  |
| PrecoCusto | decimal(12,2) | Não |  |
| PrecoVenda | decimal(12,2) | Não |  |
| EstoqueAtual | int | Não |  |
| EstoqueMinimo | int | Não |  |
| Ativo | bit | Não |  |
| Versao | timestamp | Não | Controle de edição simultânea |

## Servicos

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| ServicoID | int | Não | Identificador automático |
| Nome | nvarchar(150) | Não |  |
| Descricao | nvarchar(500) | Sim |  |
| Preco | decimal(12,2) | Não |  |
| DuracaoEstimadaMinutos | int | Sim |  |
| Ativo | bit | Não |  |
| DataCadastro | datetime2 | Não |  |
| Versao | timestamp | Não | Controle de edição simultânea |

## SessoesAplicacao

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| Token | uniqueidentifier | Não |  |
| UsuarioID | int | Não |  |
| ExpiraEm | datetime2 | Não |  |

## TerminaisCaixa

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| TerminalCaixaID | int | Não | Identificador automático |
| Nome | nvarchar(100) | Não |  |
| Ativo | bit | Não |  |

## Usuarios

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| UsuarioID | int | Não | Identificador automático |
| Nome | nvarchar(100) | Não |  |
| Usuario | nvarchar(50) | Não |  |
| SenhaHashCodificada | varchar(512) | Não | Salt e derivação PBKDF2; não é senha em texto puro |
| PerfilID | int | Não |  |
| Ativo | bit | Não |  |
| DataCadastro | datetime2 | Não |  |
| TentativasLogin | int | Não |  |
| BloqueadoAte | datetime2 | Sim |  |

## Vendas

| Campo | Tipo | Nulo | Observação |
|---|---|---|---|
| VendaID | int | Não | Identificador automático |
| NumeroVenda | varchar(20) | Não |  |
| ClienteID | int | Sim |  |
| UsuarioID | int | Não |  |
| CaixaID | int | Não |  |
| DataVenda | datetime2 | Não |  |
| SubTotal | decimal(12,2) | Não |  |
| Desconto | decimal(12,2) | Não |  |
| TotalVenda | decimal(12,2) | Sim | Calculado pelo banco |
| Status | varchar(12) | Não |  |
| UsuarioCancelamentoID | int | Sim |  |
| DataCancelamento | datetime2 | Sim |  |
| MotivoCancelamento | nvarchar(250) | Sim |  |
| ChaveOperacao | uniqueidentifier | Sim |  |
