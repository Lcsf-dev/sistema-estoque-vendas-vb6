param([string]$Servidor='.\SQLEXPRESS',[switch]$ConservarBanco)
$ErrorActionPreference='Stop'
$banco='SEV_TESTES_'+[Guid]::NewGuid().ToString('N')
$criado=$false
$config=New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$config['Data Source']=$Servidor;$config['Initial Catalog']='master';$config['Integrated Security']=$true;$config['TrustServerCertificate']=$true
$cn=New-Object System.Data.SqlClient.SqlConnection $config.ConnectionString
$cn.Open()
function Executar([string]$sql){$c=$cn.CreateCommand();$c.CommandTimeout=120;$c.CommandText=$sql;[void]$c.ExecuteNonQuery()}
function Escalar([string]$sql){$c=$cn.CreateCommand();$c.CommandText=$sql;return ,$c.ExecuteScalar()}
function Verificar([bool]$condicao,[string]$nome){if(!$condicao){throw "FALHOU: $nome"};Write-Output "OK: $nome"}
function Chamar([string]$nome,[hashtable]$valores){
 $c=$cn.CreateCommand();$c.CommandType='StoredProcedure';$c.CommandText='dbo.'+$nome;$c.CommandTimeout=120
 [System.Data.SqlClient.SqlCommandBuilder]::DeriveParameters($c)
 foreach($p in $c.Parameters){if($p.Direction -in @('Input','InputOutput')){$p.Value=[DBNull]::Value}}
 if($c.Parameters.Contains('@Token')){$c.Parameters['@Token'].Value=$script:token}
 foreach($k in $valores.Keys){if($null -eq $valores[$k]){$c.Parameters['@'+$k].Value=[DBNull]::Value}else{$c.Parameters['@'+$k].Value=$valores[$k]}}
 $d=New-Object System.Data.DataSet;$a=New-Object System.Data.SqlClient.SqlDataAdapter $c;[void]$a.Fill($d)
 if($c.Parameters.Contains('@ServicoID')){$script:servico=[int]$c.Parameters['@ServicoID'].Value}
 return ,$d
}
function Recusar([string]$nome,[hashtable]$valores,[int]$codigo){
 try{$null=Chamar $nome $valores;throw "Aceitou operação inválida: $nome"}
 catch{ $e=$_.Exception;while($e.InnerException){$e=$e.InnerException};if($e -isnot [System.Data.SqlClient.SqlException] -or $e.Number -ne $codigo){throw};Write-Output "OK: $nome recusou operação ($codigo)" }
}
try {
 Write-Output ("Banco temporário: "+$banco)
 & (Join-Path $PSScriptRoot '..\Instalar-Banco.ps1') -Servidor $Servidor -BancoDestino $banco
 $criado=$true
 $cn.ChangeDatabase($banco)
 $script:token=[Guid]::Empty
 $credencial=('A'*32)+':'+('B'*64)
 $null=Chamar 'usp_PrimeiroAdministrador' @{Nome='Administrador de validação';Usuario='qa_admin';Credencial=$credencial}
 $login=Chamar 'usp_AutenticarUsuario' @{Usuario='qa_admin';Prova=('B'*64)}
 $script:token=[Guid]$login.Tables[0].Rows[0].Token
 Verificar ($script:token -ne [Guid]::Empty) 'Administrador e sessão'
 Executar (Get-Content -LiteralPath (Join-Path $PSScriptRoot '002_TestarValidacoesServico.sql') -Raw -Encoding UTF8)
 Write-Output 'OK: nome, preço e duração inválidos rejeitados com sessão autenticada'
 $null=Chamar 'usp_CategoriasSalvar' @{ID=0;Nome='Categoria QA'}
 $categoria=Escalar 'SELECT MAX(CategoriaID) FROM dbo.Categorias'
 $null=Chamar 'usp_FornecedoresSalvar' @{ID=0;NomeRazao='Fornecedor QA';CNPJ='';Telefone='';Email=''}
 $fornecedor=Escalar 'SELECT MAX(FornecedorID) FROM dbo.Fornecedores'
 $null=Chamar 'usp_ClientesSalvar' @{ID=0;Nome='Cliente QA';CPF_CNPJ='';Telefone='';Email='';Endereco=''}
 $cliente=Escalar 'SELECT MAX(ClienteID) FROM dbo.Clientes'
 $null=Chamar 'usp_ProdutosSalvar' @{ID=0;Nome='Película QA';CodigoBarras='789000000001';CategoriaID=$categoria;FornecedorID=$fornecedor;PrecoCusto=8;PrecoVenda=20;EstoqueMinimo=2}
 $produto=Escalar 'SELECT MAX(ProdutoID) FROM dbo.Produtos'
 $null=Chamar 'usp_ServicoInserir' @{Nome='Aplicação QA';Descricao='Serviço de validação';Preco=15;DuracaoEstimadaMinutos=10}
 $null=Chamar 'usp_EstoqueMovimentar' @{ProdutoID=$produto;Tipo='E';Quantidade=10;Motivo='Estoque de validação'}
 $null=Chamar 'usp_AbrirCaixa' @{TerminalCaixaID=1;SaldoInicial=100}
 $caixa=Escalar 'SELECT MAX(CaixaID) FROM dbo.Caixas'
 Recusar 'usp_AbrirCaixa' @{TerminalCaixaID=1;SaldoInicial=0} 50401
 $chave=[Guid]::NewGuid()
 $venda=@{CaixaID=$caixa;ClienteID=$cliente;Desconto=0;Chave=$chave;Itens="<itens><item tipo='P' id='$produto' qtd='1'/><item tipo='S' id='$script:servico' qtd='1'/></itens>";Pagamentos="<pagamentos><pagamento forma='DINHEIRO' valor='20' recebido='50'/><pagamento forma='PIX' valor='15' recebido='15'/></pagamentos>"}
 $r=Chamar 'usp_FinalizarVenda' $venda;$id=$r.Tables[0].Rows[0].VendaID
 Verificar ((Escalar "SELECT EstoqueAtual FROM dbo.Produtos WHERE ProdutoID=$produto") -eq 9) 'Venda baixa somente produto'
 Verificar ((Escalar "SELECT TotalVenda FROM dbo.Vendas WHERE VendaID=$id") -eq 35) 'Venda mista total 35'
 Verificar ((Escalar "SELECT SUM(Troco) FROM dbo.PagamentosVenda WHERE VendaID=$id") -eq 30) 'Troco 30'
 $null=Chamar 'usp_FinalizarVenda' $venda
 Verificar ((Escalar 'SELECT COUNT(*) FROM dbo.Vendas') -eq 1) 'Repetição não duplica venda'
 $resumo=Chamar 'usp_CaixaResumo' @{CaixaID=$caixa}
 Verificar ($resumo.Tables[0].Rows[0].DinheiroEsperado -eq 120 -and $resumo.Tables[0].Rows[0].EletronicosLiquidos -eq 15) 'Caixa separa dinheiro e PIX'
 $venda.Chave=[Guid]::NewGuid();$venda.Itens="<itens><item tipo='P' id='$produto' qtd='100'/></itens>"
 Recusar 'usp_FinalizarVenda' $venda 50507
 Verificar ((Escalar 'SELECT COUNT(*) FROM dbo.Vendas') -eq 1) 'Falha de estoque não grava venda parcial'
 $venda.Itens="<itens><item tipo='S' id='$script:servico' qtd='1'/></itens>";$venda.Desconto=1
 Recusar 'usp_FinalizarVenda' $venda 50509
 $venda.Desconto=0
 Recusar 'usp_FinalizarVenda' $venda 50510
 $venda.Pagamentos="<pagamentos><pagamento forma='CREDITO' valor='15' recebido='15'/></pagamentos>"
 $null=Chamar 'usp_FinalizarVenda' $venda
 Verificar ((Escalar "SELECT EstoqueAtual FROM dbo.Produtos WHERE ProdutoID=$produto") -eq 9) 'Venda só de serviço preserva estoque'
 $null=Chamar 'usp_CancelarVenda' @{VendaID=$id;CaixaEstornoID=$caixa;Motivo='Cancelamento de validação'}
 Verificar ((Escalar "SELECT EstoqueAtual FROM dbo.Produtos WHERE ProdutoID=$produto") -eq 10) 'Cancelamento devolve produto'
 Recusar 'usp_CancelarVenda' @{VendaID=$id;CaixaEstornoID=$caixa;Motivo='Repetição'} 50522
 $null=Chamar 'usp_MovimentarCaixa' @{CaixaID=$caixa;Tipo='SUPRIMENTO';Valor=50;Motivo='Validação'}
 $null=Chamar 'usp_MovimentarCaixa' @{CaixaID=$caixa;Tipo='SANGRIA';Valor=20;Motivo='Validação'}
 Recusar 'usp_MovimentarCaixa' @{CaixaID=$caixa;Tipo='SANGRIA';Valor=999;Motivo='Validação'} 50404
 $r=Chamar 'usp_FecharCaixa' @{CaixaID=$caixa;Conferido=129}
 Verificar ($r.Tables[0].Rows[0].Esperado -eq 130 -and $r.Tables[0].Rows[0].Diferenca -eq -1) 'Fechamento calcula saldo e diferença'
 $venda.Chave=[Guid]::NewGuid()
 Recusar 'usp_FinalizarVenda' $venda 50504
 $versao=Escalar "SELECT Versao FROM dbo.Categorias WHERE CategoriaID=$categoria"
 $null=Chamar 'usp_CategoriasSalvar' @{ID=$categoria;Versao=$versao;Nome='Categoria alterada'}
 Recusar 'usp_CategoriasSalvar' @{ID=$categoria;Versao=$versao;Nome='Sobrescrita indevida'} 50100
 Verificar ((Escalar "SELECT COUNT(*) FROM dbo.Auditoria WHERE Entidade='Categorias' AND Detalhes LIKE '%Categoria QA%'") -ge 1) 'Auditoria preserva valores anteriores'
 $null=Chamar 'usp_UsuariosSalvar' @{ID=0;Nome='Vendedor QA';Usuario='qa_vendedor';PerfilID=3;Ativo=$true;Credencial=$credencial}
 $r=Chamar 'usp_AutenticarUsuario' @{Usuario='qa_vendedor';Prova=('B'*64)};$script:token=[Guid]$r.Tables[0].Rows[0].Token
 Recusar 'usp_CategoriasSalvar' @{ID=0;Nome='Não permitido'} 50201
 Recusar 'usp_EstoqueMovimentar' @{ProdutoID=$produto;Tipo='E';Quantidade=1;Motivo='Não permitido'} 50201
 $null=Chamar 'usp_Sair' @{}
 Recusar 'usp_CategoriasListar' @{Busca='';IncluirInativos=$true} 50200
 Write-Output 'INTEGRAÇÃO SQL CONCLUÍDA. Dados exclusivamente no banco de validação.'
 } finally {
 try {
  if($criado -and !$ConservarBanco){
   if($banco -notmatch '^SEV_TESTES_[a-f0-9]{32}$'){throw 'Nome inesperado: limpeza recusada.'}
   $cn.ChangeDatabase('master')
   [System.Data.SqlClient.SqlConnection]::ClearAllPools()
   Executar ("DROP DATABASE ["+$banco+"];")
   Write-Output 'Banco temporário removido.'
  }
 } finally {$cn.Dispose()}
}
