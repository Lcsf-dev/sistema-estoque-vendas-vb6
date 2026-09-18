param(
    [string]$Servidor='.\SQLEXPRESS',
    [string]$BancoDestino='SEV_DB'
)
$ErrorActionPreference='Stop'
if($BancoDestino -notmatch '^[A-Za-z][A-Za-z0-9_]{0,63}$'){throw 'Nome de banco inválido.'}
$config=New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$config['Data Source']=$Servidor;$config['Initial Catalog']='master';$config['Integrated Security']=$true;$config['TrustServerCertificate']=$true
$conexao=New-Object System.Data.SqlClient.SqlConnection $config.ConnectionString
$conexao.Open()
$transacao=$null
function ExecutarArquivo([string]$relativo){
    $sql=Get-Content -LiteralPath (Join-Path $PSScriptRoot $relativo) -Raw -Encoding UTF8
    $sql=$sql.Replace('SEV_DB',$BancoDestino)
    foreach($lote in [regex]::Split($sql,'(?im)^\s*GO\s*\r?$')){
        if([string]::IsNullOrWhiteSpace($lote)){continue}
        $comando=$conexao.CreateCommand();$comando.CommandTimeout=120;$comando.CommandText=$lote
        if($null -ne $script:transacao){$comando.Transaction=$script:transacao}
        [void]$comando.ExecuteNonQuery()
    }
    Write-Output ('Aplicado: '+$relativo)
}
try {
    $comando=$conexao.CreateCommand();$comando.CommandText='SELECT DB_ID(@Nome);'
    [void]$comando.Parameters.AddWithValue('@Nome',$BancoDestino)
    if($comando.ExecuteScalar() -isnot [DBNull]){throw 'O banco de destino já existe. Este instalador é exclusivo para banco novo; nenhum arquivo foi aplicado.'}
    ExecutarArquivo 'estrutura\001_EstruturaInicial.sql'
    $conexao.ChangeDatabase($BancoDestino)
    $script:transacao=$conexao.BeginTransaction()
    foreach($arquivo in @('003_Operacoes.sql','002_Cadastros.sql','004_IntegrarServicoExistente.sql','005_AuditoriaCadastros.sql','006_PainelRelatorios.sql')){
        ExecutarArquivo ('migracoes\'+$arquivo)
    }
    $script:transacao.Commit();$script:transacao=$null
    Write-Output "Banco $BancoDestino instalado. Não há usuários nem senhas padrão; crie o administrador no aplicativo."
} catch {
    if($null -ne $script:transacao){try{$script:transacao.Rollback()}catch{}}
    throw
} finally {$conexao.Dispose()}
