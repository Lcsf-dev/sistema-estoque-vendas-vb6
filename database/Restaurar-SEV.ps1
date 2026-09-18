param(
    [Parameter(Mandatory=$true)][string]$ArquivoBackup,
    [string]$Servidor='.\SQLEXPRESS',
    [string]$BancoDestino=('SEV_Recuperado_'+(Get-Date -Format 'yyyyMMdd_HHmmss')),
    [string]$OrigemEsperada='SEV_DB',
    [switch]$SubstituirPrincipal
)
$ErrorActionPreference='Stop'
if($BancoDestino -notmatch '^SEV_[A-Za-z0-9_]{1,90}$'){throw 'Nome de destino inválido.'}
if($OrigemEsperada -notmatch '^SEV_[A-Za-z0-9_]{1,90}$'){throw 'Nome de origem inválido.'}
if($SubstituirPrincipal -and ($BancoDestino -ne 'SEV_DB' -or $OrigemEsperada -ne 'SEV_DB')){throw 'Substituição permitida somente de SEV_DB por um backup de SEV_DB.'}
if($BancoDestino -eq 'SEV_DB' -and !$SubstituirPrincipal){throw 'Para substituir SEV_DB, use -SubstituirPrincipal. A operação pedirá confirmação e criará um backup preventivo.'}
if(![System.IO.Path]::IsPathRooted($ArquivoBackup)){throw 'Informe o caminho absoluto do backup, acessível ao serviço SQL Server.'}
$ArquivoBackup=[System.IO.Path]::GetFullPath($ArquivoBackup)
$conexao=New-Object System.Data.SqlClient.SqlConnection
$config=New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$config['Data Source']=$Servidor;$config['Initial Catalog']='master';$config['Integrated Security']=$true;$config['TrustServerCertificate']=$true
$conexao.ConnectionString=$config.ConnectionString
$conexao.Open()
function Literal([string]$valor){return "N'"+$valor.Replace("'","''")+"'"}
function Consultar([string]$sql){
    $comando=$conexao.CreateCommand();$comando.CommandTimeout=600;$comando.CommandText=$sql
    $dados=New-Object System.Data.DataTable
    $adaptador=New-Object System.Data.SqlClient.SqlDataAdapter $comando
    [void]$adaptador.Fill($dados)
    return ,$dados
}
function Executar([string]$sql){$comando=$conexao.CreateCommand();$comando.CommandTimeout=600;$comando.CommandText=$sql;[void]$comando.ExecuteNonQuery()}
$modoExclusivo=$false
try {
    $origem=Literal $ArquivoBackup
    $cabecalho=Consultar "RESTORE HEADERONLY FROM DISK=$origem;"
    if($cabecalho.Rows.Count -ne 1 -or $cabecalho.Rows[0].DatabaseName -ne $OrigemEsperada -or $cabecalho.Rows[0].BackupType -ne 1){throw 'Use um arquivo contendo exatamente um backup completo do banco esperado.'}
    Executar "RESTORE VERIFYONLY FROM DISK=$origem WITH CHECKSUM;"
    $arquivos=Consultar "RESTORE FILELISTONLY FROM DISK=$origem;"
    $pastas=Consultar "SELECT CONVERT(nvarchar(500),SERVERPROPERTY('InstanceDefaultDataPath')) Dados,CONVERT(nvarchar(500),SERVERPROPERTY('InstanceDefaultLogPath')) Logs,CONVERT(nvarchar(500),SERVERPROPERTY('InstanceDefaultBackupPath')) Backups;"
    $existe=(Consultar ('SELECT DB_ID('+(Literal $BancoDestino)+') AS ID;')).Rows[0].ID -isnot [DBNull]
    if($existe -and !$SubstituirPrincipal){throw 'O banco de destino já existe. Escolha outro nome; nada foi substituído.'}
    $movimentos=@();$indice=0
    foreach($arquivo in $arquivos.Rows){
        $indice++
        if($arquivo.Type -notin @('D','L')){throw 'Este utilitário suporta apenas arquivos de dados e log comuns.'}
        $pasta=if($arquivo.Type -eq 'L'){$pastas.Rows[0].Logs}else{$pastas.Rows[0].Dados}
        if([string]::IsNullOrWhiteSpace($pasta)){throw 'Pasta padrão da instância não encontrada.'}
        $extensao=if($arquivo.Type -eq 'L'){'.ldf'}else{'.mdf'}
        $destino=Join-Path $pasta ($BancoDestino+'_'+[Guid]::NewGuid().ToString('N')+'_'+$indice+$extensao)
        $movimentos+=('MOVE '+(Literal $arquivo.LogicalName)+' TO '+(Literal $destino))
    }
    if($SubstituirPrincipal){
        Write-Host 'Feche o SEV em todos os computadores. A restauração substitui os dados atuais pelos dados do backup.'
        if((Read-Host 'Digite RESTAURAR SEV_DB para confirmar') -cne 'RESTAURAR SEV_DB'){throw 'Restauração cancelada.'}
        if($existe){
            $preventivo=Join-Path $pastas.Rows[0].Backups ('SEV_DB_antes_restauracao_'+[Guid]::NewGuid().ToString('N')+'.bak')
            Executar ('BACKUP DATABASE [SEV_DB] TO DISK='+(Literal $preventivo)+' WITH COPY_ONLY,CHECKSUM;')
            Executar ('RESTORE VERIFYONLY FROM DISK='+(Literal $preventivo)+' WITH CHECKSUM;')
            Write-Host ('Backup preventivo verificado: '+$preventivo)
            Executar 'ALTER DATABASE [SEV_DB] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;'
            $modoExclusivo=$true
        }
    }
    $opcoes=$movimentos -join ','
    if($SubstituirPrincipal){$opcoes+=',REPLACE'}
    Executar ("RESTORE DATABASE [$BancoDestino] FROM DISK=$origem WITH $opcoes,RECOVERY,CHECKSUM;")
    Executar "ALTER DATABASE [$BancoDestino] SET MULTI_USER;"
    $modoExclusivo=$false
    Executar "DBCC CHECKDB ([$BancoDestino]) WITH NO_INFOMSGS;"
    Executar "USE [$BancoDestino]; IF OBJECT_ID('dbo.SessoesAplicacao','U') IS NOT NULL DELETE dbo.SessoesAplicacao;"
    Write-Output "Restauração concluída e integridade verificada: $BancoDestino"
    if(!$SubstituirPrincipal){Write-Output 'O SEV_DB original foi preservado. Este banco recuperado serve para conferir os dados antes de uma substituição.'}
} finally {
    if($modoExclusivo){try{Executar 'ALTER DATABASE [SEV_DB] SET MULTI_USER;'}catch{Write-Warning 'Confira o estado do SEV_DB no SSMS.'}}
    $conexao.Dispose()
}
