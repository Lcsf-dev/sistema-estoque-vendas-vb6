# Backup e restauração controlada

## Criar backup

Entre como administrador e abra **Configurações e backup → Criar e verificar backup**. O aplicativo executa `BACKUP DATABASE ... WITH COPY_ONLY, CHECKSUM`, verifica o arquivo com `RESTORE VERIFYONLY` e registra a operação na auditoria.

Aguarde a conclusão. O caminho completo é exibido na tela. O serviço SQL Server precisa ter acesso à pasta; o Explorer pode exigir privilégios administrativos para copiar o arquivo dessa pasta. A conta Windows que usa o aplicativo precisa de permissão de backup no banco.

## Recuperar para conferência, preservando o banco principal

Abra o PowerShell no computador do banco. Execute o script da pasta do projeto:

```powershell
& 'D:\Desenvolvimento\Sistema SEV\SistemaEstoqueVendas\database\Restaurar-SEV.ps1' `
  -ArquivoBackup 'C:\caminho\SEV_DB_arquivo.bak'
```

Substitua somente o caminho do `.bak` pelo arquivo real. O script cria um banco `SEV_Recuperado_DATA_HORA`, valida a origem, verifica o backup e executa `DBCC CHECKDB`. Não sobrescreve um banco de destino existente. No SSMS, confira as tabelas, os últimos registros e as datas. O aplicativo continua usando o SEV_DB original.

O arquivo deve estar acessível ao serviço SQL Server. Não é necessário que o PowerShell leia seu conteúdo: a leitura é feita pelo próprio SQL Server.

## Substituir o principal após conferir

Feche o SEV em todos os computadores. A substituição retorna o banco ao momento do backup e remove do banco ativo as alterações posteriores. O utilitário cria antes um backup preventivo do estado atual.

```powershell
& 'D:\Desenvolvimento\Sistema SEV\SistemaEstoqueVendas\database\Restaurar-SEV.ps1' `
  -ArquivoBackup 'C:\caminho\SEV_DB_arquivo.bak' `
  -BancoDestino SEV_DB `
  -SubstituirPrincipal
```

O script solicita que você digite **RESTAURAR SEV_DB**. Sem essa confirmação, não substitui nada. Se a criação/verificação do backup preventivo falhar, a substituição não começa. O procedimento encerra conexões do banco somente depois da confirmação e do backup preventivo.

Depois da restauração, as sessões do aplicativo são removidas; faça login novamente. As senhas e cadastros serão os existentes no backup restaurado. Guarde o caminho do backup preventivo exibido no terminal.

Se o Windows impedir a execução de scripts por política da organização, solicite a execução ao administrador responsável. Este utilitário não muda a política de execução nem as proteções do sistema.
