# Banco de dados SEV

## O que está versionado

Código T-SQL e utilitários de instalação, testes e recuperação. Não são enviados `.bak`, `.mdf`, `.ldf`, contas reais nem dumps de dados. Os dados iniciais da estrutura são apenas três perfis, um terminal e uma configuração; não há usuário ou senha padrão.

## Instalação nova

Na raiz do repositório, execute `./database/Instalar-Banco.ps1`. O padrão é `.\SQLEXPRESS` e `SEV_DB`, com autenticação Windows. O instalador recusa qualquer banco de destino já existente. Pode-se informar `-Servidor` e `-BancoDestino`; ao usar outro nome, ajuste também a conexão do aplicativo.

A estrutura inicial é aplicada primeiro. As migrações são aplicadas dentro de uma transação; em caso de falha, a criação inicial do banco pode permanecer para diagnóstico, mas o instalador não anuncia sucesso. Não tenta apagar automaticamente um banco de instalação.

Alternativa com sqlcmd instalado, também na raiz do repositório:

```powershell
sqlcmd -S '.\SQLEXPRESS' -E -b -i database\InstalarBanco.sql
```

`InstalarBanco.sql` usa diretivas SQLCMD e `:ON ERROR EXIT`. Não execute como consulta T-SQL comum. Essa alternativa para no primeiro erro, mas não envolve todas as migrações em uma única transação como o instalador PowerShell.

## Ordem dos arquivos

1. `estrutura/001_EstruturaInicial.sql` — apenas para banco novo/vazio.
2. `migracoes/003_Operacoes.sql` — sessão e operações.
3. `migracoes/002_Cadastros.sql` — depende da sessão criada em 003.
4. `migracoes/004_IntegrarServicoExistente.sql`.
5. `migracoes/005_AuditoriaCadastros.sql`.
6. `migracoes/006_PainelRelatorios.sql`.

Os números 002/003 refletem a história do desenvolvimento; a ordem de instalação é a indicada acima, não a ordem alfabética. `procedures/001_usp_ServicoInserir.sql` é somente uma cópia de referência sincronizada; a instalação usa a migração 004.

## Evoluir o banco existente

Faça backup, crie uma migração nova com nome descritivo e valide-a em ambiente isolado. Não reexecute a estrutura inicial nem substitua o banco por dados de demonstração. A modernização visual da versão 1.2 não exige recriar o SEV_DB.

## Navegação

- `diagnostico/001_ConferirEstrutura.sql`: metadados, constraints, relacionamentos e índices, somente leitura.
- `consultas/001_VendasPorPeriodo.sql`: exemplo de análise de vendas, sem exportação automática.
- `testes/Executar-Testes.ps1`: instalação e validação descartáveis.
- `Restaurar-SEV.ps1`: recuperação controlada, documentada em `../docs/BACKUP_RESTAURACAO.md`.

Consulte [o inventário completo](../docs/INVENTARIO_SQL.md) para localizar cada procedure, trigger e view.
