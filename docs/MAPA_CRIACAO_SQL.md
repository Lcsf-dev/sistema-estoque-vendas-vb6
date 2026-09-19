# 🗄️ Mapa de criação do banco

Este documento relaciona os arquivos necessários para reconstruir o schema a partir do código-fonte. Não contém exportação de registros de usuários ou de operações comerciais.

## Ordem de instalação

| Ordem | Arquivo | Responsabilidade |
|---|---|---|
| 1 | [001_EstruturaInicial.sql](../database/estrutura/001_EstruturaInicial.sql) | CREATE DATABASE, 16 tabelas iniciais, chaves, constraints, índices e registros de configuração |
| 2 | [003_Operacoes.sql](../database/migracoes/003_Operacoes.sql) | Tabela de sessões e procedures de autenticação, estoque, caixa e venda |
| 3 | [002_Cadastros.sql](../database/migracoes/002_Cadastros.sql) | Evolução de cadastros, rowversion e procedures de manutenção |
| 4 | [004_IntegrarServicoExistente.sql](../database/migracoes/004_IntegrarServicoExistente.sql) | Procedure de serviço compatível com a sessão autenticada |
| 5 | [005_AuditoriaCadastros.sql](../database/migracoes/005_AuditoriaCadastros.sql) | Cinco triggers de auditoria |
| 6 | [006_PainelRelatorios.sql](../database/migracoes/006_PainelRelatorios.sql) | View de vendas finalizadas e procedures de indicadores/relatórios |

**003 antecede 002 por dependência de sessão**, não por ordenação numérica. Prefira o [instalador PowerShell](../database/Instalar-Banco.ps1), que codifica essa sequência. A estrutura possui sua própria transação; as migrações seguintes são agrupadas em outra transação. A criação do banco não é desfeita automaticamente caso uma etapa falhe.

## As 17 tabelas

| Domínio | Tabelas | Finalidade |
|---|---|---|
| Acesso | Perfis, Usuarios, SessoesAplicacao | Permissões, credenciais derivadas e sessão temporária |
| Cadastros | Clientes, Categorias, Fornecedores, Produtos, Servicos | Entidades comerciais e atributos atuais |
| Caixa | TerminaisCaixa, Caixas, MovimentosCaixa | Terminal, ciclos de abertura/fechamento e movimentações |
| Venda | Vendas, ItensVenda, PagamentosVenda | Cabeçalho, composição e distribuição financeira |
| Estoque | MovimentacoesEstoque | Histórico das entradas e saídas por produto |
| Administração | Configuracoes, Auditoria | Preferências operacionais e rastreabilidade |

Todas são criadas na estrutura inicial, exceto `SessoesAplicacao`, criada na migração 003. Veja os campos no [dicionário](DICIONARIO.md), os vínculos no [DER](DER.md) e as definições completas no SQL, sem depender de screenshots do SSMS.

## Integridade e índices

- PKs identificam as entidades; FKs preservam as referências entre cadastros e operações.
- CHECKs delimitam estados, valores e a associação de um item a produto ou serviço.
- Colunas calculadas representam totais e troco conforme as expressões do schema.
- Índices únicos e filtrados protegem regras específicas, além das validações nas procedures.
- Os índices da estrutura devem ser avaliados com planos de execução e volume real antes de qualquer promessa de desempenho.

O [diagnóstico somente leitura](../database/diagnostico/001_ConferirEstrutura.sql) lista objetos, CHECKs, FKs e índices. Ele não instala nem corrige estruturas.

## Dados iniciais e dados de operação

A instalação insere **três perfis, um terminal de caixa e uma configuração**. Não cadastra clientes, produtos, vendas ou usuário administrador. O administrador é criado no primeiro acesso com credencial própria. As fixtures do runner ficam restritas ao banco temporário de testes.

## Instalar não é atualizar

Os arquivos atuais compõem a instalação completa. Não existe um executor genérico de migrações pendentes com controle de versão instalado. Para evoluir uma base existente, planeje uma migração específica; não reexecute a estrutura inicial. O nome alternativo pode ser informado ao instalador, mas a conexão do VB6 precisa apontar explicitamente para ele.
