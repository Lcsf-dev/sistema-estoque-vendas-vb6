# Arquitetura e manutenção

## Continuidade do projeto

Foram preservados `modConexao.AbrirConexao`, `modServicos.InserirServico`, o formulário principal e o formulário de serviços. O cadastro de serviços ganhou pesquisa, edição e inativação. A procedure original `usp_ServicoInserir` mantém nome e parâmetros existentes e recebe um parâmetro adicional de sessão. A versão anterior é guardada em backup.

## Camadas

1. Formulários VB6 coletam e validam entradas, chamam os módulos e exibem resultados.
2. `modDados.Consultar` executa procedures com parâmetros ADO e devolve recordsets desconectados. O módulo ajusta GUIDs e o tamanho de NVARCHAR(MAX) para o driver OLE DB 19.
3. `modSeguranca` deriva senhas com PBKDF2-HMAC-SHA256, 600 mil iterações, salt aleatório de 16 bytes e saída de 32 bytes; utiliza a API BCrypt do Windows. A entrada da senha é UTF-16LE. O formato salvo é `saltHex:hashHex`.
4. SQL Server valida sessão/perfil, regras de negócio, constraints e transações. O preço enviado pela tela não determina o valor gravado: a procedure obtém o preço atual do cadastro.

Autenticação gera um token de sessão com validade de 12 horas. Cinco falhas de login bloqueiam temporariamente o usuário por cinco minutos. Fechar o aplicativo remove sua sessão. O token é mantido apenas em memória. A conexão local usa autenticação Windows e criptografia obrigatória com confiança no certificado local, conforme a configuração existente.

## Migrações

Não execute novamente `001_EstruturaInicial.sql` no banco já criado. A implantação aplica as seguintes migrações em transação, na ordem:

1. `003_Operacoes.sql`: sessões, usuários, estoque, caixa, venda, cancelamento, consultas e backup.
2. `002_Cadastros.sql`: rowversion e manutenção de cadastros.
3. `004_IntegrarServicoExistente.sql`: integração da procedure original de serviços com sessão.
4. `005_AuditoriaCadastros.sql`: triggers com antes/depois das alterações.
5. `006_PainelRelatorios.sql`: view, painel e relatório de produtos/serviços.

A ordem numérica difere da ordem de execução porque os cadastros dependem da validação de sessão criada em 003. Os arquivos de migração instalados ficam em `database\migracoes`. Scripts usam UTF-8; fontes VB6 usam Windows-1252.

## Regras centrais

- Produto movimenta estoque; serviço não. Valores e descrições dos itens são copiados para preservar o histórico.
- Finalizar venda, registrar pagamento, baixar estoque e movimentar caixa formam uma única transação.
- Produtos são bloqueados em ordem de código; uma unidade não pode ser vendida duas vezes se estoque negativo estiver desabilitado.
- Cada tentativa de venda recebe GUID; a repetição da mesma chave devolve a venda existente.
- Pagamentos precisam somar exatamente o total. Troco só existe em dinheiro. Valores financeiros usam DECIMAL(12,2) no SQL e Currency no VB6.
- Cancelamento não exclui a venda; altera status e cria movimentos inversos. Não pode ser repetido.
- Sangria e estorno em dinheiro respeitam o saldo disponível.
- Fechamento calcula dinheiro esperado, dinheiro contado e diferença. PIX e cartões são contabilizados separadamente do numerário.
- Cadastros usam rowversion para impedir que uma edição antiga sobrescreva uma nova.
- Auditoria de cadastros usa triggers na mesma transação; operações de caixa/venda/usuário têm auditoria nas procedures.

## Compilar

Abra o `.vbp`, confirme a referência **Microsoft ActiveX Data Objects 2.8 Library** e use **File → Make SistemaEstoqueVendas.exe**. O projeto é 32 bits. Requer VB6 para desenvolvimento e runtime VB6 + OLE DB 19 de 32 bits para execução. Não copie os projetos de QA para a instalação: eles usam dados fictícios e instrumentação de testes.

## Manutenção

Mantenha backups antes de mudanças de schema. Não permita atualizações diretas de saldo nas telas; use a procedure de movimentação. Não concatene entradas do usuário em SQL. O usuário Windows com privilégios de administrador do SQL não pode ser restringido pelos perfis internos do SEV; uma implantação compartilhada deve separar as permissões administrativas da conta operacional.
