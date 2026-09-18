:ON ERROR EXIT
-- Executar em modo SQLCMD, a partir da raiz do repositório.
-- Instalação NOVA: aborta se SEV_DB já contiver tabelas de usuário.
-- Não cria usuários nem senhas padrão.
:r "database\estrutura\001_EstruturaInicial.sql"
:r "database\migracoes\003_Operacoes.sql"
:r "database\migracoes\002_Cadastros.sql"
:r "database\migracoes\004_IntegrarServicoExistente.sql"
:r "database\migracoes\005_AuditoriaCadastros.sql"
:r "database\migracoes\006_PainelRelatorios.sql"
