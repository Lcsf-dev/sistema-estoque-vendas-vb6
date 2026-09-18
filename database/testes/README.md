# Testes do banco

Execute `./database/testes/Executar-Testes.ps1` a partir do repositório. Pode informar `-Servidor '.SQLEXPRESS'`. São usados apenas dados fictícios num banco novo cujo nome começa com `SEV_TESTES_` e termina com um GUID.

O runner testa a instalação completa antes das regras. Uma falha lança uma exceção e encerra com erro. A limpeza no bloco `finally` remove exclusivamente o banco criado pelo runner. Use `-ConservarBanco` para inspecionar o resultado no SSMS.

Os valores fixos de credencial usados no teste são fixtures artificiais para exercitar as procedures, não senhas reais nem um cadastro padrão da aplicação. O teste de PBKDF2 no VB6 foi uma verificação separada, descrita em `docs/VALIDACAO.md`.

`002_TestarValidacoesServico.sql` deve ser chamado pelo runner, depois da criação da conta fictícia. O script recusa bancos fora do prefixo de testes e desfaz cada tentativa com rollback. A sessão é criada dentro da transação de cada caso.

Cobertura: serviço inválido, sessão, cadastro, rowversion, venda mista, troco, duplicação de requisição, venda sem estoque/caixa, desconto acima do limite, pagamento divergente, cancelamento, suprimento/sangria, fechamento e autorização por perfil.

Esses testes não fazem benchmark, não simulam queda de energia e não substituem validação operacional. O teste simultâneo de duas vendas da última unidade foi executado durante a implementação, conforme o registro de validação; o runner deste diretório ainda não automatiza esse cenário paralelo.
