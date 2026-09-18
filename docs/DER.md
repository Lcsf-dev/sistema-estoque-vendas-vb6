# Relacionamentos

```mermaid
erDiagram
    Perfis ||--o{ Usuarios : define
    Usuarios ||--o{ SessoesAplicacao : possui
    Usuarios ||--o{ Vendas : realiza
    Clientes o|--o{ Vendas : identifica
    Categorias o|--o{ Produtos : classifica
    Fornecedores o|--o{ Produtos : fornece
    TerminaisCaixa ||--o{ Caixas : recebe
    Usuarios ||--o{ Caixas : abre_fecha
    Caixas ||--o{ Vendas : registra
    Vendas ||--|{ ItensVenda : contem
    Produtos o|--o{ ItensVenda : produto
    Servicos o|--o{ ItensVenda : servico
    Vendas ||--o{ PagamentosVenda : recebe
    Produtos ||--o{ MovimentacoesEstoque : movimenta
    Vendas o|--o{ MovimentacoesEstoque : origina
    Caixas ||--o{ MovimentosCaixa : movimenta
    Vendas o|--o{ MovimentosCaixa : origina
    Usuarios o|--o{ Auditoria : executa
```

`ItensVenda.TipoItem` distingue P (produto) de S (serviço). Cada linha aponta para exatamente um produto ou serviço, conforme as constraints da estrutura existente. Preço, custo e descrição do item vendido representam a data da venda. Configuracoes é um registro único com código 1. Não há exclusão de cadastros pela interface; a inativação mantém os relacionamentos históricos.
