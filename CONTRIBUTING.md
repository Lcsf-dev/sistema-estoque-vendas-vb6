# Desenvolvimento

Use branches curtas para cada mudança. Preserve Windows-1252 nos fontes VB6 e UTF-8 nos arquivos SQL/documentação. Não versione dados reais, backups, instaladores de terceiros ou credenciais.

Para alterações de banco, crie migrações explícitas, atualize o inventário/documentação e execute o runner em banco descartável. Para alterações visuais, confira a janela normal e maximizada, foco de teclado e operação dos botões; compile o projeto.

Sugestões para os próximos commits, conforme cada etapa ficar pronta:

- `feat(ui): reorganizar menu principal e painel`
- `feat(ui): adaptar formularios ao redimensionamento`
- `feat(ui): modernizar cadastro de produtos`
- `fix(pdv): corrigir alinhamento dos pagamentos`

Não crie commits artificiais para simular etapas já concluídas. O primeiro commit registra a base atual, e os próximos documentam as mudanças reais.
