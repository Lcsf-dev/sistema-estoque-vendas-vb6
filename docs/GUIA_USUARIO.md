# SEV — Sistema de Estoque e Vendas

Versão 1.1, continuação do projeto existente em Visual Basic 6, ADO e SQL Server Express. Banco principal: **SEV_DB**. Instância configurada: **.\SQLEXPRESS**. Aplicação local; não requer internet para funcionar.

## Começar

1. Abra `src\SistemaEstoqueVendas.vbp` no VB6 e pressione **F5**. O ponto de entrada agora é `Sub Main`, que apresenta o login.
2. No primeiro acesso, informe seu nome, um nome de usuário e uma senha de **12 a 128 caracteres**. Confirme a senha e clique em **Criar administrador**. Não existe senha padrão.
3. Em **Configurações e backup**, informe o nome do estabelecimento e os limites de desconto dos perfis. Os limites iniciais são zero. Estoque negativo fica desabilitado.
4. Cadastre categorias e fornecedores; depois produtos, serviços e clientes.
5. Em **Entrada / ajuste de estoque**, registre a quantidade inicial de cada produto, com motivo. Serviços não têm estoque.
6. Em **Abrir / movimentar / fechar caixa**, selecione o terminal principal e informe o dinheiro inicial.
7. Em **PDV**, escolha produto ou serviço, pesquise por nome ou código de barras, selecione o resultado, informe a quantidade e clique em **Adicionar**.
8. Informe os pagamentos. **Parcela da venda** é o valor aplicado ao total; **Valor recebido** é o dinheiro entregue. Exemplo: venda de R$ 35,00, recebido R$ 50,00 → troco R$ 15,00. Em PIX/cartão os dois valores devem ser iguais.
9. Clique em **Finalizar venda**. Após a confirmação, use **Nova venda**.
10. Ao encerrar o dia, confira fisicamente o dinheiro e feche o caixa. A tela informa a diferença entre o saldo esperado e o contado.

O executável é `bin\SistemaEstoqueVendas.exe`. A versão final instalada abriu normalmente e apresentou a tela **SEV - Acesso**. Você pode utilizá-la sem abrir o editor VB6.

## Funcionalidades

- Login, administrador inicial, usuários, perfis e expiração de sessão.
- Pesquisa, inclusão, alteração e ativação/inativação de categorias, clientes, fornecedores, produtos e serviços.
- Produtos com custo, preço, código de barras, estoque atual e mínimo; serviços com preço e duração opcional.
- Entrada/saída de estoque com motivo e histórico.
- Abertura e fechamento de caixa, suprimento e sangria.
- Venda de produtos, serviços ou ambos; pagamento misto, troco e desconto por perfil.
- Cancelamento com estorno financeiro e reposição apenas dos produtos.
- Histórico, relatórios por período/vendedor/itens, estoque mínimo, auditoria, painel de hoje e exportação CSV.
- Backup com checksum e verificação; utilitário separado de restauração controlada.

## Uso dos cadastros

Clique em um registro da lista para editar; **Novo** (ou **Limpar**, em serviços) inicia outra inclusão. **Ativar / desativar** preserva o histórico. Se outro operador alterar o registro, a gravação é recusada e você deve pesquisar novamente.

Nas pesquisas são mostrados até 500 cadastros. Refine o texto para encontrar outros registros. As listas auxiliares de clientes/categorias/fornecedores também usam esse limite. Consultas operacionais detalhadas mostram até 1000 linhas; os resumos e totais consideram o período completo. A exportação contém o resultado da consulta exibida. Quantidades são unidades inteiras; venda fracionada não faz parte desta versão. O carrinho comporta 200 linhas e 20 pagamentos.

O leitor de código de barras pode preencher o campo de pesquisa como um teclado. Enter pesquisa; selecione o resultado e clique em Adicionar. O sistema não cobra PIX nem cartão: registra uma operação cujo recebimento é conferido pelo operador.

## Perfis

| Perfil | Permissões |
|---|---|
| Administrador | Todos os módulos, usuários, configurações e backup |
| Gerente | Cadastros, estoque, caixa, vendas, cancelamento e consultas |
| Vendedor | Vendas, seu caixa e consultas permitidas; sem alteração de cadastros, usuários ou configurações |

O limite de desconto é configurado por perfil. Cada caixa pertence ao usuário que o abriu; gerente/administrador pode intervir. O SQL Server valida a sessão e as permissões nas procedures. Administradores do Windows/SQL Server continuam tendo acesso administrativo direto ao banco; os perfis do aplicativo não substituem essas permissões.

## Recuperação de falhas

- Após falha de conexão na finalização, mantenha o carrinho e repita **Finalizar** com os mesmos dados. A chave da operação impede gravar duas vezes a mesma tentativa.
- Se fechar o programa ou iniciar outra venda, consulte primeiro o histórico para confirmar o resultado. O carrinho não é persistido após encerrar o programa.
- Uma rejeição por preço atualizado exige conferir o cadastro e iniciar novamente o carrinho, após consultar o histórico. Os preços são recalculados no servidor.
- Para cancelar, use **Histórico**, informe o código interno da venda, escolha um caixa aberto e descreva o motivo. A devolução real de PIX/cartão é feita externamente. O caixa precisa ter dinheiro para a parte em espécie.

## Backup e restauração

Crie o backup em **Configurações e backup**. O caminho aparece na tela; a pasta padrão pertence ao serviço SQL Server. Guarde cópias em outro dispositivo com as permissões necessárias. Um backup somente no mesmo disco não protege contra falha do disco.

Veja `docs\BACKUP_RESTAURACAO.md`. A restauração normalmente cria um banco separado para conferência. Substituir o principal exige confirmação explícita e backup preventivo.

## Estrutura técnica

- `src\`: fontes VB6, codificação Windows-1252. Preserve essa codificação ao editar.
- `database\migracoes\`: atualizações incrementais para o banco existente. Ordem documentada em `docs\ARQUITETURA.md`.
- `database\Restaurar-SEV.ps1`: utilitário de restauração.
- `bin\`: executável compilado.
- `docs\`: arquitetura, regras, DER, dicionário e validação.
- `backups\`: cópia do projeto anterior à atualização, criada na instalação.

O projeto usa o driver OLE DB 19 de 32 bits (`MSOLEDBSQL19`) e a referência ADO 2.8. A conexão Windows está centralizada em `modConexao.bas`; não utiliza o login `sa` nem senha SQL no código. A validação foi feita no SQL Server 17 da instância instalada. As migrações utilizam `CREATE OR ALTER` e JSON, exigindo SQL Server 2016 SP1 ou posterior; não são compatíveis integralmente com 2014.

## Limites do escopo

Sem emissão fiscal, NFC-e, impressão, TEF, API bancária, PIX real, integração externa ou ordem de serviço. Os valores registrados dependem da conferência do operador. A versão foi validada com dados fictícios em banco separado; confira o fluxo operacional da sua loja antes de iniciar registros reais.
