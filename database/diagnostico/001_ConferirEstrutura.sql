-- Somente leitura. Execute com SEV_DB selecionado no SSMS.
SET NOCOUNT ON;
SELECT DB_NAME() AS BancoAtual;
SELECT type_desc AS Tipo,COUNT(*) AS Quantidade
FROM sys.objects WHERE is_ms_shipped=0 AND type IN ('U','P','V','TR') GROUP BY type_desc;
SELECT s.name AS Esquema,o.name AS Objeto,o.type_desc AS Tipo
FROM sys.objects o JOIN sys.schemas s ON s.schema_id=o.schema_id
WHERE o.is_ms_shipped=0 AND o.type IN ('U','P','V','TR') ORDER BY o.type_desc,o.name;
SELECT OBJECT_NAME(parent_object_id) AS Tabela,name AS Restricao,definition AS Regra
FROM sys.check_constraints ORDER BY Tabela,Restricao;
SELECT OBJECT_NAME(parent_object_id) AS Tabela,name AS Relacionamento,OBJECT_NAME(referenced_object_id) AS Referencia
FROM sys.foreign_keys ORDER BY Tabela,Relacionamento;
SELECT OBJECT_NAME(object_id) AS Tabela,name AS Indice,is_unique AS Unico,filter_definition AS Filtro
FROM sys.indexes WHERE object_id IN(SELECT object_id FROM sys.tables WHERE is_ms_shipped=0) AND name IS NOT NULL
ORDER BY Tabela,Indice;
