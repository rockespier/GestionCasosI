-- RT-114: Adicionar paises Curazao, Inglaterra, Irlanda del Norte; renombrar Holanda -> Paises Bajos.
-- pais.codigo es numerico: si es IDENTITY lo genera SQL Server, si no se usa MAX(codigo)+1.
-- Holanda conserva su codigo (Caso.idPais lo referencia).
-- Ejecutar en euroamer_innovadoctors. Revisar el resultado del SELECT final.

SET XACT_ABORT ON;
BEGIN TRANSACTION;

DECLARE @esIdentity int = COLUMNPROPERTY(OBJECT_ID('pais'), 'codigo', 'IsIdentity');
DECLARE @nuevos TABLE (orden int, nombre nvarchar(100));
INSERT INTO @nuevos VALUES (1, N'Curazao'), (2, N'Inglaterra'), (3, N'Irlanda del Norte');

DECLARE @i int = 1, @nombre nvarchar(100);
WHILE @i <= 3
BEGIN
    SELECT @nombre = nombre FROM @nuevos WHERE orden = @i;

    IF NOT EXISTS (SELECT 1 FROM pais WHERE nombre = @nombre)
    BEGIN
        IF @esIdentity = 1
            INSERT INTO pais (nombre) VALUES (@nombre);
        ELSE
            INSERT INTO pais (codigo, nombre)
            SELECT ISNULL(MAX(codigo), 0) + 1, @nombre FROM pais WITH (UPDLOCK, HOLDLOCK);
    END

    SET @i = @i + 1;
END

UPDATE pais SET nombre = N'Pa' + NCHAR(237) + N'ses Bajos' WHERE nombre = N'Holanda';

COMMIT TRANSACTION;

SELECT codigo, nombre FROM pais
WHERE nombre IN (N'Curazao', N'Inglaterra', N'Irlanda del Norte', N'Pa' + NCHAR(237) + N'ses Bajos', N'Holanda')
ORDER BY codigo;
