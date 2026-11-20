-- Corrige incidencia.fchCreacion de los registros antiguos a hora de Lima.
-- Antes: se grababa con getdate() (hora del servidor) y la pantalla restaba 2 horas al mostrar.
-- Ahora: se graba en hora de Lima y la pantalla muestra el valor tal cual.
-- Este script resta 2 horas a las incidencias grabadas ANTES del despliegue.
--
-- El corte se hace por idIncidencia (no por fecha): cerca del despliegue las fechas
-- antiguas (hora servidor) y las nuevas (hora Lima) se solapan unas 2 horas.
--
-- PASO 1: identificar la ultima incidencia antigua. Revisar las incidencias cercanas
-- al despliegue y anotar el idIncidencia de la ultima grabada antes de desplegar:
--
--   SELECT TOP 50 idIncidencia, fchCreacion, codUsuario
--   FROM incidencia ORDER BY idIncidencia DESC;
--
-- PASO 2: poner ese id en @ultimoIdAntiguo y ejecutar todo el script.
-- Se guarda un respaldo en incidencia_fch_bkp (ver reversa al final).

SET XACT_ABORT ON;

DECLARE @ultimoIdAntiguo int = NULL;   -- <<< COMPLETAR
DECLARE @horas int = -2;

IF @ultimoIdAntiguo IS NULL
BEGIN
    RAISERROR('Completar @ultimoIdAntiguo antes de ejecutar.', 16, 1);
    RETURN;
END

IF OBJECT_ID('incidencia_fch_bkp') IS NOT NULL
BEGIN
    RAISERROR('incidencia_fch_bkp ya existe: el script ya se ejecuto. No se aplica de nuevo.', 16, 1);
    RETURN;
END

BEGIN TRANSACTION;

SELECT idIncidencia, fchCreacion
INTO incidencia_fch_bkp
FROM incidencia
WHERE idIncidencia <= @ultimoIdAntiguo
  AND fchCreacion IS NOT NULL;

UPDATE i
SET fchCreacion = DATEADD(hour, @horas, b.fchCreacion)
FROM incidencia i
JOIN incidencia_fch_bkp b ON b.idIncidencia = i.idIncidencia;

PRINT 'Incidencias corregidas: ' + CAST(@@ROWCOUNT AS varchar(20));

COMMIT TRANSACTION;

-- Verificacion: las ultimas antiguas y las primeras nuevas deben quedar en orden cronologico.
SELECT TOP 20 i.idIncidencia, b.fchCreacion AS fch_anterior, i.fchCreacion AS fch_lima
FROM incidencia i
LEFT JOIN incidencia_fch_bkp b ON b.idIncidencia = i.idIncidencia
WHERE i.idIncidencia > @ultimoIdAntiguo - 10
ORDER BY i.idIncidencia;

-- REVERSA (solo si hace falta deshacer):
--   UPDATE i SET fchCreacion = b.fchCreacion
--   FROM incidencia i JOIN incidencia_fch_bkp b ON b.idIncidencia = i.idIncidencia;
--   DROP TABLE incidencia_fch_bkp;
