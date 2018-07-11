IF EXISTS (
  SELECT * 
    FROM INFORMATION_SCHEMA.ROUTINES 
   WHERE SPECIFIC_SCHEMA = 'dbo'
     AND SPECIFIC_NAME = 'spPMActualizaProximoNumCheque' 
)
   DROP PROCEDURE dbo.spPMActualizaProximoNumCheque;
GO
-----------------------------------------------------------------------------
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Propósito: Actualiza el próximo número de cheque
-- 2/7/18 jcf Creación
--
-- =============================================
CREATE PROCEDURE dbo.spPMActualizaProximoNumCheque 
	@CHEKBKID varchar(15), @CURCHNUM varchar(20), @NXTCHNUM varchar(20)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from interfering with SELECT statements.
	SET NOCOUNT ON;
	declare @msj varchar(150), @rc int;
	set @rc = 1;

	BEGIN TRANSACTION 
    UPDATE dbo.CM00100 SET NXTCHNUM = @NXTCHNUM 
		WHERE NXTCHNUM = @CURCHNUM 
		AND CHEKBKID = @CHEKBKID
	set @rc = @@rowcount;
	commit transaction;
	--
	if @rc = 0 
	begin
		set @msj = 'Posible número de cheque repetido. Id. Cheque: '+ @CHEKBKID + ' Núm: ' + @CURCHNUM
		raiserror (@msj, 16, 1);
	end
END
GO
--------------------------------------------------------------------------------
IF (@@Error = 0) PRINT 'Creación exitosa de: spPMActualizaProximoNumCheque'
ELSE PRINT 'Error en la creación de: spPMActualizaProximoNumCheque'
GO

--------------------------------------------------------------------
--exec dbo.spPMActualizaProximoNumCheque 'bbva', '1', '2'
