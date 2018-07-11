IF EXISTS (
  SELECT * 
    FROM INFORMATION_SCHEMA.ROUTINES 
   WHERE SPECIFIC_SCHEMA = 'dbo'
     AND SPECIFIC_NAME = 'spPMActualizaNumChequeDePagoManual' 
)
   DROP PROCEDURE dbo.spPMActualizaNumChequeDePagoManual;
GO
-----------------------------------------------------------------------------
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- ================================================
-- Propósito: Obtiene y actualiza el próximo número de cheque
-- 2/7/18 jcf Creación
--
-- ================================================
CREATE PROCEDURE dbo.spPMActualizaNumChequeDePagoManual 
	@CHEKBKID varchar(15), @CURCHNUM varchar(20) output
AS
BEGIN
	SET NOCOUNT ON;

	declare @NXTCHNUM varchar(20), @NXTCHNUM_NUM numeric(20)
	select @CURCHNUM = NXTCHNUM
	from dbo.CM00100
	where CHEKBKID = @CHEKBKID

	if ISNUMERIC( @CURCHNUM) = 1
		select @NXTCHNUM = convert(varchar(20), convert(numeric(20), @CURCHNUM)+1);
	else
		select @NXTCHNUM = '1'

	exec dbo.spPMActualizaProximoNumCheque @CHEKBKID, @CURCHNUM, @NXTCHNUM;

END
GO
--------------------------------------------------------------------------------
IF (@@Error = 0) PRINT 'Creación exitosa de: spPMActualizaNumChequeDePagoManual'
ELSE PRINT 'Error en la creación de: spPMActualizaNumChequeDePagoManual'
GO

