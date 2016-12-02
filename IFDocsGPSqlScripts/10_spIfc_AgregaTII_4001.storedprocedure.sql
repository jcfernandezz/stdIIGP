

IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = object_id (N'[dbo].spIfc_AgregaTII_4001') AND OBJECTPROPERTY(id, N'IsProcedure') = 1) 
	DROP PROCEDURE [dbo].spIfc_AgregaTII_4001
GO

CREATE PROCEDURE [dbo].spIfc_AgregaTII_4001 (
	@VCHRNMBR char(21), 	
	@DOCTYPE smallint,
	@TII_Num_Autorizacion char(21),
	@TII_Num_Control char(21),
	@VENDORID char(15),
	@DOCNUMBR char(21),
	@Tipo_Comprobante smallint,
	@O_iErrorState int output,			--Return value: 0 = No Errors, Any Errors > 0
	@oErrString varchar(255) output		--Return Error Code List
)
--Propósito. Agrega datos adicionales de la factura para el servicio de impuestos de Bolivia
--Utilizado por. Interfaz de integración de compras desde excel
--15/10/13 jcf Creación
--
AS

	declare @seqnumbr int, @lineas int, @negativo int, @tripcode char(31),  @iStatus smallint, @iAddCodeErrState int
	select  @O_iErrorState = 0, @oErrString = '', @iStatus = 0, @iAddCodeErrState = 0
  
	begin try
		delete from tii_4001
		where VCHRNCOR = @VCHRNMBR
		and DOCTYPE = @DOCTYPE

		insert into tii_4001 (VCHRNCOR, DOCTYPE, TII_Num_Autorizacion, TII_Num_Control, VENDORID, DOCNUMBR, tipo_comprobante, Fact_Nota, IDSUCURSAL, RCRNGTRX)
		values (@VCHRNMBR, @DOCTYPE , 	@TII_Num_Autorizacion , @TII_Num_Control , 	@VENDORID , @DOCNUMBR , @Tipo_Comprobante, 0, '01', 0)
	
		return (@O_iErrorState)
	end try
	begin catch
		select @oErrString = 'Excepción al agregar datos adicionales de la factura [spIfc_AgregaTII_4001] ' + left(error_message(), 200)
	    select @O_iErrorState = 35001 
	    exec @iStatus = taUpdateString @O_iErrorState, @oErrString, @oErrString output, @iAddCodeErrState output
	    return (@O_iErrorState)
	end catch

go
------------------------------------------------------------------------------------------------------
--TEST
--select * from POP10310

--declare	@O_iErrorState int ,	--Return value: 0 = No Errors, Any Errors > 0
--	@oErrString varchar(255)	--Return Error Code List
--exec [dbo].spIfc_AgregaTII_4001 'RCT00000000000088', @O_iErrorState output, @oErrString output
--select @O_iErrorState, @oErrString

