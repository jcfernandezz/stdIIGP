SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
 
 --Propósito. Asigna el próximo número de cheque de un pago manual
 --11/7/18 jcf Creación
 go
 ALTER procedure [dbo].[taPMManualCheckPost]	@I_vBACHNUMB char(15) ,  @I_vPMNTNMBR char(20) ,  @I_vVENDORID char(15) ,  @I_vDOCNUMBR char(20) ,  @I_vDOCAMNT  numeric(19,5) ,  @I_vDOCDATE  datetime ,  @I_vPSTGDATE datetime ,  @I_vPYENTTYP smallint ,  @I_vCARDNAME char(15) ,  @I_vCURNCYID char(15) ,  
												@I_vCHEKBKID char(15) ,  @I_vTRXDSCRN char(30) ,  @I_vXCHGRATE numeric(19,7) ,  @I_vRATETPID char(15) ,  @I_vEXPNDATE datetime ,  @I_vEXCHDATE datetime ,  @I_vEXGTBDSC char(30) ,  @I_vEXTBLSRC char(50) ,  @I_vRATEEXPR smallint ,  @I_vDYSTINCR smallint ,  
												@I_vRATEVARC numeric(19,7) ,  @I_vTRXDTDEF smallint ,  @I_vRTCLCMTD smallint ,  @I_vPRVDSLMT smallint ,  @I_vDATELMTS smallint ,  @I_vTIME1 datetime ,  @I_vMDFUSRID char(15) ,  @I_vPTDUSRID char(15) ,  @I_vBatchCHEKBKID char(15) ,  @I_vCREATEDIST smallint ,  
												@I_vRequesterTrx smallint,  @I_vUSRDEFND1 char(50),      @I_vUSRDEFND2 char(50),      @I_vUSRDEFND3 char(50),      @I_vUSRDEFND4 varchar(8000),  @I_vUSRDEFND5 varchar(8000),  @O_iErrorState int output,  @oErrString varchar(255) output   
 as  
 set nocount on  
 begin try
    if @I_vPYENTTYP = 0 --cheque
	begin
		select @O_iErrorState = 0  
		declare @CURCHNUM varchar(20)

		exec dbo.spPMActualizaNumChequeDePagoManual @I_vCHEKBKID, @CURCHNUM output
 
 		update pm10400 set docnumbr =  @CURCHNUM
		where PMNTNMBR = @I_vPMNTNMBR
		and BACHNUMB = @I_vBACHNUMB;
	end
 end try
 begin catch
	 select @O_iErrorState = 16;
	 select @oErrString = ERROR_MESSAGE();
 end catch

 return (@O_iErrorState)
