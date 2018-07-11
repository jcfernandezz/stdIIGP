SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

 ALTER procedure [dbo].[taPMTransactionInsertPost]  @I_vBACHNUMB char(15),         @I_vVCHNUMWK char(17),           @I_vVENDORID char(15),           @I_vDOCNUMBR char(20),           @I_vDOCTYPE  smallint,           @I_vDOCAMNT  numeric(19,5),     @I_vDOCDATE  datetime,    @I_vPSTGDATE datetime,         @I_vVADCDTRO char(15),   @I_vVADDCDPR char(15),         @I_vPYMTRMID char(20),     @I_vTAXSCHID char(15),         @I_vDUEDATE  datetime,         @I_vDSCDLRAM numeric(19,5),    @I_vDISCDATE datetime,         @I_vPRCHAMNT numeric(19,5),     @I_vCHRGAMNT numeric(19,5),     @I_vCASHAMNT numeric(19,5),     @I_vCAMCBKID char(15),         @I_vCDOCNMBR char(20),         @I_vCAMTDATE datetime,         @I_vCAMPMTNM char(20),         @I_vCHEKAMNT numeric(19,5),     @I_vCHAMCBID char(15),         @I_vCHEKDATE datetime,         @I_vCAMPYNBR char(20),         @I_vCRCRDAMT numeric(19,5),     @I_vCCAMPYNM char(20),         @I_vCHEKNMBR char(20),         @I_vCARDNAME char(15),         @I_vCCRCTNUM char(20),        @I_vCRCARDDT datetime,         @I_vCHEKBKID char(15),         @I_vTRXDSCRN char(30),         @I_vTRDISAMT numeric(19,5),     @I_vTAXAMNT numeric(19,5),      @I_vFRTAMNT numeric(19,5),      @I_vTEN99AMNT numeric(19,5),    @I_vMSCCHAMT numeric(19,5),     @I_vPORDNMBR char(20),         @I_vSHIPMTHD char(15),         @I_vDISAMTAV numeric(19,5),     @I_vDISTKNAM numeric(19,5),     @I_vAPDSTKAM numeric(19,5),     @I_vMDFUSRID char(15),         @I_vPOSTEDDT datetime,         @I_vPTDUSRID char(15),         @I_vPCHSCHID char(15),         @I_vFRTSCHID char(15),         @I_vMSCSCHID char(15),         @I_vPRCTDISC numeric(19,2),  @I_vTax_Date datetime,     @I_vCURNCYID char(15),   @I_vXCHGRATE numeric(19,7),  @I_vRATETPID char(15),   @I_vEXPNDATE datetime,   @I_vEXCHDATE datetime,   @I_vEXGTBDSC char(30),   @I_vEXTBLSRC char(50),   @I_vRATEEXPR smallint,     @I_vDYSTINCR smallint,   @I_vRATEVARC numeric(19,7),  @I_vTRXDTDEF smallint,   @I_vRTCLCMTD smallint,   @I_vPRVDSLMT smallint,   @I_vDATELMTS smallint,   @I_vTIME1 datetime,   @I_vBatchCHEKBKID char(15),  @I_vCREATEDIST smallint,  @I_vRequesterTrx smallint,  @I_vUSRDEFND1 char(50),   @I_vUSRDEFND2 char(50),   @I_vUSRDEFND3 char(50),   @I_vUSRDEFND4 varchar(8000),  @I_vUSRDEFND5 varchar(8000),  
												@O_iErrorState int output,	/* <Return value:  0=No Errors, 1=Error Occurred>										*/
												@oErrString varchar(255) output	/* <Return Error Code List:>													*/
as  
 --Propósito. Ingresa la retención en la factura.
 --Requisito. El id de retención debe estar configurado en la localización argentina
 --Utilizado por. Interfaz de compras GSPN, GUSA
 --17/08/16 jcf Creación
 --
 set nocount on  
 select @O_iErrorState = 0  
 begin try

	if (isnull(@I_vUSRDEFND2, '') != '')
	begin
		if not exists(select * from nfRET_GL00060 where nfRET_plan_de_retencione = isnull(@I_vUSRDEFND2, ''))
		begin
			 select @O_iErrorState = 1;
			 select @oErrString = 'La retención '+ rtrim(@I_vUSRDEFND2) +' no existe en GP. Configure esta retención en la ventana Mantenimiento de Plan de Retenciones y vuelva a intentar. [econnect taPMTransactionInsertPost]';
		end			

		if @O_iErrorState = 0  
		begin
			if not exists(select VENDORID from nfRET_GL10030 where VENDORID = @I_vVENDORID and DOCTYPE = @I_vDOCTYPE and DOCNUMBR = @I_vDOCNUMBR and VCHRNMBR = @I_vVCHNUMWK)
				insert into nfRET_GL10030(VENDORID,DOCTYPE,DOCNUMBR,VCHRNMBR,nfRET_plan_de_retencione, nfRET_Applied)
							values(@I_vVENDORID, @I_vDOCTYPE, @I_vDOCNUMBR, @I_vVCHNUMWK, @I_vUSRDEFND2, 0 )
			else
				update nfRET_GL10030 set nfRET_plan_de_retencione = @I_vUSRDEFND2
				where VENDORID = @I_vVENDORID and DOCTYPE = @I_vDOCTYPE and DOCNUMBR = @I_vDOCNUMBR and VCHRNMBR = @I_vVCHNUMWK
		end
		--en el caso de localización USA estándar actualizar sólo cuando @I_vUSRDEFND2 > 0
		--UPDATE PM10000 SET aplywith = 1, ppstaxrt = convert(smallint, @I_vUSRDEFND2) *100
		--where BACHNUMB = @I_vBACHNUMB
		--and VCHNUMWK = @I_vVCHNUMWK
	end
 end try
 begin catch
	 select @O_iErrorState = 1;
	 select @oErrString = ERROR_MESSAGE();
 end catch

 return (@O_iErrorState)   
GO
----------------------------------------------------------------------------------
--SELECT ppstaxrt, aplywith, *
--FROM PM10000

--SELECT ppstaxrt, aplywith, *
--FROM PM20000

--SELECT ppstaxrt, aplywith, *
--FROM PM30200

--sp_columns PM10000
--sp_statistics pm10000

--select top 100 *
--from dynamics..taerrorcode
--where errorkeyfields like '%PM%'

--select VENDORID,DOCTYPE,DOCNUMBR,VCHRNMBR, 
--	nfRET_plan_de_retencione, nfRET_Applied
--from nfRET_GL10030

--SP_STATISTICS nfRET_GL10030

