IF OBJECT_ID('spIfc_awli_pm00400','P') IS NOT NULL
DROP PROC dbo.spIfc_awli_pm00400
GO

--Propósito. Ingresa datos predeterminados de localización arg de la factura pm
--08/12/16 jcf Creación
--
create PROCEDURE dbo.spIfc_awli_pm00400
@VENDORID                     char(15),
@DOCTYPE                      smallint,
@DOCNUMBR                     char(21),
@VCHRNMBR                     char(21)
AS

MERGE dbo.awli_pm00400 AS Target
USING (select c.VENDORID, @VCHRNMBR vchrnmbr, @DOCTYPE doctype, @DOCNUMBR awli_docNumber, c.RESP_TYPE, c.OP_ORIGIN, c.COD_COMP, c.FORMAT_DOC, c.CHECK_CAI, c.CUIT_Pais, c.DEST_CODE, c.CUST_CODE, c.GrossIncomeStatus, c.GrossIncomeNumber
		FROM awli_pm00200 c				--configuración de localización arg del proveedor
		where c.vendorid = @VENDORID 
		) AS Source
ON (Target.vendorid = Source.vendorid
	and Target.vchrnmbr = Source.vchrnmbr
	and Target.doctype = Source.doctype
	and Target.awli_docNumber = Source. awli_docNumber)
WHEN MATCHED THEN
    UPDATE SET 	Target.COD_COMP = Source.COD_COMP,
				Target.OP_ORIGIN = Source.OP_ORIGIN,
				Target.CUST_CODE = Source.CUST_CODE ,
				Target.DEST_CODE = Source.DEST_CODE ,
				Target.FORMAT_DOC = Source.FORMAT_DOC,
				Target.RESP_TYPE = Source.RESP_TYPE 
WHEN NOT MATCHED BY TARGET THEN
    INSERT (VENDORID,VCHRNMBR,DOCTYPE,AWLI_DocNumber,COD_COMP,OP_ORIGIN,CAI, TO_DT, CUST_CODE,DEST_CODE,NRO_DESP,DIGVERIF_NRODESP,FORMAT_DOC,CURR_CODE,T_CAMBIO,CNTRLLR,RESP_TYPE)
    VALUES (Source.vendorid, @VCHRNMBR, @DOCTYPE, @DOCNUMBR, Source.COD_COMP, Source.OP_ORIGIN, '', '1/1/1900', Source.CUST_CODE, Source.DEST_CODE, '', '', Source.FORMAT_DOC, '', 0, 0, Source.RESP_TYPE);
--OUTPUT $action, Inserted.*, Deleted.*; 
 
GO
---------------------------------------------------------------------------------------
--exec dbo.spIfc_awli_pm00400 '2', 1, 'b', 'c';

--select *
--from dbo.awli_pm00400
--where vendorid = '2'

--select *
----update a set cod_comp = '11'
--FROM awli_pm00200 a
--where a.vendorid = '2'
