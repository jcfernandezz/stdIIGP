IF OBJECT_ID('spIfc_Nfret_gl10030','P') IS NOT NULL
DROP PROC dbo.spIfc_Nfret_gl10030
GO

CREATE PROCEDURE dbo.spIfc_Nfret_gl10030
@VENDORID                     char(15),
@DOCTYPE                      smallint,
@DOCNUMBR                     char(21),
@VCHRNMBR                     char(21),
@nfRET_plan_de_retencione     char(21) = NULL,
@nfRET_Applied                tinyint = NULL
AS

IF EXISTS (SELECT 1 FROM dbo.nfret_gl10030
WHERE VENDORID = @VENDORID
   AND DOCTYPE = @DOCTYPE
   AND DOCNUMBR = @DOCNUMBR
   AND VCHRNMBR = @VCHRNMBR
 )
BEGIN
 
UPDATE dbo.nfret_gl10030
   SET nfRET_plan_de_retencione     = @nfRET_plan_de_retencione,
       nfRET_Applied                = @nfRET_Applied
 WHERE VENDORID = @VENDORID
   AND DOCTYPE = @DOCTYPE
   AND DOCNUMBR = @DOCNUMBR
   AND VCHRNMBR = @VCHRNMBR
 
 
END
ELSE
BEGIN
 
INSERT INTO dbo.nfret_gl10030
(VENDORID,DOCTYPE,DOCNUMBR,VCHRNMBR,nfRET_plan_de_retencione,nfRET_Applied)
SELECT @VENDORID,@DOCTYPE,@DOCNUMBR,@VCHRNMBR,@nfRET_plan_de_retencione,@nfRET_Applied
 
END
 
GO
-----------------------------------------------
--SELECT *
--FROM nfret_gl10030

--SELECT *
----update a set grossincomeNumber = '0'
--FROM AWLI_PM00200 a
----where a.grossincomeNumber = ''

--SELECT *
----update a set cod_comp = '11'
--FROM AWLI_PM00400 a

--select *
--from awli_pm00200

--sp_statistics nfret_gl10030
--sp_statistics awli_pm00200
--sp_columns awli_pm00200
--sp_statistics awli_pm00400
--sp_columns awli_pm00400
