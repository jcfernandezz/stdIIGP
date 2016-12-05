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
