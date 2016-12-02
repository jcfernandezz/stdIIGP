
-----------------------------------------------------------------------------------------
use dynamics
go

IF EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[vwIfcCompannias]') AND OBJECTPROPERTY(id,N'IsView') = 1)
    DROP view dbo.vwIfcCompannias;
GO
create view dbo.vwIfcCompannias as
--Propósito. Obtiene la lista de compañías que integran facturas de compra
--Utilizado por. Integración de facturas de compra
--14/12/10 jcf Creación
--
select ci.CMPANYID, ci.INTERID, ci.CMPNYNAM, ci.CCode
from DYNAMICS..SY01500 ci			--sy_company_mstr
WHERE INTERID in ('GUSA')				

go
IF (@@Error = 0) PRINT 'Creación exitosa de la vista: vwIfcCompannias'
ELSE PRINT 'Error en la creación de la vista: vwIfcCompannias'
GO
-----------------------------------------------------------------------------------------
--select * from vwIfcCompannias
--SELECT * FROM dynamics..vwCfdCompannias
--select *
--from sop10100
--order by sopnumbe

