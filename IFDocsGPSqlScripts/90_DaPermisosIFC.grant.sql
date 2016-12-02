--GILA USA
--Interfaz de compras. Royalties.
--Propósito. Rol que da accesos a objetos de interfaz de compras
--Requisitos. Ejecutar en la compañía.
--21/07/16 JCF Creación
--
-----------------------------------------------------------------------------------
use DYNAMICS
go
IF DATABASE_PRINCIPAL_ID('rol_InterfazCompras') IS NULL
	create role rol_InterfazCompras;
GO
GRANT select ON dbo.vwIfcCompannias TO dyngrp, rol_InterfazCompras; 
go


use gspn
GO
IF DATABASE_PRINCIPAL_ID('rol_InterfazCompras') IS NULL
	create role rol_InterfazCompras;
GO

grant select, update, insert on dbo.pm00200 to rol_InterfazCompras;
grant select, update, insert on dbo.pm10000 to rol_InterfazCompras;
grant select, update, insert on dbo.pm10500 to rol_InterfazCompras;
grant select, update, insert on dbo.pm10100 to rol_InterfazCompras;
grant select on dbo.pm20000 to rol_InterfazCompras;
grant select on dbo.pm30200 to rol_InterfazCompras;

--retenciones loc Argentina
grant select, insert, update on nfRET_GL10030 to rol_InterfazCompras;
grant select on nfRET_GL00060 to rol_InterfazCompras;

--grant select on vwIfcProveedores to rol_InterfazCompras;

--datos impositivos loc Bolivia
--grant execute on spIfc_AgregaTII_4001 to rol_InterfazCompras;
go

GO
