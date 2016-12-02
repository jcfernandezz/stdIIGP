--GILA USA
--Interfaz de compras. Royalties.
--Propósito. Rol que da accesos a objetos de interfaz de compras
--Requisitos. Ejecutar en la compañía.
--21/07/16 JCF Creación
--
-----------------------------------------------------------------------------------
use gbra
GO
IF DATABASE_PRINCIPAL_ID('rol_InterfazVentas') IS NULL
	create role rol_InterfazVentas;
GO
grant select on dbo.rm00101 to rol_InterfazVentas;


go

GO
