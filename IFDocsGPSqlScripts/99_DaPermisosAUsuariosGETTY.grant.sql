--GETTY
--Facturas AP
--Propósito. Accesos a objetos de carga de contribuidores AP
--Requisitos. Ejecutar antes los permisos 
-------------------------------------------------------------------------------------------
--Permiso a usuarios Windows:
-------------------------------------------------------------------------------------------
use gspn; 
--create user [GILA\mayra.garcia ] for login [GILA\mayra.garcia];
EXEC sp_addrolemember 'rol_InterfazCompras', 'GILA\tiiselam' ;
EXEC sp_addrolemember 'rol_InterfazCompras', 'GILA\ext-tiiselam4';
EXEC sp_addrolemember 'rol_InterfazCompras', 'Gila\consultor';
