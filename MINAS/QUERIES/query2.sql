-- CREATE PROCEDURE [dbo].[usp_MenuListarCsv]
-- AS
-- BEGIN
	-- select 
	-- (select STUFF((Select ';'+ convert(varchar,Menu_Id) +'|' + Menu_Nombre+ '|'+ Menu_accion + '|' + convert(varchar,Menu_PadreId)
	-- From Menu
	-- For XML PATH('')),1,1,''))
-- end
-- go
-- arquitectura CTE  -- COMMON TABLE EXPRESION


;with
tmp001_sep(t,r,ls,ss) as(
    select*from(values('|','~','^',';'))t(sepCamp,sepReg,sepLista,sepAux)
)
select stuff((select ss, menu_id, t, menu_nombre, t, menu_accion, t, menu_padreId
from dbo.menu cross apply tmp001_sep
for xml path, type).value('.','varchar(max)'),1,1,'')




select*from(values('maria',23,86.4))t(a1,a2,a3)

select*from(values('maria'),('23'),('86.4'))t(a1)

