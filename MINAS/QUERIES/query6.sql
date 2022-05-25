if exists(select 1 from sys.default_constraints where name = 'df_001_area')
alter table areas drop constraint df_001_area
go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.UDFPK_areas','fn'))
drop function dbo.UDFPK_areas
go
create function dbo.UDFPK_areas()
returns smallint
as
begin
declare @clave smallint = (select count(1) + 1 from areas)
return @clave
end
go
alter table areas add constraint df_001_area default(dbo.UDFPK_areas()) for area_id
go
