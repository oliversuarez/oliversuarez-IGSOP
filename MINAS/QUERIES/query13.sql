
;with
tmp001_maximo as(
    select max(indid)+1 maxi from sys.sysindexes where id > 1024
)
select tabla, rows from(
select convert(varchar, t.name) tabla, rows, row_number()over(partition by t.name order by rows desc) item
from sys.tables t cross apply sys.sysindexes cross apply tmp001_maximo cross apply master.dbo.spt_values tt
where tt.type='p' and number < maxi and object_id = id and indid = number)t where item = 1 and 2=3
order by rows desc

-- select*from dbo.t_ContratosModalidades
-- select*from demo.t_ContratosModalidades

declare @datos varchar(max)
-- ='demo.t_ContratosModalidades~ModCon_Id|ModCon_Descripcion~|estudo dialog'
='MARIA^demo.t_ContratosModalidades~ModCon_Id|ModCon_Descripcion~15|LUCIANA^demo.t_ContratosModalidadesDetalle~item|ModCon_Id|detalles~||estado~||saludo~||mono'
-- exec dbo.uspMantenimiento_simple @datos output
exec dbo.uspCrear_cabeceraDetalle @datos output
select @datos salida

select*from demo.t_ContratosModalidades
select*from demo.t_ContratosModalidadesDetalle



-- declare @salida varchar(max) ='prueba1~item|dato~|patricia~|toro~|pajaro~|paloma~|aveztruz'
-- exec dbo.uspMantenimiento_simple @salida output
-- select @salida

-- alter table demo.t_ContratosModalidades add constraint df_contrato_001 default(-1) for ModCon_Disponible
-- alter table demo.t_ContratosModalidades add constraint df_contrato_002 default(dbo.[UDFPK_demo.t_ContratosModalidades]()) for ModCon_id

-- go
-- create function dbo.[UDFPK_demo.t_ContratosModalidades]()
-- returns smallint
-- as begin
-- declare @id smallint = (select isnull(max(ModCon_Id),0)+1 from demo.t_ContratosModalidades)
-- return @id
-- end
-- go

-- go
-- alter function masivo.[UDFPK_demo.t_ContratosModalidades](
--     @id smallint
-- )returns table as return(
--     select isnull(max(ModCon_Id),0)+@id Id from demo.t_ContratosModalidades
-- )
-- go



-- select u.*, p.*
-- from(
-- select row_number()over(order by modcon_id) item, *
-- from demo.t_ContratosModalidades)p cross apply masivo.[UDFPK_demo.t_ContratosModalidades](iif(p.modcon_id != 6,item,null))u


