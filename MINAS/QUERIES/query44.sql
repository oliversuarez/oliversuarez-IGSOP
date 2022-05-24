declare
    @codEmpl varchar(6)= '000024', --'001281', --'000024', --'001281', -- '000024',
    @fecha  datetime ='2022-02-01',
    @cta  tinyint = 1,
    @data varchar(max) = 


-- 'OLIVER^tdHojaTiempoDetalle~CodEmpr|CodEmpl|FecHTC|HorIniHTD|HorTerHTD|CodClie|CodOTP|CodOT|CodArea|AsuHTD|TarHorHTD|TieHTD|FunHTD|TipoFijoHTD~\
-- 1|001281|2022-02-01|||001526|E|017874E|LYS|esto es una prueba|144.32|160|Dr. serpa|'


using(select 

isnull(x.fech1, HorIniHTD) HorIniHTD, isnull(x.fech2, HorTerHTD) HorTerHTD     [salida de datos 1]  ,xxxxxxxxxxxxxxxxxxxxxxxxxxxxx(sin campo por default [los pk default])
from(select row_number()over(partition by HorIniHTD order by HorIniHTD) item,   [salida de datos 2]
*from #tmp001DatoPrueba  ---- ESTA LINEA ES CONSTANTE
)a cross apply masivo.UDFPK_tdHojaTiempoDetalle(codEmpl, fecHTC,(case HorIniHTD when '1900-01-01' then 2*item-1 end))x [salida de datos 3]

)s

-- ==========================QUIEBRE AQUI ========================================================

using(select 

[salida de datos 1]                            ,CodEmpr,CodEmpl,FecHTC,CodClie,CodOTP,CodOT,CodArea,AsuHTD,TarHorHTD,TieHTD,FunHTD,TipoFijoHTD,FecMod,UsuMod,FecCre,UsuCre
[salida de datos 2]

*from #tmp001Datos

[salida de datos 3]

)s 


-- return
-- select fech1, fech2, t.*
-- from(select row_number()over(order by codempr) item,*from tdHojaTiempoDetalle
-- where codempr = 1 and codEmpl = @codEmpl and fecHTC = @fecha and estreg = 'A')t 
-- outer apply(select * from masivo.UDFPK_tdHojaTiempoDetalle(isnull(codEmpl, @codEmpl), isnull(fecHTC, @fecha), (2*item-1)))tt

-- select * from masivo.UDFPK_tdHojaTiempoDetalle(@codEmpl, @fecha,2*1-1)



-- using(

-- select [salida de datos 1],CodEmpr,CodEmpl,FecHTC,CodClie,CodOTP,CodOT,CodArea,AsuHTD,TarHorHTD,TieHTD,FunHTD,TipoFijoHTD,FecMod,UsuMod,FecCre,UsuCre
-- [salida prueba 2]*from #tmp001Datos[xxxx]

-- )s 
-- on(t.CodEmpr=s.CodEmpr and t.CodEmpl=s.CodEmpl and t.FecHTC=s.FecHTC and t.HorIniHTD=s.HorIniHTD and t.HorTerHTD=s.HorTerHTD)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 

