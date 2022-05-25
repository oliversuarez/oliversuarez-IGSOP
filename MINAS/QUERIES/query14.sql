go
set nocount on
go
if exists(select 1 from sys.sysobjects where id = object_id('mastertable','if'))
drop function mastertable
go
create function mastertable(
    @par_nombreTabla varchar(max)
)returns table as return(
  select top 1000 convert(varchar, c.name) name, c.column_id, convert(varchar,type_name(c.system_type_id)) type, c.max_length,
  convert(varchar,c.collation_name) collation_name, c.is_nullable, c.is_identity, c.default_object_id, i.is_primary_key
  from sys.columns c outer apply(select coalesce((
  select i.is_primary_key from sys.index_columns ic cross apply sys.indexes i
  where  i.object_id       = ic.object_id and
         i.index_id        = ic.index_id  and
         i.is_primary_key  = 1            and
         ic.object_id      = c.object_id  and
         ic.column_id      = c.column_id), 0) is_primary_key)i
  where c.object_id = object_id(@par_nombreTabla, 'U') order by c.column_id
)
go

go
if exists(select 1 from sys.sysobjects where id = object_id('mastertabletemp','if'))
drop function mastertabletemp
go
create function mastertabletemp(
    @par_nombreTabla varchar(max)
)returns table as return(
  select top 1000 convert(varchar, c.name) name, c.column_id, convert(varchar,type_name(c.system_type_id)) type, c.max_length,
  convert(varchar,c.collation_name) collation_name, c.is_nullable, c.is_identity, c.default_object_id, i.is_primary_key
  from tempdb.sys.columns c outer apply(select coalesce((
  select i.is_primary_key from tempdb.sys.index_columns ic cross apply tempdb.sys.indexes i
  where  i.object_id       = ic.object_id and
         i.index_id        = ic.index_id  and
         i.is_primary_key  = 1            and
         ic.object_id      = c.object_id  and
         ic.column_id      = c.column_id), 0) is_primary_key)i
  where c.object_id = object_id(@par_nombreTabla, 'U') order by c.column_id
)
go

/* Nota: aqui no se le suma los miliseconds ( + datepart(ms, @fecha)) pues no se necesita */
go
if exists(select 1 from sys.sysobjects where id = object_id('dbo.datediff_big','if'))
drop function dbo.datediff_big
go
create function dbo.datediff_big(
    @par_fecha datetime
)returns table as return(
  select datediff(ss,'19700101', @par_fecha)*cast(1000 as bigint) millis
)

go
if exists(select 1 from sys.sysobjects where id = object_id('dbo.datediff2_big','if'))
drop function dbo.datediff2_big
go
create function dbo.datediff2_big(
    @par_fecha1 datetime,
    @par_fecha2 datetime
)returns table as return(
  select datediff(ss,'19700101', @par_fecha1)*cast(1000 as bigint) millis1,
         datediff(ss,'19700101', @par_fecha2)*cast(1000 as bigint) millis2
)


go
if exists(select 1 from sys.sysobjects where id = object_id('dbo.millis_date','if'))
drop function dbo.millis_date
go
create function dbo.millis_date(
  @par_millis bigint
)returns table as return(
  select dateadd(ms, @par_millis % 1000, dateadd(ss, @par_millis / 1000, '19700101')) mfecha
)

go
if exists(select 1 from sys.sysobjects where id = object_id('textoCab', 'if'))
drop function textoCab
go
create function textoCab(
  @par_cta int
)returns table as return(
  select cab = stuff((select ',a',number from master.dbo.spt_values
  where type = 'p' and number <  @par_cta
  for xml path, type).value('.', 'varchar(max)'),1,1,'')
)


go
if exists(select 1 from sys.sysobjects where id = object_id('dbo.fun_auditoria', 'if'))
drop function dbo.fun_auditoria
go
create function dbo.fun_auditoria(
    @insUpd int = -1
)returns table as return(
select top 1000 10000 + row_number()over(order by esInsert, esFecha desc) item, esInsert, campoAuditoria
from(select distinct esInsert, esFecha, campoAuditoria from(values

(1, 1, 'fechacrea'),
(1, 0, 'usucrea'),
(0, 1, 'fechamodi'),
(0, 0, 'usumodi')


)t(esInsert, esFecha, campoAuditoria) where (@insUpd = -1))t order by esInsert, esFecha desc)
go


-- =============================================================================================
-- NOTA: SON FUNCTIONS DE LOGICA DE GENERACION DE CLAVE PRIMARIA COMPUESTA SI ES QUE SE REQUIERE
-- =============================================================================================

go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.fun_retornaCampo','if'))
drop function dbo.fun_retornaCampo
go
create function dbo.fun_retornaCampo(
    @par_tabla varchar(100)
)returns table as return(
    select col_name(parent_object_id, parent_column_id) campo
    from sys.default_constraints 
    where object_name(parent_object_id) = parsename(@par_tabla,1) and schema_name(schema_id) = isnull(parsename(@par_tabla,2), schema_name(schema_id)) and
    patindex('%UDFPK_'+ @par_tabla +'%', definition) > 0
)
go

go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.validaExistsFuncion','if'))
drop function dbo.validaExistsFuncion
go
create function dbo.validaExistsFuncion(
    @par_funcion varchar(100)
)returns table as return(
  select coalesce((select 1 from sys.sysobjects where id = object_id(@par_funcion, 'if')),0) okey
)
go

go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.retornaLogicaINS','if'))
drop function dbo.retornaLogicaINS
go
create function dbo.retornaLogicaINS(
  @tabla varchar(200),
  @campo varchar(200)
)returns table as return(
select*from(values

(
'demo.t_ContratosModalidades','ModCon_Id',
'isnull(Id, ModCon_Id) ModCon_Id',
' from(select row_number()over(partition by ModCon_Id order by ModCon_Id) item,',
')p cross apply masivo.[UDFPK_demo.t_ContratosModalidades](iif(ModCon_Id = 0, item, null))'
)



)t(tablax, campox, salida1, salida2, salida3)where tablax = @tabla and campox = @campo)
go

-- =============================================================================================
-- =============================================================================================

go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.fun_validaPKs','if'))
drop function dbo.fun_validaPKs
go
create function dbo.fun_validaPKs(
    @tabla1 varchar(100),
    @tabla2 varchar(100)
)returns table as return(
    select datosPK,
    (case tb1_ident when 1 then 'set identity_insert #tmpDatoCab on;' else '' end) tb1_ident, (case tb2_ident when 1 then 'set identity_insert #tmpDatoDet on;' else '' end) tb2_ident
    from(select stuff((select ',t.', name, '=tt.', name
    from(select t.name from dbo.mastertable(@tabla1)t cross apply dbo.mastertable(@tabla2)tt
    where t.name = tt.name and t.is_primary_key = tt.is_primary_key and t.is_primary_key = 1)t
    for xml path, type).value('.','varchar(max)'),1,1,'update t set '))pk(datosPK)
    cross apply(
    select distinct tb1_ident = max(cast(t.is_identity as int))over(), tb2_ident = max(cast(tt.is_identity as int))over() 
    from dbo.mastertable(@tabla1)t cross apply dbo.mastertable(@tabla2)tt)t
)
go

-- ========================================================================================
-- NOTA: FUNCION QUE LISTA EL RETORNO DEL GENERICO DE MANTENIMIENTO uspMantenimiento_simple
-- ========================================================================================

go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.fun_lstProcedimientoXtabla','if'))
drop function dbo.fun_lstProcedimientoXtabla
go
create function dbo.fun_lstProcedimientoXtabla(
  @schema varchar(50),
  @tabla  varchar(200)
)returns table as return(
select procedimiento from(values

('dbo','tbarea','usptbAreaListar')


)t(esquema, tabla, procedimiento)where esquema = @schema and tabla = @tabla)
go


go
if exists(select 1 from sys.sysobjects where id=object_id('usp_retornoPK','p'))
drop procedure usp_retornoPK
go
create procedure usp_retornoPK
    @cadPk varchar(max) output
as
begin
set nocount on
  create table #tmp001_Data777(item tinyint identity, datos varchar(max))
  select @cadPk = 'select*from(values('''+ replace(@cadPk, '~','''),(''') +'''))t(a)';insert into #tmp001_Data777 exec(@cadPk)
  ;with
  tmp001_Data as(
      select*from #tmp001_Data777
  )
  ,tmp001_tabla(tabla) as(
      select datos from tmp001_Data where item = 1
  )
  ,tmp001_campos(campos) as(
      select replace(datos, '|', ',') from tmp001_Data where item = 2
  )
  ,tmp001_datos(datos) as(
     select datos +')t('+ campos +')' from(select 'select* from(values'+ stuff((select ',',datos from(
     select '('''+ replace(replace(datos, '|', ''','''),'~','''),(''') +''')' datos from tmp001_Data where item > 2
     )t for xml path, type).value('.','varchar(max)'),1,1,''))t(datos)cross apply tmp001_campos 
  )
  select @cadPk = datos from tmp001_datos
end
go

