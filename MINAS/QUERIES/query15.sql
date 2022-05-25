if exists(select 1 from sys.sysobjects where id=object_id('dbo.uspMantenimiento_simple','p'))
drop procedure dbo.uspMantenimiento_simple
go
create procedure dbo.uspMantenimiento_simple
    @Data varchar(max) output
as
begin
set nocount on
set language english
begin try

declare @audita varchar(max), @tabla varchar(200), @metaData varchar(max), @dato varchar(max)
create table #tmp001Data(val tinyint, item smallint, esInsert tinyint, name varchar(100), column_id int, type varchar(200), max_length smallint, collation_name varchar(300),
is_nullable tinyint, is_identity tinyint, default_object_id int, is_primary_key tinyint, salida1 varchar(max), salida2 varchar(max), salida3 varchar(max))
create table #tmp001Campos(item int identity not null, campo varchar(100))

;with
tmp001_sep(t,r,ls,a) as(
  select*from(values('|','~','^','*'))t(sepCampo,sepRegistro,sepLista,sepAudita)
),
tmp001_split as(
  select audita, tabla, substring(data3,0, charindex(r,data3)) metaData, stuff(data3,1, charindex(r,data3),'') data,
  'select*from(values(''' consulta, 'cross apply(values(dateadd(hh,-5,getutcdate()),''' consultAudi
  from(select substring(data2,0,charindex(r,data2)) tabla, stuff(data2,1,charindex(r,data2),'') data3,*
  from(select substring(data,0,posAud) audita, stuff(data, 1, posAud, '') data2,*from(select charindex(ls, @Data) posAud, @Data data,*from tmp001_sep)t)t)t
),
tmp001_preData as(
  select
  consultAudi + replace(nullif(audita,''), a, ''',''') +'''))tt('+ t.cab +')'  audita,
  consultAudi + replace(nullif(audita,''), a, ''',''') +'''))ttt('+ t.cab +')' audita2, tabla,
  consulta + replace(metaData, t, '''),(''') +'''))t(a)' metaData,
  consulta + replace(replace(data, t, ''','''),r, '''),(''') + '''))t('+ tt.cab +')' data
  from tmp001_split cross apply tmp001_sep
  cross apply dbo.textoCab(len(audita)-len(replace(audita,a,''))+2)t cross apply dbo.textoCab(len(metaData)-len(replace(metaData,t,''))+1)tt
),
tmp001_auditoria as(
  select t.* from tmp001_preData cross apply dbo.fun_auditoria(case when not audita is null then -1 end)t
),
tmp001_maestra as(
  select t.*, (case s.okey when 1 then ttt.salida1 else null end) salida1, ttt.salida2, ttt.salida3
  from tmp001_preData tt cross apply dbo.mastertable(tt.tabla)t cross apply dbo.validaExistsFuncion('masivo.[UDFPK_'+ tt.tabla +']')s
  outer apply(select salida1, salida2, salida3 from(select*from dbo.fun_retornaCampo(tt.tabla)u cross apply dbo.retornaLogicaINS(tt.tabla, u.campo))i
  where i.tablax = tt.tabla and i.campox = t.name)ttt
),
tmp001_salida as(
  select null val, null item, null esInsert, ttt.* from tmp001_maestra ttt union all
  select 1 val, t.item, t.esInsert, tt.*from tmp001_auditoria t cross apply tmp001_maestra tt where t.campoAuditoria = tt.name
)
insert into #tmp001Data
select*from tmp001_salida union all
select null,null,null,null,null,null,null,null,null,null,null,null, '1', isnull(audita + audita2,''), '91919' from tmp001_preData union all
select null,null,null,null,null,null,null,null,null,null,null,null, '2', tabla, '91919' from tmp001_preData union all
select null,null,null,null,null,null,null,null,null,null,null,null, '3', metaData, '91919' from tmp001_preData union all
select null,null,null,null,null,null,null,null,null,null,null,null, '4', data, '91919' from tmp001_preData

select @audita = salida2 from #tmp001Data where salida3 = '91919' and rtrim(salida1) = '1'
select @tabla = salida2 from #tmp001Data where salida3 = '91919' and rtrim(salida1) = '2'
select @metaData = salida2 from #tmp001Data where salida3 = '91919' and rtrim(salida1) = '3'
select @dato = salida2 from #tmp001Data where salida3 = '91919' and rtrim(salida1) = '4'
insert into #tmp001Campos exec(@metaData)

;with
tmp001_matriz as(
  select tt.*from #tmp001Data tt outer apply(select*from(values('91919'))t(salida) where t.salida = tt.salida3)t where t.salida is null
),
tmp001_DataCRUD as(
  select*from(
  select t.val, tt.item, 0 esInsert, t.name, t.column_id, t.type, t.max_length, t.collation_name, t.is_nullable, t.is_identity, t.default_object_id, t.is_primary_key, t.salida1, t.salida2, t.salida3
  from tmp001_matriz t cross apply #tmp001Campos tt where t.name = tt.campo)t union all select*from tmp001_matriz where val = 1
),
tmp001_cabecera as(
  select*from(select stuff((select ',',name from tmp001_DataCRUD order by item for xml path, type).value('.','varchar(max)'),1,1,''))t(conCampo)
  cross apply(select (select (case (row_number()over(order by item)) when 1 then salida1 else '' end), ',', dato from(
  select (case when default_object_id != 0 and is_primary_key = 1 and not salida1 is null then null else name end) dato, item, max(salida1)over() salida1
  from tmp001_DataCRUD)t where not dato is null order by item for xml path, type).value('.','varchar(max)'))tt(sinCampo)
  cross apply(select stuff((select '+''|''+convert(varchar(50),isnull(inserted.'+ name + ',deleted.'+ name +'),121)'
  from tmp001_DataCRUD where is_primary_key = 1 order by item for xml path, type).value('.','varchar(max)'),1,5,'') +')t(accion,pks);')ttt(outputs)
  cross apply(select coalesce((select 'set identity_insert #tmp001Datos ON;' from tmp001_DataCRUD where is_identity = 1 for xml path, type).value('.','varchar(max)'),''))tttt(identidad)
  cross apply(select stuff((select ' and t.', name, '=s.', name from tmp001_DataCRUD where is_primary_key = 1 order by item for xml path, type).value('.','varchar(max)'),1,5,''))ttttt(pkeys)
  cross apply(select (select (case row_number()over(order by item) when 1 then max(salida2)over() end) salida2 from tmp001_DataCRUD where not salida2 is null
  for xml path, type).value('.','varchar(max)'))tttttt(salida2)
  cross apply(select (select case row_number()over(order by item) when 1 then max(salida3)over() end salida3 from tmp001_DataCRUD where not salida3 is null
  for xml path, type).value('.','varchar(max)'))ttttttt(salida3)
  cross apply(select stuff((select '|', name from tmp001_DataCRUD where is_primary_key = 1 order by item for xml path, type).value('.','varchar(max)'),1,1, @tabla +'~'))tttttttt(pkeysCadena)
),
tmp001_preMerge as(
  select bindTabla =
  'select '+ conCampo +' into #tmp001Datos from '+ @tabla +' where 1=2;'+ identidad +'insert into #tmp001Datos('+ conCampo +')'+ @dato + @audita,
  preMerge = 'insert into #tmp001Salida select left(accion,1),pks from(merge into '+ @tabla +' t using(select '+
  (case when salida2 is null then '' else sinCampo end) + isnull(salida2,'') + '*from #tmp001Datos'+ isnull(salida3,'') +')s on('+ pkeys +')'
  from tmp001_cabecera
),
tmp002_DataCRUD as(
  select*from tmp001_DataCRUD union all
  select t.val, t.item, 1 esInsert, t.name, t.column_id, t.type, t.max_length, t.collation_name, t.is_nullable, t.is_identity, t.default_object_id, t.is_primary_key, t.salida1, t.salida2, t.salida3
  from tmp001_DataCRUD t where t.val is null
),
tmp001_salida as(
  select*from(select*from(select*from tmp002_DataCRUD where not(esInsert = 0 and is_primary_key = 1))t where is_identity = 0)t
  where not(default_object_id != 0 and is_primary_key = 1 and salida1 is null)
),
tmp001_mergeData as(
  select *, max(seq)over(partition by rnk2) maxi from(
  select esInsert, pvt, item, row_number()over(order by esInsert, pvt, item) seq, rank()over(order by esInsert, pvt) rnk2, val, name
  from(select 0 pvt,*from tmp001_salida union all select 1 pvt,*from tmp001_salida)t)t where not (esInsert = 0 and pvt = 0 and isnull(val,0) = 1)
)
select @dato = bindTabla + preMerge + mergeTotal + outputs, @audita = pkeysCadena from(select(select(case
   when esInsert = 0 and pvt = 0 and seq  = rnk2 then ' when matched and s.'+ name +'=''0'''
   when esInsert = 0 and pvt = 0 and seq != rnk2 and seq != maxi then ' and s.'+ name +'=''0'''
   when esInsert = 0 and pvt = 0 and seq != rnk2 and seq  = maxi then ' and s.'+ name +'=''0'''
   when esInsert = 0 and pvt = 1 and seq  = rnk2 then ' then delete when matched then update set t.'+ name +'=s.'+ name
   when esInsert = 0 and pvt = 1 and seq != rnk2 and seq != maxi then ' ,t.'+ name +'=s.'+ name
   when esInsert = 0 and pvt = 1 and seq != rnk2 and seq  = maxi then ' ,t.'+ name +'=s.'+ name
   when esInsert = 1 and pvt = 0 and seq  = rnk2 then ' when not matched then insert('+ name
   when esInsert = 1 and pvt = 0 and seq != rnk2 and seq != maxi then ' ,'+ name
   when esInsert = 1 and pvt = 0 and seq != rnk2 and seq  = maxi then ' ,'+ name
   when esInsert = 1 and pvt = 1 and seq  = rnk2 then ')values(s.'+ name
   when esInsert = 1 and pvt = 1 and seq != rnk2 and seq != maxi then ' ,s.'+ name
   when esInsert = 1 and pvt = 1 and seq != rnk2 and seq  = maxi then ' ,s.'+ name end)
from tmp001_mergeData order by seq
for xml path, type).value('.','varchar(max)') + ')output $action,')t(mergeTotal)cross apply tmp001_preMerge cross apply tmp001_cabecera 

create table #tmp001Salida(item int identity not null, accion char(1), valorespk varchar(max))
exec(@dato)

select @metaData = null, @dato = null
select @metaData = (select '~',accion,'|',valorespk from #tmp001Salida order by item for xml path, type).value('.','varchar(max)')
select @Data = coalesce(((case when @metaData is null then '' else @audita end)+ @metaData), @tabla +'~upd')

end try
begin catch
    select @Data = 'ERROR U: '+ @tabla +': '+ error_message()
end catch
end
go
