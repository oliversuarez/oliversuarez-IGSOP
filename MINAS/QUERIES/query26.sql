-- restore database dbMunizEVM from disk='E:\DB\dbMunizEVM_17.bak' with replace, move 'dbMuniz_Data' to 'E:\DB\dbmunizData.mdf', move 'dbMuniz_Log' to 'E:\DB\dbmuniz_log.ldf'
-- NOTA : realizar el CRUD  de tbArea:
-- select text from sys.syscomments where id=object_id('usptbAreaListar','p')

if exists(select 1 from sys.sysobjects where id=object_id('dbo.usp_tipoDeAreaListar','p'))
drop procedure dbo.usp_tipoDeAreaListar
go
create procedure dbo.usp_tipoDeAreaListar
as
begin
set nocount on  
set language english  

;with 
tmp001_sep(t,r,ls,pk)as(
    select*from(values('|','~','^','_'))t(sepCamp,sepReg,sepLis,sepPK)
)
,tmp001_cab(cab) as(
    select 'ID|CODIGO|CODIGOREP|AREA|TIPO|RESPONSABLE|ESTADO~50|50|60|200|80|200|100~String|String|String|String|String|String|String'
)
,tmp001_area(dato) as(
    -- select cab + (select r, t.CodEmpr, pk, t.CodArea, t, t.CodArea, t, t.CodAreaReporte, t, t.DesArea, t, tt.DesTE, t, te.NomBusEmpl, t,
    -- (case t.EstReg when 'A' then 'ACTIVO' else 'INACTIVO' end)
    select cab + (select r, t.CodEmpr, pk, t.CodArea, t, t.CodArea, t, t.CodAreaReporte, t, t.DesArea, t, t.TipArea, t, t.codEmpl, t, t.estreg
    from dbo.tbArea t 
    -- cross apply dbo.tbTipoDeEmpleado tt cross apply dbo.tbEmpleado te
    -- where t.TipArea  = tt.CodTE and t.CodEmpl=te.CodEmpl
    for xml path, type).value('.','varchar(max)')
    from tmp001_sep cross apply tmp001_cab
)
,tmp001_indices(dato) as(  
    select '^4|5|6'
)
,tmp001_tipoempleado(dato) as(  
    select stuff((select r, codTE, t, DesTE from dbo.tbTipoDeEmpleado order by DesTE
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_responsable(dato) as(  
    select stuff((select r, CodEmpl, t, NomBusEmpl from dbo.tbEmpleado order by NomBusEmpl
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_estado(dato) as(  
    select '^A|ACTIVO~I|INACTIVO'
)
select t.dato + t1.dato + t2.dato + t3.dato + t4.dato
from tmp001_area t cross apply tmp001_indices t1 cross apply tmp001_tipoempleado t2 cross apply tmp001_responsable t3 cross apply tmp001_estado t4

end
go

-- select convert(varchar(100),hashbytes('sha1','oliver suarez tinoco'),2)


go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.usptbAreaListar','p'))
drop procedure dbo.usptbAreaListar
go
create procedure dbo.usptbAreaListar
as  
begin  
set nocount on  
set language english  
  
;with tmp_area(dato) as(select  
'ID|CODIGO|CODIGOREP|AREA|TIPO|RESPONSABLE|ESTADO¬50|50|60|200|80|200|100¬String|String|String|String|String|String|String'+  
(select  
reg, t.CodEmpr, '_' ,t.CodArea , tab, t.CodArea , tab, t.CodAreaReporte , tab, t.DesArea , tab, tt.DesTE , tab, te.NomBusEmpl,tab,  
case t.EstReg when 'A' then 'ACTIVO' else 'INACTIVO' end  
from dbo.tbArea t cross apply dbo.tbTipoDeEmpleado tt   
cross apply dbo.tbEmpleado te   
cross apply (values('¬','|'))ttt(reg,tab)  
where t.TipArea  = tt.CodTE and t.CodEmpl=te.CodEmpl  
for xml path, type).value('.','varchar(max)')  
),   
  
tmp_indices(dato) as(  
 select '~4|5|6'  
),   
  
tmp_tipoempleado(dato) as(  
select '~'+  
stuff((select reg, codTE, tab, DesTE  from dbo.tbTipoDeEmpleado cross apply (values('¬','|'))ttt(reg,tab)  
order by DesTE   
for xml path, type).value('.','varchar(max)'),1,1,'')  
),  
  
tmp_responsable(dato) as(  
select '~'+  
stuff((select reg, CodEmpl , tab, NomBusEmpl   from dbo.tbEmpleado  cross apply (values('¬','|'))ttt(reg,tab)  
order by NomBusEmpl    
for xml path, type).value('.','varchar(max)'),1,1,'')  
),  
  
tmp_estado(dato) as(  
 select '~A|ACTIVO¬I|INACTIVO'  
)  
  
select (select dato from(  
select dato from tmp_area union all   
select dato from tmp_indices union all   
select dato from tmp_tipoempleado union all  
select dato from tmp_responsable union all   
select dato from tmp_estado  
)t  
for xml path, type).value('.','varchar(max)')  

end  
go


