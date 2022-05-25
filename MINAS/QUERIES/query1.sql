declare @data varchar(max)

select @data = (select 'select text from sys.syscomments where id=object_id(''', name, ''')' 
from sys.procedures where not name like '%diagram%'
for xml path, type).value('.','varchar(max)')
-- exec(@data)

-- go
-- use master
-- go
-- alter database db_a7c315_gesum
-- collate Modern_Spanish_CI_AS
-- go
-- select collation_name,*from sys.databases

exec sp_spaceused Postulantes

select tabla, reg
from(
select convert(varchar, t.name) tabla, rows reg, 
row_number()over(partition by t.name order by rows desc) item
from sys.tables t cross apply sys.sysindexes 
where t.object_id = id and indid in (0,1,2,3,4,5,6) and not t.name like '%diagram%'
)t where item = 1 order by 2 desc
