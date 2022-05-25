-- select current_timestamp, sysutcdatetime(), getutcdate(), getdate()
declare @query varchar(max)

-- select concat('select*from(values(''', replace(a,' ',''','''),'''))t(a,b,c)') from(values(sysdatetimeoffset()))t(a)
-- contenido de los procedimientos
-- ===============================
select @query = (select 'select text from sys.syscomments where id=object_id(''', name,''',''p'')'
from sys.procedures where name not like '%diagram%'
for xml path, type).value('.','varchar(max)')
-- exec(@query)

-- select*from sys.tables order by 1

select*from areas