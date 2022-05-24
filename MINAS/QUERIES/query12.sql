-- alter function dbo.udf_tabla(
-- 	@data varchar(max),
-- 	@cols tinyint
-- )returns table as return(
-- 	select concat('select*from(values(''', replace(replace(@data,'|',''','''),'~','''),('''), '''))t(',
-- 	stuff((select ',a',number from master.dbo.spt_values where type='p' and number < @cols
-- 	for xml path, type).value('.','varchar(max)'),1,1,''),')') data
-- )
-- go
declare @data varchar(max)=
'1|maria|esquivel|12~2|juana|palacios|32~3|rosa|estrada|22',
@cta tinyint = 4

select @data = data from dbo.udf_tabla(@data,@cta)
exec(@data)
