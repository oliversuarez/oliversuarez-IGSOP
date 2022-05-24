-- use dbConflictCheck
go
if exists(select 1 from sys.sysobjects where id=object_id('dbo.uspCrear_cabeceraDetalle','p'))
drop procedure dbo.uspCrear_cabeceraDetalle
go
create procedure dbo.uspCrear_cabeceraDetalle
    @data varchar(max) output
as
begin
set nocount on
set language english

declare @aud varchar(2000), @cab varchar(4000), @det varchar(max), @cabecera varchar(max), @detalle varchar(max), @dataCabecera varchar(3000)
create table #tmp001_Data(aud varchar(1000), cab varchar(4000), det varchar(max))
create table #tmp002_Data(item tinyint identity, cab varchar(max))
create table #tmp003_Data(item tinyint identity, det varchar(max))
create table #tmp001_Merge(dato varchar(max))
create table #tmp002_Merge(dato varchar(max))
select @data = 'select*from(values('''+ replace(@data, '^',''',''') +'''))t(a,b,c)'; insert into #tmp001_Data exec(@data)
select @aud = (case aud when '' then '' else aud+'^' end), @cab = cab, @detalle = det from #tmp001_Data
select @cab = @aud + @cab
exec dbo.uspMantenimiento_simple @cab output
if (substring(@cab, 1, 5) = 'ERROR')select @data = @cab
else begin
    select @cabecera = replace(replace(replace(@cab, '~U|','~'),'~I|','~'),'~D|','~'), @dataCabecera = @cab
    select @cab = 'select*from(values('''+ replace(@cab, '~','''),(''') +'''))t(a)'; insert into #tmp002_Data exec(@cab)
    select @det = 'select*from(values('''+ replace(@detalle, '~','''),(''') +'''))t(a)'; insert into #tmp003_Data exec(@det)
    select @cab = cab from #tmp002_Data where item = 2
    select @det = det from #tmp003_Data where item = 2
    select @cab = 'select*from(values('''+ replace(@cab, '|','''),(''') +'''))t(a)'; insert into #tmp001_Merge exec(@cab)
    select @det = 'select*from(values('''+ replace(@det, '|','''),(''') +'''))t(a)'; insert into #tmp002_Merge exec(@det)
    select @det = (select ',','tt.',t.dato from #tmp002_Merge t outer apply(select 1 dato from #tmp001_Merge tt where tt.dato = t.dato)tt where tt.dato is null
    for xml path, type).value('.','varchar(max)') +' into #tmp001_campos'
    exec dbo.usp_retornoPK @cabecera output
    exec dbo.usp_retornoPK @detalle output
    -- select @cabecera
    -- select @detalle
    select @cabecera = replace(@cabecera, 'select* from','cross apply'), @detalle = replace(replace(@detalle, ')t(', ')tt('), 'select*', 'select t.*'+@det)
    select @detalle += @cabecera
    -- exec(@detalle+'select*from #tmp001_campos')
    delete from #tmp001_Merge
    insert into #tmp001_Merge exec(@detalle +'select stuff((select '','',''''''|'''''','','',name '+
    'from dbo.mastertabletemp(''tempdb.dbo.#tmp001_campos'')for xml path,type).value(''.'',''varchar(max)''),1,4,''''''~'''''')')
    select @cabecera = 'select (select '+ dato +' from #tmp001_campos for xml path, type).value(''.'',''varchar(max)'')' from #tmp001_Merge
    -- select @cabecera
    delete from #tmp001_Merge
    insert into #tmp001_Merge exec(@detalle + @cabecera)
    select @cabecera = dato from #tmp001_Merge
    -- select @cabecera
    delete from #tmp001_Merge
    insert into #tmp001_Merge exec(@detalle +
    'select stuff((select ''|'',name from dbo.mastertabletemp(''tempdb.dbo.#tmp001_campos'')for xml path,type).value(''.'',''varchar(max)''),1,1,''~'')')
    select @detalle = @aud + substring(det,0,charindex('~',det)) + dato + @cabecera from #tmp001_Merge cross apply #tmp001_Data

    exec dbo.uspMantenimiento_simple @detalle output
    select @data = @dataCabecera +'^'+ @detalle

end
end
go

-- declare @x varchar(max)='OLIVER*PCWEB^tbConflicto~cf_Codigo|cf_Razon|cf_Accionista|cf_Funcionario|cf_Ruc|cf_Asunto~|SCOTIA||||^tbconflictoDetalle~cd_Codigo|cf_Codigo|cd_sede|cd_empresa|cd_tipoEmpresa|cd_direccion|cd_pais|cd_tipoCliente|cd_attache|cd_abogadoContacto|cd_estadoCliente|cd_codigoCliente|fecha_ht|fecha_factu~||LIMA|BANK OF NOVA SCOTIA|CLIENTE|207583|CANADA|EXTRAORDINARIO|||INACTIVO|207583||2012-10-08~||LIMA|CREDISCOTIA FINANCIERA S.A.|CLIENTE|20255993225|PERU|CONSTANTE|||ACTIVO|001080||2021-08-17~||LIMA|SCOTIA SOCIEDAD TITULIZADORA S.A.|CLIENTE|20426911869|PERU|EXTRAORDINARIO|||INACTIVO|206547|2013-06-28|2013-08-20~||LIMA|SCOTIABANK PERU SAA|CLIENTE|20100043140|PERU|EXTRAORDINARIO|||ACTIVO|001247||2021-08-12~||LIMA|SCOTIACAPITAL|CLIENTE|206442|PERU|EXTRAORDINARIO|||INACTIVO|206442|2012-03-13|2013-03-14~||LIMA|CREDISCOTIA FINANCIERA S.A.|EMPRESA VINCULADA|20255993225|PERU|||||||~||LIMA|SCOTIA FONDOS SOCIEDAD ADMINISTRADORA DE FONDOS S.A.|EMPRESA VINCULADA|20384959416|PERU|||||||~||LIMA|SCOTIABANK|EMPRESA VINCULADA|000000000000000|CANADA|||||||~||LIMA|SCOTIABANK & TRUST (CAYMAN) LTD.|EMPRESA VINCULADA|000000000000000|ISLAS CAIMAN|||||||~||LIMA|SCOTIABANK PERU SAA|EMPRESA VINCULADA|20100043140|PERU|||||||~||AREQUIPA|CREDISCOTIA FINANCIERA S.A.|CLIENTE|20255993225|PERÚ|EXTRAORDINARIO|||ACTIVO|000391||~||AREQUIPA|SCOTIABANK PERU SAA|CLIENTE|20100043140|PERÚ|EXTRAORDINARIO|||ACTIVO|000240||2017-06-28~||CUZCO|CREDISCOTIA FINANCIERA S.A.|CLIENTE|20255993225|PERÚ|EXTRAORDINARIO|||ACTIVO|000760||~||TRUJILLO|CREDISCOTIA FINANCIERA S.A|CLIENTE|20255993225|PERÚ|EXTRAORDINARIO|||ACTIVO|000216||~||TRUJILLO|SCOTIABANK PERU SAA|CLIENTE|20100043140|PERÚ|EXTRAORDINARIO|||ACTIVO|000343||'
-- exec dbo.uspCrear_cabeceraDetalle @x output; select @x

-- declare @y varchar(max)= 'EVERA^tbtarifa~CodEmpr|CodTari|DesTari|CodMone~1|abc28|hola prueba|2^tbtarifaDetalle~CodEmpr|CodTE|CodCate|CodTari|ImpTarDet~1|L|AA||12.34~1|L|AB||458.36~1|L|AC||1068.71'
-- exec dbo.uspCrear_cabeceraDetalle @y output; select @y
