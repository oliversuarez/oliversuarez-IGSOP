declare @par_funcion varchar(100) = 'masivo.[UDFPK_demo.t_ContratosModalidades]'

select coalesce((select 1 from sys.sysobjects where id = object_id(@par_funcion, 'if')),0) okey


declare    @par_tabla varchar(100) = 'demo.t_ContratosModalidades'
    select col_name(parent_object_id, parent_column_id) campo
    from sys.default_constraints 
    where object_name(parent_object_id) = parsename(@par_tabla,1) and schema_name(schema_id) = isnull(parsename(@par_tabla,2), schema_name(schema_id)) and 
    patindex('%UDFPK_'+ @par_tabla +'%', definition) > 0


select*from(values(1))t(dd) cross apply dbo.retornaLogicaINS('demo.t_ContratosModalidades','ModCon_Id')

