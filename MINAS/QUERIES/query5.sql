if exists(select 1 from sys.sysobjects where id=object_id('dbo.usp_PostulantesManteGenCsv_o','p'))
drop procedure dbo.usp_PostulantesManteGenCsv_o
go
create procedure dbo.usp_PostulantesManteGenCsv_o
as
begin
set nocount on
set language english

;with
tmp001_sep(t,r,ls,e) as(
    select*from(values('|','~','^',' '))t(sepCamp,sepReg,sepList,sepEsp)
)
,tmp001_cab(cab)as (
    select 
    'ID|TIPO|NUMERO|APELLIDOS Y NOMBRES|F.NACIMIENTO|UNIDAD|EMPRESA|AREA|PUESTO|ESTADO~'+
    '10|200|150|600|80|150|100|100|300|100~'+
    'Int32|String|String|String|Date|String|String|String|String|String'
)
,tmp001_postulante(dato) as(
    select cab + (select r,
    Pos_id, t, Pos_DocTipoId, t, Pos_DocNumero, t, Pos_ApPat, e, Pos_ApMat, e, Pos_Nombres, t,
    convert(varchar, Pos_NacFecha, 103), t,
    Pos_UMId, t, Pos_EmpresaId, t, Pos_AreaId, t, Pos_PuestoPostulaId, t, Pos_EstadoId
    from Postulantes
    order by row_number()over(order by Pos_ApPat, Pos_ApMat, Pos_Nombres)
    for xml path, type).value('.','varchar(max)')
    from tmp001_sep cross apply tmp001_cab
)
,tmp001_indices(dato) as(
    select '^1,5,6,7,8,9'
)
,tmp001_tipDocum(dato) as(
    select stuff((select r, DocId_Id, t, DocId_Nombre from t_DocsIdentidad
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_undMin(dato) as(
    select stuff((select r, UM_Id, t, UM_Nombre from UMs
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_empresa(dato) as(
    select stuff((select r, Emp_Id, t, Emp_NombreCorto from Empresas
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_areas(dato) as(
    select stuff((select r, Area_Id, t, Area_Descripcion from Areas
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_puestos(dato) as(
    select stuff((select r, Pues_Id, t, Pues_Nombre from Puestos
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
,tmp001_estados(dato) as(
    select stuff((select r, PosEst_Id, t, PosEst_Nombre from PostulantesEstados
    for xml path, type).value('.','varchar(max)'),1,1,ls)
    from tmp001_sep
)
select t.dato + t1.dato + t2.dato + t3.dato + t4.dato + t5.dato + t6.dato + t7.dato
from tmp001_postulante t cross apply tmp001_indices t1 cross apply tmp001_tipDocum t2
cross apply tmp001_undMin t3 cross apply tmp001_empresa t4 cross apply tmp001_areas t5
cross apply tmp001_puestos t6 cross apply tmp001_estados t7

end
go
