-- CREATE Procedure [dbo].[usp_PostulantesObtenerPorIdCsv]
-- 	@Data nchar(5)
-- As
-- Begin
-- 	select isnull(
-- 	(Select convert(varchar,Pos_id) + '|' + 
-- 	convert(varchar,Pos_DocTipoId) + '|' + Pos_DocNumero + '|' +
-- 	Pos_ApPat + '|' + Pos_ApMat +  '|' + Pos_Nombres +  '|' + 
-- 	IsNull(convert(varchar,Pos_NacFecha,103),'1/1/1900')+'|'+ 
-- 	convert(varchar, Pos_UMId) + '|'+
-- 	convert(varchar, Pos_EmpresaId) + '|' +
-- 	convert(varchar, Pos_AreaId  ) + '|' +
-- 	convert(varchar, Pos_PuestoPostulaId  ) + '|' +
-- 	convert(varchar, Pos_EstadoId  ) 
-- 	From Postulantes
-- 	Where Pos_Id=@Data),'')
-- End                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     

declare @data varchar(max)='105'

;with
tmp001_sep(t,r,ls,ss) as(
    select*from(values('|','~','^',';'))t(sepCamp,sepReg,sepLista,sepAux)
)
select concat(pos_id, t, pos_doctipoId, t, pos_docNumero, t, pos_appat, t, pos_apmat, t, 
pos_nombres, t, convert(varchar, pos_nacFecha, 103),
t, pos_umid, t, pos_empresaId, t, pos_areaId, t, pos_puestoPostulaId, t, pos_estadoId)
from dbo.Postulantes cross apply tmp001_sep
where pos_id = @data


