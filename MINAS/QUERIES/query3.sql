-- CREATE PROCEDURE [dbo].[usp_UsuarioValidarLoginCsv]
-- 	@Login varchar(max)
-- AS
-- BEGIN
-- 	Declare @Usuario varchar(20)
-- 	Declare @Clave varchar(max)
-- 	Declare @pos int

-- 	Set @pos = CHARINDEX('|',@Login,0)
-- 	Set @Usuario = SUBSTRING(@Login,1, @pos - 1)
-- 	set @Clave = SUBSTRING(@Login, @pos + 1, len(@Login)-@Pos)

-- 	Select u.USER_ApPaterno +'|'+ u.USER_ApMaterno+'|'+ u.USER_Nombres+'|'+ d.DocId_Nombre +'|'+ u.USER_DocNumero
-- 	from Usuarios u
-- 	inner join t_DocsIdentidad d ON u.USER_DocTipo = d.DocId_Id
-- 	where u.USER_Usuario =@Usuario and u.USER_Clave = @Clave
-- END

-- select collation_name,*from sys.databases

declare @login varchar(max) = 'VARRIETA|03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4'

-- create table #tmp001_parametro(usuario varchar(80) collate database_default,clave varchar(max) collate database_default)
select top 0 cast(null as varchar(80)) usuario, cast(null as varchar(max)) clave into #tmp001_parametro
select @login = concat('select*from(values(''', replace(@login, '|', ''','''), '''))t(a,b)')
insert into #tmp001_parametro exec(@login)

;with
tmp001_sep(t,r,ls,ss) as(
    select*from(values('|','~','^',';'))t(sepCamp,sepReg,sepLista,sepAux)
)
,tmp001_parametro as(
    select*from #tmp001_parametro
)
select concat(t.user_appaterno,t, t.user_apmaterno, t, t.user_nombres, t, tt.docid_nombre, t, t.user_docnumero)
from usuarios t cross apply t_DocsIdentidad tt cross apply tmp001_parametro pp cross apply tmp001_sep
where t.user_doctipo = tt.docid_id and t.user_usuario = pp.usuario and t.user_clave = pp.clave

