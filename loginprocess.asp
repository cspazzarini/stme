<!-- #include FILE="./Includes/cnx.inc" -->
<%
 


'validaciones
usuario=trim(request.form("f1"))
clave=trim(request.form("f2"))
pin=trim(request.form("f3"))

if usuario="" then response.redirect("default.asp")
if clave="" then response.redirect("default.asp")
if pin="" then response.redirect("default.asp")







'BORRA DATOS DE USUARIO DE PRUEBA
cadena="delete from instrumentos where numeroafiliado='24485168'"
conexion.execute(cadena)


cadena="SELECT     userid, nivel, accesos, nombre, pin FROM         dbo.Usuarios "
cadena=cadena & " WHERE     (usuario = N'" & usuario & "') AND (password = N'" & clave & "') AND (pin = N'" & pin & "')"





set rs=conexion.execute(cadena)

if rs.eof=false then
	session("userid")=rs.fields(0)
	session("nivel")=rs.fields(1)
	session("accesos")=rs.fields(2) 
	session("usuario")=rs.fields(3)
	session("pin")=rs.fields(4)
end if

rs.close
set rs=nothing
conexion.close

set conexion=nothing

if session("userid")="" then
	response.redirect("default.asp")
else
	response.redirect("menu.asp")
end if

%>