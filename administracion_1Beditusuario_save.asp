<%
f3=replace(trim(request.form("f3")),"'","''")
f4=replace(trim(request.form("f4")),"'","''")

opciones=request.form("checkbox")

ok=true

if f3="" then ok=false
if f4="" then ok=false
if opciones="" then ok=false


if ok=true then
	Dim conexion
	Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
	conexion.connectionstring = conn
	conexion.open	
	
		cadena="update usuarios set "
		cadena=cadena & " password='" & f3 & "', "
		cadena=cadena & " pin='" & f4 & "', "
		
		a=Split(opciones,",")
		for each x in a
			subcadena=subcadena & "*" & trim(x) & "*"
		next
		if instr(subcadena,"*2") > 0 then
			subcadena=subcadena & "*2*"
		end if
		if instr(subcadena,"*1") > 0 then
			subcadena=subcadena & "*1*"
		end if
		if instr(subcadena,"*3") > 0 then
			subcadena=subcadena & "*3*"
		end if

		cadena=cadena & "accesos='" & subcadena & "' where userid=" & request.form("userid")
		
		conexion.execute(cadena)
		conexion.close
		set conexion=nothing
		response.redirect("mensaje.asp?mensaje=Perfil de usuario modificado satisfactoriamente&destino=administracion")
		
end if

%>
<script>javascript:history.go(-1)</script>