<!-- #include FILE="./Includes/cnx.inc" -->
<%
f1=replace(trim(request.form("f1")),"'","''")
f2=replace(trim(request.form("f2")),"'","''")
f3=replace(trim(request.form("f3")),"'","''")
f4=replace(trim(request.form("f4")),"'","''")

opciones=request.form("checkbox")

ok=true

if f1="" then ok=false
if f2="" then ok=false
if f3="" then ok=false
if f4="" then ok=false
if opciones="" then ok=false


if ok=true then
	
	
	cadena="select usuario from usuarios where usuario='" & f2 & "'"
	
	set rs=conexion.execute(cadena)
	if rs.eof=false then
		ok=false
	end if
	rs.close
	set rs=nothing
	
	if ok=true then
		cadena="insert into usuarios (usuario, password, pin, nombre, accesos) values ("
		cadena=cadena & "'" & f2 & "', "
		cadena=cadena & "'" & f3 & "', "
		cadena=cadena & "'" & f4 & "', "
		cadena=cadena & "'" & f1 & "', "
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
		cadena=cadena & "'" & subcadena & "')"
		
		response.write(cadena)
		
		
		conexion.execute(cadena)
		conexion.close
		set conexion=nothing
		response.redirect("mensaje.asp?mensaje=Usuario dado de alta&destino=administracion")
		
		
	
	end if
end if

%>
<script>javascript:history.go(-1)</script>