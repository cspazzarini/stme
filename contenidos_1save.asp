<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*11*") =0 then response.redirect("sinacceso.asp")%>


<!-- #include FILE="./Includes/cnx.inc" -->

<%
ierror=0

f1=trim(request.form("f1"))
if f1="" then ierror=1



f40=trim(request.form("f40"))
if f40="" then ierror=1


f2=trim(request.form("f2"))
if f2="" then ierror=1



f8=trim(request.form("f8"))
if f8="" then ierror=1


f11_2=trim(request.form("f11_2"))
if f11_2="" then ierror=1



f16=trim(request.form("f16"))
if f16="" then ierror=1



f18=trim(request.form("f18"))
if f18="" then ierror=1


facn1=trim(request.form("facn1"))




if ierror=0 then
			'terminaron validaciones
			'comienzo grabacion de datos
			
			numeroafiliado=request.form("f40")
			nombre=request.form("f1")
			documento=request.form("f2")
			fechanacimiento=request.form("f3_3") & right("00" & request.form("f3_1"),2) & right("00" & request.form("f3_2"),2)
			estadocivil=request.form("f4")
			tipoafiliado=request.form("f5")
			cantidadhijos=request.form("f6")
			email=request.form("f7")
			domicilio=request.form("f8")
			telefono=request.form("f9")
			fechaingreso=request.form("f10_3") & right("00" & request.form("f10_1"),2) & right("00" & request.form("f10_2"),2)
			situacionanses=request.form("f11_1")
			beneficiarioanses=request.form("f11_2")
			mutual=request.form("f12")
			grupofamiliar=request.form("f13")
			agrupamiento=request.form("f14")
			categoria=request.form("f15")
			lugartrabajo=request.form("f16")
			situacionlaboral=request.form("f17")
			cuentabanco=request.form("f18")
			userid=session("userid")
			
			
			
			
			
			cadena="select nombre from afiliados where numeroafiliado='" & numeroafiliado & "'"
			set rs=conexion.execute(cadena)
			if rs.eof=false then
				xn=rs.fields(0)
				rs.close
				set rs=nothing
				conexion.close
				set conexion=nothing
				%>
				<script>javascript:history.go(-1);alert("Numero de afiliado <%=numeroafiliado%> ya existe en base de datos y pertenece a <%=ucase(xn)%>")</script>
				<%
			end if
			rs.close
			set rs=nothing
			
			
			cadena="select nombre from afiliados where documento='" & documento & "'"
			set rs=conexion.execute(cadena)
			if rs.eof=false then
				xn=rs.fields(0)
				rs.close
				set rs=nothing
				conexion.close
				set conexion=nothing
				%>
				<script>javascript:history.go(-1);alert("El DNI <%=documento%> ya existe en base de datos y pertenece a  <%=ucase(xn)%>")</script>
				<%
			end if
			rs.close
			set rs=nothing
			
			
			cadena="insert into afiliados (numeroafiliado, nombre, documento, fechanacimiento, estadocivil, tipoafiliado, cantidadhijos, email, domicilio, telefono, fechaingreso, "
			CADENA=CADENA & " situacionanses, mutual, beneficiario, grupofamiliar, agrupamiento, categoria, lugartrabajo, situacionlaboral, cuentabco, userid) values ("
			cadena=cadena & "'" & numeroafiliado & "', "
			cadena=cadena & "'" & Ucase(nombre) & "', "
			cadena=cadena & "'" & documento & "', "
			cadena=cadena & fechanacimiento & ", "
			cadena=cadena & "'" & estadocivil & "', "
			cadena=cadena & tipoafiliado & ", "
			cadena=cadena & cantidadhijos & ", "
			cadena=cadena & "'" & email & "', "
			cadena=cadena & "'" & domicilio & "', "
			cadena=cadena & "'" & telefono & "', "
			cadena=cadena & fechaingreso & ", "
			cadena=cadena & "'" & situacionanses & "', "
			cadena=cadena & "'" & mutual & "', "
			cadena=cadena & "'" & beneficiarioanses & "', "
			cadena=cadena & "'" & grupofamiliar & "', "
			cadena=cadena & agrupamiento & ", "
			cadena=cadena & "'" & categoria & "', "
			cadena=cadena & "'" & lugartrabajo & "', "
			cadena=cadena & "'" & situacionlaboral & "', "
			cadena=cadena & "'" & cuentabanco & "', "
			cadena=cadena & userid & ")"
			
			conexion.execute(cadena)
			
			
			set rs=conexion.execute("select max(id) as maximo from afiliados where numeroafiliado='" & numeroafiliado & "' and documento='" & documento & "'")
			if rs.eof=false then
				id=rs.fields(0)
			end if
			rs.close
			set rS=nothing
			

			legajo=0
			for t=1 to 15
				nombre=trim(request.form("facn" & t))
				parentesco=trim(request.form("facp" & t))
				sexo=trim(request.form("facs" & t))
				fechanacimiento=request.form("facano" & t) & right("00" & request.form("facmes" & t),2) & right("00" & request.form("facdia" & t),2)
				userid=session("userid")
				if nombre <> "" then
					legajo=legajo+1
					cadena="insert into afiliados_parientes (id, numeroafiliado, sublegajo, nombre, parentesco, sexo, fechanacimiento, userid) values ("
					cadena=cadena & id & ", "
					cadena=cadena & numeroafiliado & ", "
					cadena=cadena & legajo & ", "
					cadena=cadena & "'" & nombre & "', "
					cadena=cadena & "'" & parentesco & "', "
					cadena=cadena & "'" & sexo & "', "
					cadena=cadena & fechanacimiento & ", "
					cadena=cadena & userid & ")"
					conexion.execute(cadena)
				end if
			next
			conexion.close
			set conexion=nothing
			
			response.redirect("done.asp?destino=contenidos")
else
%>
<script>javascript:history.go(-1);alert("Ingrese todos los campos requeridos")</script>
<%
end if
%>