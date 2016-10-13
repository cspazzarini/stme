<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*11*") =0 then response.redirect("sinacceso.asp")%>

<%
f1=replace(request.form("f1"),"'","''")
f2=replace(request.form("f1"),"'","''")
f3=replace(request.form("f1"),"'","''")
f4=replace(request.form("f1"),"'","''")

			
			
			Dim conexion
			Set Conexion = CreateObject("ADODB.Connection")
			conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
			conexion.connectionstring = conn
			conexion.open	
			

			cadena="insert into comercios (comercio, direccion, telefono, email, estado) values ("
			cadena=cadena & "'" & f1 & "', "
			cadena=cadena & "'" & f2 & "', "
			cadena=cadena & "'" & f3 & "', "
			cadena=cadena & "'" & f4 & "', 1) "
			
			conexion.execute(cadena)
			
			conexion.close
			set conexion=nothing
			
			response.redirect("done.asp?destino=contenidos")
%>