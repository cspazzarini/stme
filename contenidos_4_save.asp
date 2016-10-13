<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*11*") =0 then response.redirect("sinacceso.asp")%>

<%
f1=replace(request.form("f1"),"'","''")
f2=replace(request.form("f2"),"'","''")
f3=replace(request.form("f3"),"'","''")
f4=replace(request.form("f4"),"'","''")
id=request("id")
			
			
			Dim conexion
			Set Conexion = CreateObject("ADODB.Connection")
			conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
			conexion.connectionstring = conn
			conexion.open	
			
			cadena="update comercios set comercio='" & f1 & "', direccion='" & f2 & "', telefono='" & f3 & "', email='" & f4 & "' where cid=" & id
			conexion.execute(cadena)
			
			conexion.close
			set conexion=nothing
			
			response.redirect("done.asp?destino=contenidos")
%>