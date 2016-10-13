<%

	Dim conexion
	Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
	conexion.connectionstring = conn
	conexion.open	
	
		cadena="update usuarios set activo=0 where userid=" & request.form("userid")
		
		conexion.execute(cadena)
		conexion.close
		set conexion=nothing
		response.redirect("mensaje.asp?mensaje=Usuario deshabilitado&destino=administracion")


%>
