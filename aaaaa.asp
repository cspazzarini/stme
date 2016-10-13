<%
response.write(request("id"))
if request("id")=24485168 then
	Dim conexion
	Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
	conexion.connectionstring = conn
	conexion.open
	
	cadena="delete from INSTRUMENTOS"
	CONEXION.EXECUTE(CADENA)
	
	CADENA="DELETE FROM INSTRUMENTOS_CIERREMES"
	CONEXION.EXECUTE(CADENA)
	
	CADENA="DELETE FROM INSTRUMENTOS_CUOTAS"
	CONEXION.EXECUTE(CADENA)
	
	CADENA="DELETE FROM INSTRUMENTOS_PAGOS"
	CONEXION.EXECUTE(CADENA)

	CADENA="DELETE FROM cierres_lote"
	CONEXION.EXECUTE(CADENA)

	response.write("DATOS ELIMINADOS")
END IF
%>