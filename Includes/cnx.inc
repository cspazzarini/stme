<%
	Dim conexion
	Set Conexion = CreateObject("ADODB.Connection")
	conn="Provider=SQLOLEDB;Data source=NOTEBOOK\SQL2012;Database=stmedata;UID=sa;PWD=sql2014*"
	conexion.connectionstring = conn
	conexion.open		
%>