<%



Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
'conn="Provider=SQLOLEDB;Data source=Data source=sql2106.shared-servers.com,1087;Initial catalog=stmedata;User Id=usuariostme;Password=Redalert7424;"
conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"

conexion.connectionstring = conn
conexion.open

'BORRA DATOS DE USUARIO DE PRUEBA
cadena="UPDATE INSTRUMENTOS SET TIPOINSTRUMENTO='ADELANTO' WHERE TIPOINSTRUMENTO LIKE '%ADELANTO%'"
conexion.execute(cadena)


%>