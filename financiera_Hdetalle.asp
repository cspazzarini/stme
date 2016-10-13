<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if instr(session("accesos"), "*28*") =0 then response.redirect("sinacceso.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<!-- #include FILE="./Includes/cnx.inc" -->
<%



cadena="select legajo, importe, vencimiento from cierres_lote where mes_procesado=" & request("mes") & " and legajo <> '24485168'"

set rs=conexion.execute(cadena)

if rs.eof=false then
	do until rs.eof
		response.write("3;" & rs.fields(0) & ";841;" & rs.fields(1) & ";0;0;" & rs.fields(2) & "<br>")
		rs.movenext
	loop

end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing

%>

</body>
</html>
