<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if instr(session("accesos"), "*28*") =0 then response.redirect("sinacceso.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Archivo de Salida - Ministerio de Salud</title>
</head>

<body>
<!-- #include FILE="./Includes/cnx.inc" -->
<%




cadena="select legajo,   replace ( cast ( cast (sum(importe) as decimal(18,2)) as varchar(10)) , '.', ''), vencimiento, tipoinstrumento , cierre_id = STUFF((SELECT N'|' + cast( cierre_id as varchar)   FROM dbo.CIERRES_LOTE_SALUD AS p2    WHERE p2.legajo = p.legajo and p2.tipoinstrumento = p.tipoinstrumento and p2.vencimiento = p.vencimiento    ORDER BY cierre_id    FOR XML PATH(N''), TYPE).value(N'.[1]', N'nvarchar(max)'), 1, 2, N''), nombre FROM dbo.CIERRES_LOTE_SALUD AS p inner join dbo.afiliados a on p.legajo = a.numeroafiliado where mes_procesado=" & request("mes") & " and legajo <> '24485168' group by legajo, tipoinstrumento,vencimiento,  nombre"

set rs=conexion.execute(cadena)

if rs.eof=false then
	do until rs.eof
		if (rs.fields(3) = "ALTA") then
			cpto = "281"
			importe = "000000000"
		else
			cpto = "285"
			importe = right ("000000000" & rs.fields(1),9)
		end if
		observacion = right ("0000000000" & rs.fields(4),10)
		nombre = right ("                    " & rs.fields(5),20)
		response.write("1" & "1" & "281" & right(request("mes"),2) & left(request("mes"),4) & right("0000000000" & rs.fields(0),10) & cpto & importe &   right( rs.fields(2),2) & left( rs.fields(2),4) & observacion &  nombre &";0;0;" & "<br>")
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
