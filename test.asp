<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<p>Fecha actual: <%=date%></p>
<p>Meses: 9</p>
<p>Vence: <%

vence=dateadd("m", 9, date)
MesVencimiento=year(vence) & right("00" & month(vence),2)

response.write(mesvencimiento)


%> </p>

</body>
</html>
