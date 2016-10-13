<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
<!-- #include FILE="./Includes/cnx.inc" -->
<!--
.style1 {font-family: Arial, Helvetica, sans-serif}
.style2 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style3 {font-size: 12px}
.style5 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FFFFFF; }
-->
</style>
</head>

<body>

<%

		  
cadena="SELECT     TOP (100) PERCENT cid, Comercio, Direccion, Telefono, email, estado FROM         dbo.comercios ORDER BY Comercio"		

set rs=conexion.execute(cadena) 
total=0 
total2=0
		  %>
          </p>
<%if rs.eof=false then%>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#4F5D8A"><table width="1000" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td width="7%" height="25"><div align="center" class="style5"><span class="style44">ID</span></div></td>
        <td width="20%"><div align="center" class="style5"><span class="style44">Comercio</span></div></td>
        <td width="30%"><div align="center" class="style5"><span class="style44">Direcciòn</span></div></td>
        <td width="13%"><div align="center" class="style5"><span class="style44">Telefono</span></div></td>
        <td width="22%"><div align="center" class="style5"><span class="style44">Email</span></div></td>
        <td width="8%"><div align="center" class="style5"><span class="style44">Estado</span></div></td>
        </tr>
      <%do until rs.eof%>
      <tr>
        <td height="35" bgcolor="#FFFFFF"><div align="center"><span class="style47 style1 style3"><%=rs.fields(0)%></span></div></td>
        <td bgcolor="#FFFFFF"><span class="style47 style1 style3">&nbsp;<%=rs.fields(1)%>&nbsp;</span></td>
        <td bgcolor="#FFFFFF"><span class="style47 style1 style3">&nbsp;<%=rs.fields(2)%>&nbsp;</span></td>
        <td bgcolor="#FFFFFF"><div align="right" class="style2"><span class="style47"><%=rs(3)%>&nbsp;</span></div></td>
        <td bgcolor="#FFFFFF"><div align="right" class="style2"><span class="style47"><%=rs(4)%>&nbsp;</span></div></td>
        <td bgcolor="#FFFFFF"><div align="center" class="style2"><span class="style47"><%if rs(5)=1 then response.write("ACTIVO")%></span></div></td>
        </tr>
      <%
	
		   rs.movenext
		   loop%>
    </table></td>
  </tr>
</table>
      <%end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing%>
</p>
<p>



</body>
</html>
