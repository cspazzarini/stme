<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 13px;
	font-weight: bold;
	color: #344356;
}
a:link {
	text-decoration: none;
	color: #00769D;
}
a:visited {
	text-decoration: none;
	color: #00769D;
}
a:hover {
	text-decoration: none;
	color: #00769D;
}
a:active {
	text-decoration: none;
	color: #00769D;
}
.style32 {color: #FF0000}
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
.style44 {font-family: Arial, Helvetica, sans-serif}
.style46 {font-size: 12px}
.style47 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
.style49 {color: #FFFFFF}
.style51 {font-family: Arial, Helvetica, sans-serif; color: #344356;}
.style52 {font-size: 13px; color: #344356; font-family: Arial, Helvetica, sans-serif;}
.style57 {font-size: 13px}
-->
</style>
<script src="../../../../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body>
<table width="1012" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="6" background="images/izquierda.jpg">&nbsp;</td>
    <td width="1000"><table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td height="8" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="110"><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr bgcolor="#FFFFFF">
            <td height="110" background="images/logonu_fondo.jpg"><img src="images/logonu.jpg" width="276" height="199" /></td>
            <td background="images/logonu_fondo.jpg"><p align="right" class="style36"><br />
              INTRANET DE GESTION DEL SINDICATO DE TRABAJADORES MUNICIPALES DE ENSENADA <br />
            </p></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="33" background="images/fondo_menu_top.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="153"><div align="center"><span class="style1"><a href="menu.asp">INICIO</a></span></div></td>
            <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
            <td width="170"><div align="center"><span class="style1"><a href="contenidos.asp">GESTION  DE CONTENIDOS</a></span></div></td>
            <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
            <td width="171"><div align="center"><span class="style1"><a href="financiera.asp"><span class="style32">GESTION FINANCIERA </span></a></span></div></td>
            <td width="10"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
            <td width="154"><div align="center"><span class="style1"><a href="coseguro.asp">GESTION COSEGURO </a></span></div></td>
            <td width="9"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
            <td width="155"><div align="center"><span class="style1"><a href="administracion.asp">ADMINISTRACION</a></span></div></td>
            <td width="10"><div align="center"><img src="images/separador_menu_top.jpg" width="5" height="33" /></div></td>
            <td width="152"><div align="center"><span class="style1"><a href="logout.asp">CERRAR SESION</a></span></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="2" bgcolor="#2F3863"></td>
      </tr>
      <tr>
        <td height="400" valign="top" bgcolor="#FFFFFF"><%
Dim conexion
Set Conexion = CreateObject("ADODB.Connection")
conn="Provider=SQLOLEDB;Data source=ARTONEST;Database=stmedata;UID=usuariostme;PWD=Redalert7424"
conexion.connectionstring = conn
conexion.open
		  
		  
cadena="SELECT     TOP (100) PERCENT cid, Comercio, Direccion, Telefono, email, estado FROM         dbo.comercios ORDER BY Comercio"		

set rs=conexion.execute(cadena) 
total=0 
total2=0
		  %>
          </p>
          <%if rs.eof=false then%>
          <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#4F5D8A"><table width="900" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="7%" height="25"><div align="center" class="style49 style46 style44 style5"><strong>ID</strong></div></td>
                    <td width="20%"><div align="center" class="style49 style46 style44 style5"><strong>Comercio</strong></div></td>
                    <td width="30%"><div align="center" class="style49 style46 style44 style5"><strong>Direcciòn</strong></div></td>
                    <td width="13%"><div align="center" class="style49 style46 style44 style5"><strong>Telefono</strong></div></td>
                    <td width="22%"><div align="center" class="style49 style46 style44 style5"><strong>Email</strong></div></td>
                    <td width="8%"><div align="center" class="style49 style46 style44 style5"><strong>Editar</strong></div></td>
                  </tr>
                  <%do until rs.eof%>
                  <tr>
                    <td height="35" bgcolor="#FFFFFF"><div align="center" class="style47"><span class="style3  style51"><%=rs.fields(0)%></span></div></td>
                    <td bgcolor="#FFFFFF"><span class="style52 style47 style3 style44 style46"><span class="style44 style3 style57">&nbsp;<%=rs.fields(1)%>&nbsp;</span></span></td>
                    <td bgcolor="#FFFFFF"><span class="style52 style47 style3 style44 style46"><span class="style44 style3 style57">&nbsp;<%=rs.fields(2)%>&nbsp;</span></span></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style2 style44 style46"><span class="style44"><%=rs(3)%>&nbsp;</span></div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style2 style44 style46"><span class="style44"><%=rs(4)%>&nbsp;</span></div></td>
                    <td bgcolor="#FFFFFF"><div align="center" class="style2 style44 style46"><a href="contenidos_4_editar.asp?id=<%=rs(0)%>">Editar</a></div></td>
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
          <p>&nbsp;</p></td>
      </tr>
      <tr>
        <td height="2" bgcolor="#2F3863"></td>
      </tr>
    </table></td>
    <td width="6" background="images/derecha.jpg">&nbsp;</td>
  </tr>
</table>
</body>
</html>
