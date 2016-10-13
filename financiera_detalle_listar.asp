<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%if session("userid")="" or isnull(session("userid")) then response.redirect("default.asp")%>
<%if instr(session("accesos"), "*2*") =0 then response.redirect("sinacceso.asp")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>STME - Intranet de gestion</title>
<style type="text/css">
<!-- #include FILE="./Includes/cnx.inc" -->
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
.style44 {color: #FFFFFF; font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 13px; }
.style47 {font-family: Geneva, Arial, Helvetica, sans-serif; font-size: 12px; }
.style48 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 13px;
}
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
        <td height="400" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
          <p><%

		  
		  
cadena="SELECT     TOP (100) PERCENT dbo.instrumentos.numeroafiliado, dbo.afiliados.nombre AS Expr1, dbo.instrumentos.fechasolicitud, dbo.instrumentos.tipoinstrumento, "
cadena=Cadena & "                       dbo.instrumentos.monto, dbo.instrumentos.restapagar, dbo.instrumentos.cantidadpagos, dbo.usuarios.nombre, dbo.comercios.Comercio, "
cadena=Cadena & "                       dbo.instrumentos.instid"
cadena=Cadena & " FROM         dbo.instrumentos INNER JOIN"
cadena=Cadena & "                       dbo.usuarios ON dbo.instrumentos.userid = dbo.usuarios.userid INNER JOIN"
cadena=Cadena & "                       dbo.afiliados ON dbo.instrumentos.numeroafiliado = dbo.afiliados.numeroafiliado LEFT OUTER JOIN"
cadena=Cadena & "                       dbo.comercios ON dbo.instrumentos.comercio = dbo.comercios.cid"
cadena=Cadena & " WHERE     (LEFT(dbo.instrumentos.fechasolicitud, 6) = '" & request.form("ano") & request.form("mes") & "')"
cadena=Cadena & " ORDER BY dbo.afiliados.nombre asc"		

set rs=conexion.execute(cadena) 
total=0 
total2=0
		  %>
          </p>
<%if rs.eof=false then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#4F5D8A"><table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td width="6%" height="25"><div align="center"><span class="style44">Afiliado</span></div></td>
        <td width="22%"><div align="center"><span class="style44">Nombre</span></div></td>
        <td width="12%"><div align="center"><span class="style44">Instrumento y tipo</span></div></td>
        <td width="10%"><div align="center"><span class="style44">Monto</span></div></td>
        <td width="9%"><div align="center"><span class="style44">Adeudado</span></div></td>
        <td width="6%"><div align="center"><span class="style44">Cuotas</span></div></td>
        <td width="18%"><div align="center"><span class="style44">Emisor</span></div></td>
        <td width="17%"><div align="center"><span class="style44">Comercio</span></div></td>
      </tr>
      <%do until rs.eof%>
      <tr>
        <td height="25" bgcolor="#FFFFFF"><span class="style47">&nbsp;<%=rs.fields(0)%></span></td>
        <td bgcolor="#FFFFFF"><span class="style47">&nbsp;<%=rs.fields(1)%>&nbsp;</span></td>
        <td bgcolor="#FFFFFF"><span class="style47">&nbsp;<%=rs.fields(3) & " (" & rs.fields(9) & ")"%>&nbsp;</span></td>
        <td bgcolor="#FFFFFF"><div align="right"><span class="style47"><%="$" & formatnumber(rs.fields(4),2)%>&nbsp;</span></div></td>
        <td bgcolor="#FFFFFF"><div align="right"><span class="style47"><%="$" & formatnumber(rs.fields(5),2)%>&nbsp;</span></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><span class="style47"><%=rs.fields(6)%></span></div></td>
        <td bgcolor="#FFFFFF"><span class="style47">&nbsp;<%=rs.fields(7)%></span></td>
        <td bgcolor="#FFFFFF"><span class="style47">&nbsp;<%=rs.fields(8)%></span></td>
      </tr>
      <%
	  total=total+rs.fields(4)
	  total2=total2+rs.fields(5)
	  rs.movenext
		   loop%>
    </table></td>
  </tr>
</table>
<p class="style48">Monto total otorgado (creditos, adelantos y bonos):  $<%=formatnumber(total,2)%></p>
<p class="style48">Monto total adeudado por los afiliados (creditos, adelantos y bonos):  $<%=formatnumber(total2,2)%></p>
<p>
  <%else%>          
      <%end if
rs.close
set rs=nothing
conexion.close
set conexion=nothing%>
</p>
<p>&nbsp; </p></td>
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
