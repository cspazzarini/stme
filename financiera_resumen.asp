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
.style15 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #344453;
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
.style35 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style36 {font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	font-size: 12px;
}
.style44 {font-family: Arial, Helvetica, sans-serif}
.style45 {font-family: Arial, Helvetica, sans-serif; font-size: 12px; }
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
          <p>&nbsp;</p>
          <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="30" bgcolor="#384858"><div align="center"><span class="style35">RESUMEN DE EMISION DE DOCUMENTOS</span></div></td>
            </tr>
            <tr>
              <td bgcolor="#374957"><table width="800" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="16%" height="30" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">Mes</span>/Año</div></td>
                    <td width="16%" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">Bonos</span></div></td>
                    <td width="16%" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">Créditos</span></div></td>
                    <td width="16%" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">Adelantos</span> Efectivo</div></td>
                    <td width="16%" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">Adelantos</span> Cheques</div></td>
                    <td width="16%" bgcolor="#e5e5e5" class="style1"><div align="center"><span class="style44">TOTAL</span></div></td>
                  </tr>
                  <%
			
	CADENA="SELECT     TOP (100) PERCENT LEFT(fechasolicitud, 6) AS fecha, SUM(monto) AS Expr1 FROM         dbo.instrumentos where numeroafiliado <> '1'"
	cadena=cadena & " GROUP BY LEFT(fechasolicitud, 6) ORDER BY fecha DESC"
	SET RS=CONEXION.EXECUTE(CADENA)
	IF RS.EOF=FALSE THEN
		do until rs.eof
				  %>
                  <tr>
                    <td height="30" bgcolor="#F9f9f9"><div align="center" class="style1"><%=mid(rs(0),5,2) & "/" & left(rs(0),4)%></div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style45"><%
					
					CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1"
					cadena=cadena & " FROM         dbo.instrumentos"
					cadena=cadena & " WHERE     (LEFT(fechasolicitud, 6) = '" & rs(0) & "') AND (numeroafiliado <> N'1' AND numeroafiliado <> N'24485168')"
					cadena=cadena & " GROUP BY tipoinstrumento"
					cadena=cadena & " HAVING      (tipoinstrumento = N'BONO')"

					SET RS2=CONEXION.EXECUTE(CADENA)
					IF RS2.EOF=FALSE THEN
						RESPONSE.WRITE("$" & FORMATNUMBER(RS2(0),2))
						END IF
						RS2.CLOSE
						

					%>&nbsp;&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style45"><%
					
					CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1"
					cadena=cadena & " FROM         dbo.instrumentos"
					cadena=cadena & " WHERE     (LEFT(fechasolicitud, 6) = '" & rs(0) & "') AND (numeroafiliado <> N'1' AND numeroafiliado <> N'24485168')"
					cadena=cadena & " GROUP BY tipoinstrumento"
					cadena=cadena & " HAVING      (tipoinstrumento = N'CREDITO')"

					SET RS2=CONEXION.EXECUTE(CADENA)
					IF RS2.EOF=FALSE THEN
						RESPONSE.WRITE("$" & FORMATNUMBER(RS2(0),2))
						END IF
						RS2.CLOSE
						

					%>&nbsp;&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style45"><%
					
					CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1"
					cadena=cadena & " FROM         dbo.instrumentos"
 					cadena=cadena & " WHERE     (LEFT(fechasolicitud, 6) = '" & rs(0) & "') AND (numeroafiliado <> N'1' AND numeroafiliado <> N'24485168')"
					cadena=cadena & " and (tipoadelanto is null or tipoadelanto='Efectivo') "
					cadena=cadena & " GROUP BY tipoinstrumento"
					cadena=cadena & " HAVING      (tipoinstrumento = N'ADELANTO')"
					SET RS2=CONEXION.EXECUTE(CADENA)
					IF RS2.EOF=FALSE THEN
						RESPONSE.WRITE("$" & FORMATNUMBER(RS2(0),2))
						END IF
						RS2.CLOSE
						

					%>&nbsp;&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style45">
                      <%
					
					CADENA="SELECT     TOP (100) PERCENT SUM(monto) AS Expr1"
					cadena=cadena & " FROM         dbo.instrumentos"
 					cadena=cadena & " WHERE     (LEFT(fechasolicitud, 6) = '" & rs(0) & "') AND (numeroafiliado <> N'1' AND numeroafiliado <> N'24485168')"
					cadena=cadena & " and (tipoadelanto='Cheque') "
					cadena=cadena & " GROUP BY tipoinstrumento"
					cadena=cadena & " HAVING      (tipoinstrumento = N'ADELANTO')"
					SET RS2=CONEXION.EXECUTE(CADENA)
					IF RS2.EOF=FALSE THEN
						RESPONSE.WRITE("$" & FORMATNUMBER(RS2(0),2))
						END IF
						RS2.CLOSE
						

					%>
                      &nbsp;&nbsp;&nbsp;</div></td>
                    <td bgcolor="#FFFFFF"><div align="right" class="style45">
                      <%
					
											RESPONSE.WRITE("$" & FORMATNUMBER(RS(1),2))
						

					%>
                      &nbsp;&nbsp;&nbsp;</div></td>
                  </tr>
	<%rs.movenext
	
	loop
    end if
	rs.close
	conexion.close
	
	%>
              </table></td>
            </tr>
          </table>          <p align="center" class="style15">&nbsp;</p>
          </td>
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
